/**
 * Feature 07-A — Table of Contents: service layer (Office JS operations).
 *
 * Implements TOC field insert, refresh, reconciliation, and scoped-TOC
 * bookmark creation. All Office JS calls are isolated here so the driver
 * and UI layers remain testable without a live Word host.
 *
 * ## Office JS fidelity notes (ARCHITECTURE.md §Office JS Constraints Reference)
 *
 * Row 5: `Range.insertField(InsertLocation.before, FieldType.toc)` is WordApi 1.5.
 *   - The returned `Word.Field` proxy is live only within the same Word.run context.
 *   - `field.code` must be set BEFORE the first sync after insert; the write is
 *     batched and applied at the subsequent ctx.sync().
 *   - `field.updateResult()` is called AFTER setting code, then another ctx.sync()
 *     triggers the actual TOC build. This is two syncs total per insert.
 *   - `field.id` (WordApi 1.5) is available after load("id"); we read it to
 *     persist into the registry for "refresh this TOC" and "go to" operations.
 *
 * ## F9 fallback (wordApi15 absent)
 *
 * When WordApi 1.5 is unavailable, `insertField` does not exist. We fall back
 * to `range.insertOoxml(w:fldSimple)` which inserts a static-OOXML field stub
 * that Word renders when the document is next calculated (F9 / right-click
 * "Update Field"). `updateResult()` is NOT called in the fallback path because
 * the API does not exist; the UI banner instructs the user to press F9.
 *
 * The w:fldSimple OOXML form is:
 *   <w:fldSimple w:instr=" TOC <switches> "><w:r><w:t></w:t></w:r></w:fldSimple>
 * (xmlns is inherited from the surrounding document body.)
 *
 * ## CONCERN
 *
 * The `field.id` property — used to persist the registry entry — is a WordApi 1.5
 * property, but its exact availability and uniqueness guarantees across hosts
 * (Word Online vs Win/Mac) need live-host verification. The unit tests mock it.
 * If `field.id` proves unavailable on a host, the registry entry will be stored
 * with an empty id and the reconciliation step (A3) will prune it safely.
 *
 * Additionally, the exact behaviour of `range.insertOoxml(w:fldSimple)` across
 * Word Online / Word for Mac / Word for Windows needs live-host verification.
 * The unit tests cover the branching logic; the OOXML rendering is host-only.
 *
 * Spec reference: spec/features/07-toc-precedent-templates.md §A.3 / §A.5.
 */

import { isCapable, wordRun } from '@shared/office/index.js';
import { getStorageProvider } from '@shared/storage/index.js';
import type { TocProfile, TocRegistry, TocRegistryEntry, BuildSwitchesOptions } from './types.js';
import { buildTocSwitches, buildScheduleIndexSwitches } from './switches.js';
import type { ScheduleIndexConfig } from './types.js';

// ---------------------------------------------------------------------------
// Storage keys
// ---------------------------------------------------------------------------

const DOC_SETTINGS_KEY_REGISTRY = 'toc.registry';

// ---------------------------------------------------------------------------
// w:fldSimple OOXML template (F9 fallback)
// ---------------------------------------------------------------------------

/**
 * Build a w:fldSimple OOXML fragment for the F9 fallback path.
 * The instruction string is inserted as-is; Word renders it on next update.
 *
 * Note: Word expects the namespace to be inherited; the fragment is deliberately
 * minimal — only the required attributes.
 */
export function buildFldSimpleOoxml(instruction: string): string {
  // Full XML-attribute escape. A firm style name in the switch string (e.g.
  // `\t "Heading & Title,1"`) can contain &, <, > or " — escaping only " would
  // emit malformed XML and insertOoxml would throw on the fallback path.
  // Order matters: & must be escaped first.
  const escaped = instruction
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
  return `<w:fldSimple w:instr="${escaped}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t></w:t></w:r></w:fldSimple>`;
}

// ---------------------------------------------------------------------------
// Insert a TOC field (A.5.1)
// ---------------------------------------------------------------------------

export interface InsertTocOptions {
  /** Where to place the field: at cursor ("atCursor") or top of document ("topOfDocument"). */
  placement: 'atCursor' | 'topOfDocument';
  /** Label for the registry entry. */
  scopeLabel: string;
  /** When provided, appends \b <bookmark> to the switch string (scoped TOC). */
  bookmark?: string;
}

export interface InsertTocResult {
  /** Whether the insert used the native WordApi 1.5 path. */
  usedNativeField: boolean;
  /** Whether the F9 fallback path was used (wordApi15 absent). */
  usedFallback: boolean;
  /** The registry entry that was persisted. May have empty fieldId on the fallback path. */
  registryEntry: TocRegistryEntry;
  /**
   * When usedNativeField is true and `field.result.text` was empty after
   * updateResult(), the TOC had no entries — callers may show a hint to apply
   * heading styles.
   */
  isEmpty: boolean;
}

/**
 * Insert a TOC field into the document.
 *
 * WordApi 1.5 path: insertField → set code → updateResult → read field.id.
 * F9 fallback path: insertOoxml(w:fldSimple) + banner instruction.
 */
export async function insertToc(
  profile: TocProfile,
  opts: InsertTocOptions,
): Promise<InsertTocResult> {
  const switches = buildTocSwitches(profile, { bookmark: opts.bookmark } as BuildSwitchesOptions);
  const fieldCode = 'TOC' + switches;

  if (isCapable('wordApi15')) {
    // --- Native path (WordApi 1.5) ---
    let fieldId = '';
    let isEmpty = false;

    await wordRun(async ctx => {
      // Step 1: resolve the insertion range.
      const range = opts.placement === 'topOfDocument'
        ? ctx.document.body.getRange('Start')
        : ctx.document.getSelection().getRange('Start');

      // Step 2: insertField — returns a live field proxy.
      // InsertLocation.before places the field just before the resolved range.
      const field = range.insertField(
        Word.InsertLocation.before,
        Word.FieldType.toc,
      );

      // Step 3: set field.code (batched — not yet synced).
      field.code = fieldCode;

      // Step 4: first sync — writes the code.
      await ctx.sync();

      // Step 5: update the TOC (renders entries).
      field.updateResult();

      // Step 6: load result text for empty-check.
      // Note: field.id is not in the @types/office-js type definitions but may
      // be present at runtime on some hosts (WordApi 1.5+). We access it via a
      // type cast as a best-effort identifier. If absent, fieldId stays empty
      // and the reconciliation step safely keeps the entry.
      field.load('result');

      // Step 7: second sync — triggers TOC build and loads result.
      await ctx.sync();

      // Access id via runtime property (not in typed API).
      fieldId = (field as unknown as { id?: string }).id ?? '';

      // Empty check: result.text is empty when no headings map to the TOC.
      try {
        isEmpty = (field.result?.text?.trim() ?? '') === '';
      } catch {
        isEmpty = false;
      }
    });

    const entry: TocRegistryEntry = {
      fieldId,
      scopeLabel: opts.scopeLabel,
      ...(opts.bookmark ? { bookmark: opts.bookmark } : {}),
    };

    await _appendToRegistry(entry);

    return { usedNativeField: true, usedFallback: false, registryEntry: entry, isEmpty };

  } else {
    // --- F9 fallback path (wordApi15 absent) ---
    const ooxml = buildFldSimpleOoxml(' ' + fieldCode + ' ');

    await wordRun(async ctx => {
      const range = opts.placement === 'topOfDocument'
        ? ctx.document.body.getRange('Start')
        : ctx.document.getSelection().getRange('Start');

      // insertOoxml inserts raw OOXML at the range; Word renders it on next update.
      range.insertOoxml(ooxml, Word.InsertLocation.before);

      await ctx.sync();
    });

    // The fallback path cannot reliably retrieve the field id because we have
    // no handle to the field object.
    const entry: TocRegistryEntry = {
      fieldId: '',
      scopeLabel: opts.scopeLabel,
    };

    await _appendToRegistry(entry);

    return { usedNativeField: false, usedFallback: true, registryEntry: entry, isEmpty: false };
  }
}

// ---------------------------------------------------------------------------
// Refresh this TOC (A.5.3)
// ---------------------------------------------------------------------------

/**
 * Refresh a single TOC identified by its registry entry's fieldId.
 * No-op if the fieldId is empty (fallback-path entry).
 */
export async function refreshOneToc(fieldId: string): Promise<void> {
  if (!fieldId || !isCapable('wordApi15')) return;

  await wordRun(async ctx => {
    // Locate the field by type, then find the one matching our id.
    const fields = ctx.document.fields.getByTypes([Word.FieldType.toc]);
    // Load id via cast — id is not in the @types/office-js definition
    // but is present at runtime on WordApi 1.5+ hosts.
    (fields as unknown as { load(p: string): void }).load('id');
    await ctx.sync();

    const target = fields.items.find(f =>
      (f as unknown as { id?: string }).id === fieldId,
    );
    if (target) {
      target.updateResult();
      await ctx.sync();
    }
  });
}

// ---------------------------------------------------------------------------
// Refresh all TOCs + reconcile (A.5.3)
// ---------------------------------------------------------------------------

/**
 * Refresh all TOC fields in the document (batched single sync) and reconcile
 * the registry by pruning entries whose field no longer exists.
 */
export async function refreshAllTocs(): Promise<void> {
  if (!isCapable('wordApi15')) return;

  const registry = await _loadRegistry();

  await wordRun(async ctx => {
    const fields = ctx.document.fields.getByTypes([Word.FieldType.toc]);
    // Load id via cast — same runtime-only field.id approach as refreshOneToc.
    (fields as unknown as { load(p: string): void }).load('id');
    await ctx.sync();

    // Batch updateResult — all in one sync, no per-item sync.
    for (const field of fields.items) {
      field.updateResult();
    }
    await ctx.sync();

    // Reconcile: determine which registry fieldIds still exist.
    const liveIds = new Set(
      fields.items.map(f => (f as unknown as { id?: string }).id ?? ''),
    );

    // Prune entries whose fieldId is no longer live.
    // Entries with empty fieldId (F9 fallback) are kept — they cannot be verified.
    const reconciled = registry.filter(
      entry => !entry.fieldId || liveIds.has(entry.fieldId),
    );

    if (reconciled.length !== registry.length) {
      await _saveRegistry(reconciled);
    }
  });
}

// ---------------------------------------------------------------------------
// Scoped TOC bookmark generation (A.5.2 / A4)
// ---------------------------------------------------------------------------

/**
 * Generate a stable, collision-free bookmark name for a scoped TOC.
 *
 * Format: TOC_SCOPE_<slug>_<n>
 *   slug — sanitised label (lower-case, non-alphanum → underscore, max 20 chars)
 *   n    — 1-based counter, incremented until no collision with existing bookmarks.
 *
 * Collision is checked against `document.bookmarks` in the caller's Word.run.
 * This function is pure (no Office JS) — the caller passes in the existing names.
 */
export function generateBookmarkName(
  label: string,
  existingNames: string[],
): string {
  const slug = label
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 20);

  const existing = new Set(existingNames);
  let counter = 1;
  let name = `TOC_SCOPE_${slug}_${counter}`;
  while (existing.has(name)) {
    counter++;
    name = `TOC_SCOPE_${slug}_${counter}`;
  }
  return name;
}

// ---------------------------------------------------------------------------
// Insert an Index of Schedules TOC (Strategy 1, A.5.4)
// ---------------------------------------------------------------------------

export interface InsertScheduleIndexResult {
  usedNativeField: boolean;
  usedFallback: boolean;
  isEmpty: boolean;
}

/**
 * Insert an Index of Schedules as a separate TOC field (Strategy 1).
 *
 * The field uses ONLY \t "<scheduleStyles>" with NO \o, ensuring body-heading
 * styles are excluded. Strategy 2 (HTML snapshot) is handled in schedule-index.ts.
 */
export async function insertScheduleIndexField(
  config: ScheduleIndexConfig,
  placement: 'atCursor' | 'topOfDocument',
): Promise<InsertScheduleIndexResult> {
  const switches = buildScheduleIndexSwitches(
    config.scheduleStyles,
    config.level,
    true, // always hyperlinked
  );
  const fieldCode = 'TOC' + switches;

  if (isCapable('wordApi15')) {
    let isEmpty = false;

    await wordRun(async ctx => {
      const range = placement === 'topOfDocument'
        ? ctx.document.body.getRange('Start')
        : ctx.document.getSelection().getRange('Start');

      const field = range.insertField(Word.InsertLocation.before, Word.FieldType.toc);
      field.code = fieldCode;
      await ctx.sync();

      field.updateResult();
      field.load('result');
      await ctx.sync();

      try {
        isEmpty = (field.result?.text?.trim() ?? '') === '';
      } catch {
        isEmpty = false;
      }
    });

    return { usedNativeField: true, usedFallback: false, isEmpty };

  } else {
    const ooxml = buildFldSimpleOoxml(' ' + fieldCode + ' ');

    await wordRun(async ctx => {
      const range = placement === 'topOfDocument'
        ? ctx.document.body.getRange('Start')
        : ctx.document.getSelection().getRange('Start');

      range.insertOoxml(ooxml, Word.InsertLocation.before);
      await ctx.sync();
    });

    return { usedNativeField: false, usedFallback: true, isEmpty: false };
  }
}

// ---------------------------------------------------------------------------
// Scoped TOC with bookmark (A.5.2 — calls A4 path)
// ---------------------------------------------------------------------------

/**
 * Insert a scoped TOC for the current selection.
 * Reads existing bookmarks to avoid collisions, inserts a new bookmark around
 * the selection, then inserts a TOC field with \b <bookmarkName>.
 */
export async function insertScopedToc(
  profile: TocProfile,
  scopeLabel: string,
): Promise<InsertTocResult> {
  let bookmarkName = '';
  const switches_opts: BuildSwitchesOptions = {};

  if (isCapable('wordApi15')) {
    let existingNames: string[] = [];

    // Step 1: read existing bookmark names inside one Word.run for collision check.
    await wordRun(async ctx => {
      const bookmarks = ctx.document.bookmarks;
      bookmarks.load('name');
      await ctx.sync();
      existingNames = bookmarks.items.map(b => b.name);
    });

    bookmarkName = generateBookmarkName(scopeLabel, existingNames);
    switches_opts.bookmark = bookmarkName;

    // Step 2: insert the bookmark around the current selection.
    await wordRun(async ctx => {
      const selection = ctx.document.getSelection();
      selection.insertBookmark(bookmarkName);
      await ctx.sync();
    });
  }

  // Step 3: insert the TOC field with the bookmark switch.
  const insertOpts: InsertTocOptions = {
    placement: 'atCursor',
    scopeLabel,
  };
  if (bookmarkName) {
    insertOpts.bookmark = bookmarkName;
  }
  return insertToc(profile, insertOpts);
}

// ---------------------------------------------------------------------------
// Navigate to a TOC (go-to)
// ---------------------------------------------------------------------------

/**
 * Navigate to (select) the TOC field identified by fieldId.
 * No-op when fieldId is empty or wordApi15 is absent.
 */
export async function navigateToToc(fieldId: string): Promise<void> {
  if (!fieldId || !isCapable('wordApi15')) return;

  await wordRun(async ctx => {
    const fields = ctx.document.fields.getByTypes([Word.FieldType.toc]);
    (fields as unknown as { load(p: string): void }).load('id');
    await ctx.sync();

    const target = fields.items.find(f =>
      (f as unknown as { id?: string }).id === fieldId,
    );
    if (target) {
      target.result.select();
      await ctx.sync();
    }
  });
}

// ---------------------------------------------------------------------------
// Registry helpers
// ---------------------------------------------------------------------------

async function _loadRegistry(): Promise<TocRegistry> {
  try {
    const store = getStorageProvider().docSettings;
    const raw = await store.get(DOC_SETTINGS_KEY_REGISTRY);
    if (Array.isArray(raw)) return raw as TocRegistry;
  } catch { /* non-fatal */ }
  return [];
}

async function _saveRegistry(registry: TocRegistry): Promise<void> {
  const store = getStorageProvider().docSettings;
  await store.set(DOC_SETTINGS_KEY_REGISTRY, registry as unknown);
}

async function _appendToRegistry(entry: TocRegistryEntry): Promise<void> {
  const registry = await _loadRegistry();
  registry.push(entry);
  await _saveRegistry(registry);
}

/** Load the current TOC registry for display in the UI. */
export async function loadRegistry(): Promise<TocRegistry> {
  return _loadRegistry();
}

/** Clear the full registry (used in tests and on document re-open). */
export async function clearRegistry(): Promise<void> {
  await _saveRegistry([]);
}
