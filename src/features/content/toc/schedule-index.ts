/**
 * Feature 07-A — Table of Contents: Index of Schedules.
 *
 * Two strategies per spec §A.5.4:
 *
 * Strategy 1 (preferred — stays live):
 *   A custom-style TOC field using ONLY \t "Schedule Heading,1;…" \h with NO
 *   \o switch, so body heading styles are excluded. Delegates to
 *   toc-service.insertScheduleIndexField() (which is Strategy 1-aware).
 *
 * Strategy 2 (fallback — snapshot):
 *   When Schedule paragraphs lack a consistent style, or when the user
 *   prefers a static list. Uses DocModel.getParagraphsByStyle() to collect
 *   paragraphs, inserts a bookmark at each, then inserts an HTML snapshot
 *   labelled "snapshot — Regenerate". The snapshot is NOT a live field.
 *
 *   Regenerate is idempotent (spec §A.8):
 *     - Bookmark names are counter-suffixed and verified unique against
 *       document.bookmarks (mirrors generateBookmarkName() in toc-service.ts),
 *       so a re-run never clobbers a user's pre-existing bookmark.
 *     - The snapshot is wrapped in a content control tagged
 *       IDX_SCHEDULE_SNAPSHOT; a re-run deletes the prior snapshot before
 *       inserting the new one, so exactly one snapshot ever exists — they no
 *       longer stack on repeated regeneration.
 *
 * SG/UK rationale: Schedules often use a bespoke "Schedule Heading" style and
 * restart numbering, making a plain \o "1-3" TOC miss them. Strategy 1 is the
 * correct default; Strategy 2 covers legacy precedents with inconsistent styling.
 *
 * Spec reference: spec/features/07-toc-precedent-templates.md §A.5.4.
 */

import { wordRun } from '@shared/office/index.js';
import { DocModel } from '@shared/doc-model/index.js';
import type { ScheduleIndexConfig, ScheduleSnapshotEntry } from './types.js';
import { insertScheduleIndexField } from './toc-service.js';

// ---------------------------------------------------------------------------
// Strategy 1 — live custom-style TOC field
// ---------------------------------------------------------------------------

export interface StrategyOneResult {
  strategy: 1;
  usedNativeField: boolean;
  usedFallback: boolean;
  /** True when no Schedule-styled paragraphs were found after updateResult(). */
  isEmpty: boolean;
}

/**
 * Generate the Index of Schedules using Strategy 1 (live TOC field).
 *
 * Delegates to toc-service.insertScheduleIndexField(); the switch string uses
 * ONLY \t for the configured schedule styles and NO \o.
 *
 * @param config    Schedule index configuration.
 * @param placement Where to insert the field ("atCursor" or "topOfDocument").
 */
export async function generateScheduleIndexStrategy1(
  config: ScheduleIndexConfig,
  placement: 'atCursor' | 'topOfDocument' = 'topOfDocument',
): Promise<StrategyOneResult> {
  const { usedNativeField, usedFallback, isEmpty } = await insertScheduleIndexField(
    config,
    placement,
  );
  return { strategy: 1, usedNativeField, usedFallback, isEmpty };
}

// ---------------------------------------------------------------------------
// Strategy 2 — collected snapshot (fallback)
// ---------------------------------------------------------------------------

export interface StrategyTwoResult {
  strategy: 2;
  entries: ScheduleSnapshotEntry[];
  /** True when getParagraphsByStyle() returned no matching paragraphs. */
  isEmpty: boolean;
  /** Informational message to show in the UI above the snapshot list. */
  snapshotLabel: string;
}

/** Label appended to the snapshot heading — always en-GB, never "Exhibit". */
const SNAPSHOT_LABEL = 'snapshot — Regenerate to update';

/**
 * Content-control tag marking the Strategy-2 snapshot block. On Regenerate we
 * find the prior snapshot by this tag and delete it before inserting the new
 * one, so snapshots replace-in-place rather than stacking (spec §A.8).
 */
const SNAPSHOT_CC_TAG = 'IDX_SCHEDULE_SNAPSHOT';

/**
 * Pick the next collision-free IDX_SCHEDULE_<n> bookmark name.
 *
 * Mirrors the counter-suffixed uniqueness approach of generateBookmarkName() in
 * toc-service.ts (spec §A.8 bookmark-collision rule): increment n until the name
 * is absent from the in-use set, so we never overwrite a bookmark the user (or a
 * prior snapshot) already placed. Pure — the caller owns the in-use set.
 */
function nextScheduleBookmarkName(usedNames: Set<string>): string {
  let counter = 1;
  let name = `IDX_SCHEDULE_${counter}`;
  while (usedNames.has(name)) {
    counter++;
    name = `IDX_SCHEDULE_${counter}`;
  }
  return name;
}

/**
 * Generate the Index of Schedules using Strategy 2 (HTML snapshot).
 *
 * Collects all paragraphs matching the configured schedule styles via
 * DocModel.getParagraphsByStyle(). For each paragraph a bookmark is inserted
 * (to enable anchor links), then an HTML list is inserted at the placement
 * location labelled as a snapshot.
 *
 * NOTE: This function requires DocModel to have been loaded first.
 *
 * @param config    Schedule index configuration.
 * @param placement Where to insert the snapshot HTML ("atCursor" | "topOfDocument").
 */
export async function generateScheduleIndexStrategy2(
  config: ScheduleIndexConfig,
  placement: 'atCursor' | 'topOfDocument' = 'topOfDocument',
): Promise<StrategyTwoResult> {
  // 1. Collect paragraphs by style from the cached DocModel.
  const paras = DocModel.getParagraphsByStyle(config.scheduleStyles);

  if (paras.length === 0) {
    // Regenerate-to-empty is still a regenerate: drop any prior snapshot so the
    // document reflects the current (now empty) state rather than keeping a
    // stale list. Without this, removing every Schedule then Regenerating would
    // leave the old snapshot in place and look like a bug. Idempotent (§A.8).
    await wordRun(async ctx => {
      const priorSnapshots = ctx.document.contentControls.getByTag(SNAPSHOT_CC_TAG);
      priorSnapshots.load('items');
      await ctx.sync();
      for (const cc of priorSnapshots.items) {
        cc.delete(false);
      }
      await ctx.sync();
    });

    return {
      strategy: 2,
      entries: [],
      isEmpty: true,
      snapshotLabel: SNAPSHOT_LABEL,
    };
  }

  // 2. Build entries; bookmark names are assigned inside the Word.run once the
  //    document's existing bookmark names are known (collision-checked).
  const entries: ScheduleSnapshotEntry[] = paras.map(para => ({
    paragraphId: para.id,
    text: para.text,
    bookmark: '',
  }));

  // 3. Insert bookmarks at each paragraph and then the HTML snapshot.
  await wordRun(async ctx => {
    // 3a. Stage reads needed before we mutate: the prior snapshot (to replace
    //     it) and existing bookmark names (to avoid collisions). Load both,
    //     then a single sync resolves them together.
    const priorSnapshots = ctx.document.contentControls.getByTag(SNAPSHOT_CC_TAG);
    priorSnapshots.load('items');

    const bookmarks = ctx.document.bookmarks;
    bookmarks.load('name');

    await ctx.sync();

    // 3b. Replace-in-place: delete any prior snapshot (control + its content)
    //     so regeneration never stacks a second list on top of the first.
    for (const cc of priorSnapshots.items) {
      cc.delete(false);
    }

    // 3c. Assign counter-suffixed, collision-free bookmark names. The prior
    //     snapshot's own paragraph bookmarks remain in the document, so the
    //     in-use set seeds from the live names and each new name steps past them.
    const usedNames = new Set(bookmarks.items.map(b => b.name));
    for (const entry of entries) {
      entry.bookmark = nextScheduleBookmarkName(usedNames);
      usedNames.add(entry.bookmark);
    }

    // 3d. Insert a bookmark at the start of each schedule paragraph.
    for (const entry of entries) {
      // Re-resolve a live range for this paragraph within this Word.run.
      // DocModel nodes carry a closure that re-fetches the range in the current ctx.
      const para = paras.find(p => p.id === entry.paragraphId);
      if (!para) continue;

      const range = para.getRange(ctx);
      range.insertBookmark(entry.bookmark);
    }

    // 3e. Build the HTML snapshot list and insert it at the target location,
    //     then wrap the inserted range in a tagged content control so the next
    //     Regenerate can find and replace exactly this block.
    const snapshotHtml = _buildSnapshotHtml(entries, config.label);

    const targetRange = placement === 'topOfDocument'
      ? ctx.document.body.getRange('Start')
      : ctx.document.getSelection().getRange('Start');

    const inserted = targetRange.insertHtml(snapshotHtml, Word.InsertLocation.before);
    const snapshotCc = inserted.insertContentControl();
    // tag/title are typed plain-string setters — never XML-escape (Office
    // serialises internally; escaping would double-encode).
    snapshotCc.tag = SNAPSHOT_CC_TAG;
    snapshotCc.title = config.label;

    // One sync for all mutations (batched).
    await ctx.sync();
  });

  return {
    strategy: 2,
    entries,
    isEmpty: false,
    snapshotLabel: SNAPSHOT_LABEL,
  };
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Build the HTML for the Strategy 2 snapshot.
 * Uses anchor links to the paragraph bookmarks.
 *
 * Labelled explicitly as a snapshot so users know it needs manual regeneration.
 * The label comes from config (never "Table of Authorities").
 */
function _buildSnapshotHtml(entries: ScheduleSnapshotEntry[], label: string): string {
  const heading = `<p><strong>${_escapeHtml(label)} (${SNAPSHOT_LABEL})</strong></p>`;
  const items = entries.map(entry => {
    const text = _escapeHtml(entry.text);
    // Use a named anchor link to the bookmark.
    return `<p><a href="#${entry.bookmark}">${text}</a></p>`;
  });
  return heading + items.join('');
}

/**
 * Escape special HTML characters to prevent XSS / malformed markup when
 * paragraph text is inserted into the HTML string.
 */
function _escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ---------------------------------------------------------------------------
// Exported config helper — re-export for convenience of callers
// ---------------------------------------------------------------------------

export { DEFAULT_SCHEDULE_INDEX_CONFIG } from './config.js';
