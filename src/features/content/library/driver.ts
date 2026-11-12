/**
 * Precedent / Clause Library — Preact Signals driver.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.5.1, §B.5.2, §B.5.3, §B.8
 *
 * Holds all UI state for the LibraryPanel: search results, save form, insert flow,
 * drift check. All Office JS calls go through wordRun() (shared/office).
 *
 * B4 — Save-clause flow:
 *   getSelection → block empty → getHtml + getOoxml + range.text →
 *   normalise on save → build entry (uuid) → library-store.put → update Fuse index → toast.
 *
 * B5 — Click-to-insert:
 *   jurisdiction gate (mandatory, non-suppressible) → insertHtml / insertOoxml →
 *   wrap in CLAUSE content control.
 *
 * B6 — Drift detection:
 *   scan content controls (load "tag,title") → parse CLAUSE:<id>:v<n> →
 *   compare vs library version → flag drifted inserts.
 */

import { signal, computed } from '@preact/signals';
import type { ClauseEntry, DriftResult, SaveClauseInput } from './types.js';
import { LARGE_LIBRARY_THRESHOLD } from './types.js';
import {
  validateClauseMetadata,
  validateTitle,
  validateSelectionNonEmpty,
} from './validation.js';
import {
  putClause,
  loadIndex,
  hasSharedBackend,
} from './library-store.js';
import {
  getLibrarySearch,
} from './search.js';
import type { FacetFilters, SearchResult } from './search.js';
import {
  requiresJurisdictionWarning,
  jurisdictionLabel,
  docJurisdictionLabel,
  getDocJurisdiction,
  insertClause,
  normaliseQuotesAndEncoding,
  stripControlChars,
} from './insert.js';
import { scanDrift } from './drift.js';
import { wordRun } from '@shared/office/index.js';

// ---------------------------------------------------------------------------
// UUID generation (client-side, no network)
// ---------------------------------------------------------------------------

function generateId(): string {
  // Use crypto.randomUUID when available (modern browsers + WebView2); fallback to Math.random.
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  // Fallback: RFC 4122 v4 UUID from Math.random
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.floor(Math.random() * 16);
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

// ---------------------------------------------------------------------------
// Signals — search + browse state
// ---------------------------------------------------------------------------

export const librarySearchQuery = signal<string>('');
export const libraryFacets = signal<FacetFilters>({});
export const libraryResults = signal<SearchResult[]>([]);
export const libraryBusy = signal<boolean>(false);
export const libraryError = signal<string | null>(null);
export const libraryIndexSize = signal<number>(0);

// ---------------------------------------------------------------------------
// Signals — save form state
// ---------------------------------------------------------------------------

export const saveFormOpen = signal<boolean>(false);
export const saveFormTitle = signal<string>('');
export const saveFormMetadata = signal<Partial<import('./types.js').ClauseMetadata>>({});
export const saveFormErrors = signal<import('./validation.js').ValidationError[]>([]);
export const saveFormWarnings = signal<import('./validation.js').ValidationWarning[]>([]);
/** The captured body from the selection (set when the save form opens). */
export const saveFormBody = signal<import('./types.js').ClauseBody | null>(null);
export const saveBusy = signal<boolean>(false);
export const saveSuccess = signal<string | null>(null);   // toast message

// ---------------------------------------------------------------------------
// Signals — insert confirm modal
// ---------------------------------------------------------------------------

export const insertModalOpen = signal<boolean>(false);
export const insertModalWarning = signal<string>('');
/** The entry pending insertion (waiting for user confirm). */
export const insertPendingEntry = signal<ClauseEntry | null>(null);
export const insertPreserveTables = signal<boolean>(false);

// ---------------------------------------------------------------------------
// Signals — drift detection
// ---------------------------------------------------------------------------

export const driftResults = signal<DriftResult[]>([]);
export const driftBusy = signal<boolean>(false);
export const driftError = signal<string | null>(null);

// ---------------------------------------------------------------------------
// Derived
// ---------------------------------------------------------------------------

export const saveFormValid = computed(() => {
  const titleErr = validateTitle(saveFormTitle.value);
  if (titleErr) return false;
  const result = validateClauseMetadata(
    saveFormMetadata.value,
    hasSharedBackend(),
  );
  return result.valid;
});

// ---------------------------------------------------------------------------
// B4 — Save-clause flow
// ---------------------------------------------------------------------------

/**
 * Open the save form: capture the current Word selection (getHtml + getOoxml + text).
 * Blocks if the selection is empty. On success, opens the save form and pre-populates body.
 */
export async function openSaveForm(): Promise<void> {
  libraryError.value = null;
  libraryBusy.value = true;

  try {
    let html = '';
    let ooxml: string | undefined;
    let plainText = '';

    await wordRun(async ctx => {
      const sel = ctx.document.getSelection();
      sel.load('isEmpty,text');
      await ctx.sync();

      // Block empty selection
      if (sel.isEmpty) {
        throw new Error('Select a clause in the document before saving.');
      }

      plainText = sel.text ?? '';

      // Capture HTML (always)
      const htmlResult = sel.getHtml();
      // Capture OOXML (for table-fidelity insert option)
      const ooxmlResult = sel.getOoxml();
      await ctx.sync();

      html = htmlResult.value ?? '';
      ooxml = ooxmlResult.value ?? undefined;
    });

    // Validate non-empty (defensive)
    if (!validateSelectionNonEmpty(html, plainText)) {
      throw new Error('The selection appears to be empty. Select a clause and try again.');
    }

    // Normalise smart quotes and strip control chars on save (spec §B.8)
    const normHtml = normaliseQuotesAndEncoding(stripControlChars(html));
    const normOoxml = ooxml ? stripControlChars(ooxml) : undefined;
    const normPlainText = stripControlChars(plainText);

    saveFormBody.value = {
      html: normHtml,
      plainText: normPlainText,
      ...(normOoxml !== undefined ? { ooxml: normOoxml } : {}),
    };

    // Reset form fields
    saveFormTitle.value = '';
    saveFormMetadata.value = {};
    saveFormErrors.value = [];
    saveFormWarnings.value = [];
    saveSuccess.value = null;
    saveFormOpen.value = true;

  } catch (err) {
    libraryError.value = err instanceof Error ? err.message : String(err);
  } finally {
    libraryBusy.value = false;
  }
}

/**
 * Submit the save form: validate → build entry → persist → update Fuse index → toast.
 */
export async function submitSaveForm(input: Omit<SaveClauseInput, 'body'>): Promise<void> {
  saveBusy.value = true;
  saveFormErrors.value = [];
  saveFormWarnings.value = [];
  saveSuccess.value = null;

  // Validate title
  const titleErr = validateTitle(input.title);
  if (titleErr) {
    saveFormErrors.value = [titleErr];
    saveBusy.value = false;
    return;
  }

  // Validate metadata
  const metaResult = validateClauseMetadata(
    input.metadata as Partial<import('./types.js').ClauseMetadata>,
    hasSharedBackend(),
  );
  if (!metaResult.valid) {
    saveFormErrors.value = metaResult.errors;
    saveFormWarnings.value = metaResult.warnings;
    saveBusy.value = false;
    return;
  }

  // Warn but do not block on staleness
  saveFormWarnings.value = metaResult.warnings;

  try {
    const body = saveFormBody.value;
    if (!body) throw new Error('No clause body captured. Please open the save form again.');

    const now = new Date();
    const entry: ClauseEntry = {
      id: generateId(),
      title: input.title.trim(),
      body,
      metadata: input.metadata as import('./types.js').ClauseMetadata,
      audit: {
        createdBy: input.createdBy,
        createdAt: now.toISOString(),
        updatedAt: now.toISOString(),
        version: 1,
      },
    };

    // Persist
    await putClause(entry);

    // Update Fuse index incrementally
    const search = getLibrarySearch();
    const em = entry.metadata;
    search.upsertEntry({
      id: entry.id,
      title: entry.title,
      metadata: {
        jurisdiction: em.jurisdiction,
        practiceArea: em.practiceArea,
        lastReviewed: em.lastReviewed,
        approvedBy: em.approvedBy,
        visibility: em.visibility,
        ...(em.jurisdictionOther !== undefined ? { jurisdictionOther: em.jurisdictionOther } : {}),
        ...(em.tags !== undefined ? { tags: em.tags } : {}),
        ...(em.language !== undefined ? { language: em.language } : {}),
      },
      plainText: entry.body.plainText,
      version: entry.audit.version,
    });
    libraryIndexSize.value = search.size;

    // Success toast
    saveSuccess.value = `Clause "${entry.title}" saved to the library.`;
    saveFormOpen.value = false;

  } catch (err) {
    saveFormErrors.value = [{
      field: '_',
      message: err instanceof Error ? err.message : 'An unexpected error occurred while saving.',
    }];
  } finally {
    saveBusy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Search / index loading
// ---------------------------------------------------------------------------

/**
 * Initialise (or re-initialise) the Fuse search index from the persisted store.
 * Call once when the library panel mounts.
 */
export async function initLibrarySearch(): Promise<void> {
  libraryBusy.value = true;
  libraryError.value = null;
  try {
    const records = await loadIndex();
    const search = getLibrarySearch();
    search.buildIndex(records);
    libraryIndexSize.value = search.size;
    // Run the current query (if any) after loading
    runSearch();
  } catch (err) {
    libraryError.value = err instanceof Error ? err.message : String(err);
  } finally {
    libraryBusy.value = false;
  }
}

/** Run the current search query + facets and update libraryResults. */
export function runSearch(): void {
  const search = getLibrarySearch();
  const results = search.search(librarySearchQuery.value, libraryFacets.value);
  libraryResults.value = results;
}

// ---------------------------------------------------------------------------
// B5 — Click-to-insert with jurisdiction warning gate
// ---------------------------------------------------------------------------

/**
 * Initiate the insert flow for a result card.
 * Checks the jurisdiction gate (mandatory, non-suppressible).
 * If a warning is required, opens the confirmation modal.
 * If no warning is required, inserts immediately.
 *
 * @param entry         The full ClauseEntry to insert.
 * @param preserveTables Whether to use the OOXML path for table fidelity.
 */
export async function initiateInsert(
  entry: ClauseEntry,
  preserveTables: boolean,
): Promise<void> {
  const docJur = getDocJurisdiction();
  const warn = requiresJurisdictionWarning(entry.metadata.jurisdiction, docJur);

  if (warn) {
    // Show the blocking jurisdiction warning modal (non-suppressible).
    const clauseLabel = jurisdictionLabel(
      entry.metadata.jurisdiction,
      entry.metadata.jurisdictionOther,
    );
    const docLabel = docJurisdictionLabel(docJur);
    insertModalWarning.value =
      `This clause is governed by ${clauseLabel}. `
      + `The current document is governed by ${docLabel}. `
      + `Inserting a cross-jurisdiction clause is a known professional risk. `
      + `Insert anyway?`;
    insertPendingEntry.value = entry;
    insertPreserveTables.value = preserveTables;
    insertModalOpen.value = true;
    return;
  }

  // Same jurisdiction: insert immediately.
  await performInsert(entry, preserveTables);
}

/**
 * Confirm the jurisdiction warning modal and proceed with insertion.
 * Must only be called from the modal's explicit "Insert anyway" action.
 */
export async function confirmInsert(): Promise<void> {
  const entry = insertPendingEntry.value;
  if (!entry) return;
  insertModalOpen.value = false;
  await performInsert(entry, insertPreserveTables.value);
}

/** Cancel the jurisdiction warning modal. Insert does NOT proceed. */
export function cancelInsert(): void {
  insertModalOpen.value = false;
  insertPendingEntry.value = null;
  insertModalWarning.value = '';
}

/**
 * Internal: perform the actual insert (after the jurisdiction gate has passed or been confirmed).
 */
async function performInsert(entry: ClauseEntry, preserveTables: boolean): Promise<void> {
  libraryBusy.value = true;
  libraryError.value = null;
  try {
    await insertClause(entry, { preserveTables });
  } catch (err) {
    libraryError.value = err instanceof Error ? err.message : String(err);
  } finally {
    libraryBusy.value = false;
    insertPendingEntry.value = null;
  }
}

// ---------------------------------------------------------------------------
// B6 — Drift detection
// ---------------------------------------------------------------------------

/**
 * Scan the document's content controls for CLAUSE tracking tags and compare
 * the stored version against the current library version. Surfaces drifted inserts.
 */
export async function checkDrift(): Promise<void> {
  driftBusy.value = true;
  driftError.value = null;
  driftResults.value = [];

  try {
    const results = await scanDrift();
    driftResults.value = results;
  } catch (err) {
    driftError.value = err instanceof Error ? err.message : String(err);
  } finally {
    driftBusy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Convenience: check whether the library index exceeds the large-library threshold
// ---------------------------------------------------------------------------

export function isLargeLibrary(): boolean {
  return libraryIndexSize.value > LARGE_LIBRARY_THRESHOLD;
}
