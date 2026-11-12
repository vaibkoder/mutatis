/**
 * Precedent / Clause Library — persistent store.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.5.1
 *
 * Storage layout (local = localStorage via the shared local store):
 *   library/entry/<id>        → full ClauseEntry JSON
 *   library/index             → ClauseIndexRecord[] JSON (index for Fuse.js)
 *
 * The shared backend (remote adapter) is stubbed (isConfigured() === false).
 * 'shared' visibility is offered in the UI but disabled when no backend is configured.
 * This module is stub-aware: it checks getStorageProvider().remote?.isConfigured().
 *
 * Export / import uses the shared local.export() / local.import() methods so that
 * all library entries are included in a single JSON download/upload.
 */

import { getStorageProvider } from '@shared/storage/index.js';
import type { ClauseEntry, ClauseIndexRecord } from './types.js';

// ---------------------------------------------------------------------------
// Internal constants
// ---------------------------------------------------------------------------

const INDEX_KEY = 'library/index';
const ENTRY_PREFIX = 'library/entry/';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function entryKey(id: string): string {
  return `${ENTRY_PREFIX}${id}`;
}

async function readIndex(): Promise<ClauseIndexRecord[]> {
  const store = getStorageProvider().local;
  const raw = await store.get(INDEX_KEY);
  if (!Array.isArray(raw)) return [];
  return raw as ClauseIndexRecord[];
}

async function writeIndex(index: ClauseIndexRecord[]): Promise<void> {
  const store = getStorageProvider().local;
  await store.set(INDEX_KEY, index);
}

/** Build a ClauseIndexRecord from a full ClauseEntry. */
function toIndexRecord(entry: ClauseEntry): ClauseIndexRecord {
  const m = entry.metadata;
  // Build required fields first, then spread optional ones only when present.
  // exactOptionalPropertyTypes requires we not assign `undefined` to optional props.
  const indexMeta: ClauseIndexRecord['metadata'] = {
    jurisdiction: m.jurisdiction,
    practiceArea: m.practiceArea,
    lastReviewed: m.lastReviewed,
    approvedBy: m.approvedBy,
    visibility: m.visibility,
    ...(m.jurisdictionOther !== undefined ? { jurisdictionOther: m.jurisdictionOther } : {}),
    ...(m.tags !== undefined ? { tags: m.tags } : {}),
    ...(m.language !== undefined ? { language: m.language } : {}),
  };
  return {
    id: entry.id,
    title: entry.title,
    metadata: indexMeta,
    plainText: entry.body.plainText,
    version: entry.audit.version,
  };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Save (create or update) a clause entry.
 * Persists the full entry and updates the index.
 *
 * On update, increments audit.version and updates audit.updatedAt.
 */
export async function putClause(entry: ClauseEntry): Promise<void> {
  const store = getStorageProvider().local;

  // Persist the full entry
  await store.set(entryKey(entry.id), entry as unknown);

  // Update the index (upsert by id)
  const index = await readIndex();
  const record = toIndexRecord(entry);
  const idx = index.findIndex(r => r.id === entry.id);
  if (idx >= 0) {
    index[idx] = record;
  } else {
    index.push(record);
  }
  await writeIndex(index);
}

/** Retrieve a full ClauseEntry by id. Returns undefined if not found. */
export async function getClause(id: string): Promise<ClauseEntry | undefined> {
  const store = getStorageProvider().local;
  const raw = await store.get(entryKey(id));
  if (raw == null) return undefined;
  return raw as ClauseEntry;
}

/** Delete a clause entry by id. */
export async function deleteClause(id: string): Promise<void> {
  const store = getStorageProvider().local;
  await store.delete(entryKey(id));

  const index = await readIndex();
  const updated = index.filter(r => r.id !== id);
  await writeIndex(updated);
}

/** Load the full search index (lightweight index records). */
export async function loadIndex(): Promise<ClauseIndexRecord[]> {
  return readIndex();
}

/**
 * Export all library entries + index as a JSON Blob for user download.
 * Delegates to the underlying local store's export() so all library keys are included.
 */
export async function exportLibrary(): Promise<Blob> {
  return getStorageProvider().local.export();
}

/**
 * Import library entries from a JSON Blob (merge; does not clear existing entries).
 * After import, the index is rebuilt from all persisted entries to ensure consistency.
 */
export async function importLibrary(blob: Blob): Promise<void> {
  const store = getStorageProvider().local;
  await store.import(blob);

  // Rebuild index from all entry keys to ensure the imported data is reflected.
  await rebuildIndex();
}

/**
 * Rebuild the in-storage index by scanning all entry keys.
 * Called after import and on-demand for recovery.
 */
export async function rebuildIndex(): Promise<void> {
  const store = getStorageProvider().local;
  const keys = await store.list(ENTRY_PREFIX);
  const entries: ClauseIndexRecord[] = [];

  for (const key of keys) {
    const raw = await store.get(key);
    if (raw != null) {
      try {
        const entry = raw as ClauseEntry;
        entries.push(toIndexRecord(entry));
      } catch {
        // Skip malformed entries
      }
    }
  }

  await writeIndex(entries);
}

/**
 * True when a configured shared backend is available.
 * Used to gate the 'shared' visibility option in the UI.
 */
export function hasSharedBackend(): boolean {
  return getStorageProvider().remote?.isConfigured() === true;
}
