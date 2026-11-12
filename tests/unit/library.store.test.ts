/**
 * Tests for B2 — Library store (persist, index, export/import).
 *
 * Acceptance check B-T1: put → persisted + indexed + retrievable; export/import round-trip;
 * shared visibility gated on isConfigured().
 *
 * Uses the localStorage-backed local store (available in jsdom via tests/setup.ts).
 * Storage provider singleton is a fresh instance per test because beforeEach clears localStorage.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import {
  putClause,
  getClause,
  deleteClause,
  loadIndex,
  exportLibrary,
  importLibrary,
  rebuildIndex,
  hasSharedBackend,
} from '@features/content/library/library-store.js';
import type { ClauseEntry } from '@features/content/library/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeEntry(id: string, overrides: Partial<ClauseEntry> = {}): ClauseEntry {
  return {
    id,
    title: `Test Clause ${id}`,
    body: {
      html: `<p>Content ${id}</p>`,
      plainText: `Content ${id}`,
    },
    metadata: {
      jurisdiction: 'SG',
      practiceArea: 'Commercial',
      lastReviewed: '2025-06-01',
      approvedBy: 'J. Tan',
      visibility: 'personal',
    },
    audit: {
      createdBy: 'alice@firm',
      createdAt: '2025-06-01T00:00:00Z',
      updatedAt: '2025-06-01T00:00:00Z',
      version: 1,
    },
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// B-T1: Core CRUD + index
// ---------------------------------------------------------------------------

describe('library-store — put / get / delete / index (B-T1)', () => {
  beforeEach(() => localStorage.clear());

  it('put + get round-trips a full entry', async () => {
    const entry = makeEntry('clz-001');
    await putClause(entry);
    const retrieved = await getClause('clz-001');
    expect(retrieved).toMatchObject({
      id: 'clz-001',
      title: 'Test Clause clz-001',
    });
    expect(retrieved?.metadata.jurisdiction).toBe('SG');
  });

  it('stores entry in the index after put', async () => {
    const entry = makeEntry('clz-002');
    await putClause(entry);
    const index = await loadIndex();
    expect(index.some(r => r.id === 'clz-002')).toBe(true);
  });

  it('index record carries required search fields', async () => {
    const entry = makeEntry('clz-003');
    await putClause(entry);
    const index = await loadIndex();
    const rec = index.find(r => r.id === 'clz-003');
    expect(rec).toBeDefined();
    expect(rec?.title).toBe('Test Clause clz-003');
    expect(rec?.metadata.jurisdiction).toBe('SG');
    expect(rec?.metadata.practiceArea).toBe('Commercial');
    expect(rec?.plainText).toBe('Content clz-003');
    expect(rec?.version).toBe(1);
  });

  it('upserts an existing entry (same id)', async () => {
    const entry = makeEntry('clz-004');
    await putClause(entry);
    const updated = { ...entry, title: 'Updated Title', audit: { ...entry.audit, version: 2 } };
    await putClause(updated);

    const index = await loadIndex();
    const records = index.filter(r => r.id === 'clz-004');
    expect(records).toHaveLength(1);  // not duplicated
    expect(records[0]?.title).toBe('Updated Title');
    expect(records[0]?.version).toBe(2);
  });

  it('delete removes entry and updates index', async () => {
    const entry = makeEntry('clz-005');
    await putClause(entry);
    await deleteClause('clz-005');

    expect(await getClause('clz-005')).toBeUndefined();
    const index = await loadIndex();
    expect(index.some(r => r.id === 'clz-005')).toBe(false);
  });

  it('returns undefined for a missing entry', async () => {
    expect(await getClause('nonexistent')).toBeUndefined();
  });

  it('empty index when no entries saved', async () => {
    const index = await loadIndex();
    expect(index).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// Export / import round-trip
// ---------------------------------------------------------------------------

describe('library-store — export / import (B-T1)', () => {
  beforeEach(() => localStorage.clear());

  it('export produces a JSON Blob including the entry', async () => {
    const entry = makeEntry('clz-export-1');
    await putClause(entry);
    const blob = await exportLibrary();
    expect(blob.type).toBe('application/json');
    // Blob.text() may not be available in all jsdom versions; use FileReader fallback.
    const text = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(reader.error ?? new Error('FileReader failed'));
      reader.readAsText(blob);
    });
    expect(text).toContain('clz-export-1');
  });

  it('import merges entries into a fresh store', async () => {
    // Set up, export, clear, import
    const entry = makeEntry('clz-import-1');
    await putClause(entry);
    const blob = await exportLibrary();

    localStorage.clear();

    await importLibrary(blob);
    const retrieved = await getClause('clz-import-1');
    expect(retrieved?.id).toBe('clz-import-1');
  });

  it('import rebuilds the index from imported entries', async () => {
    const entry = makeEntry('clz-import-2');
    await putClause(entry);
    const blob = await exportLibrary();
    localStorage.clear();
    await importLibrary(blob);

    const index = await loadIndex();
    expect(index.some(r => r.id === 'clz-import-2')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Shared backend gate
// ---------------------------------------------------------------------------

describe('library-store — hasSharedBackend', () => {
  it('returns false when no remote adapter is configured (stub)', () => {
    // The storage provider is initialised with stub adapters (isConfigured() === false).
    expect(hasSharedBackend()).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// rebuildIndex
// ---------------------------------------------------------------------------

describe('library-store — rebuildIndex', () => {
  beforeEach(() => localStorage.clear());

  it('rebuilds an empty index when no entries exist', async () => {
    await rebuildIndex();
    const index = await loadIndex();
    expect(index).toEqual([]);
  });

  it('rebuilds index from existing entry keys', async () => {
    const e1 = makeEntry('clz-rebuild-1');
    const e2 = makeEntry('clz-rebuild-2');
    await putClause(e1);
    await putClause(e2);
    // Corrupt the index key directly to simulate drift
    localStorage.setItem('mutatis.library/index', '[]');
    await rebuildIndex();
    const index = await loadIndex();
    expect(index.some(r => r.id === 'clz-rebuild-1')).toBe(true);
    expect(index.some(r => r.id === 'clz-rebuild-2')).toBe(true);
  });
});
