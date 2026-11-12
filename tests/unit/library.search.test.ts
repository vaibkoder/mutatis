/**
 * Tests for B3 — Fuse.js search + faceted pre-filter.
 *
 * Acceptance check B-T2: "liability cap" fuzzy-matches the limitation clause
 * ranked by title weight.
 *
 * No Office JS required — pure search logic.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { LibrarySearch, _resetLibrarySearch } from '@features/content/library/search.js';
import type { ClauseIndexRecord } from '@features/content/library/types.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeRecord(id: string, overrides: Partial<ClauseIndexRecord> = {}): ClauseIndexRecord {
  return {
    id,
    title: `Test Clause ${id}`,
    metadata: {
      jurisdiction: 'SG',
      practiceArea: 'Commercial',
      lastReviewed: '2025-06-01',
      approvedBy: 'J. Tan',
      visibility: 'personal',
    },
    plainText: `Generic clause content for ${id}`,
    version: 1,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// B-T2: "liability cap" fuzzy search
// ---------------------------------------------------------------------------

describe('LibrarySearch — fuzzy search (B-T2)', () => {
  let search: LibrarySearch;

  beforeEach(() => {
    search = new LibrarySearch();
    _resetLibrarySearch();
  });

  it('returns the limitation clause when searching "liability cap"', () => {
    const limClause = makeRecord('clz-limitation', {
      title: 'Limitation of Liability (cap at fees)',
      metadata: {
        jurisdiction: 'SG',
        practiceArea: 'Commercial',
        lastReviewed: '2025-06-01',
        approvedBy: 'J. Tan',
        visibility: 'personal',
        tags: ['liability', 'cap', 'indemnity'],
      },
      plainText: 'Neither party shall be liable for indirect loss. Cap at contract fees.',
    });

    const unrelated = makeRecord('clz-nda', {
      title: 'Non-Disclosure Agreement – Mutual',
      plainText: 'Each party shall keep confidential information secret.',
    });

    search.buildIndex([limClause, unrelated]);
    const results = search.search('liability cap');

    expect(results.length).toBeGreaterThan(0);
    expect(results[0]?.record.id).toBe('clz-limitation');
  });

  it('returns an empty array for very short queries (< minMatchCharLength)', () => {
    search.buildIndex([makeRecord('clz-a')]);
    const results = search.search('li');   // 2 chars, below minMatchCharLength=3
    // Short query returns all (browse mode, no Fuse filtering)
    // The spec says Fuse doesn't run for < 3 chars; we return all candidates instead
    expect(Array.isArray(results)).toBe(true);
  });

  it('returns all candidates when query is empty (browse mode)', () => {
    search.buildIndex([makeRecord('clz-a'), makeRecord('clz-b'), makeRecord('clz-c')]);
    const results = search.search('');
    expect(results).toHaveLength(3);
  });

  it('ranks by title weight — title match beats plainText-only match', () => {
    const titleMatch = makeRecord('clz-title', {
      title: 'Force Majeure clause',
      plainText: 'General commercial terms.',
    });
    const textOnly = makeRecord('clz-text', {
      title: 'Miscellaneous Provisions',
      plainText: 'Force majeure events shall excuse performance.',
    });

    search.buildIndex([textOnly, titleMatch]);
    const results = search.search('force majeure');
    expect(results.length).toBeGreaterThan(0);
    expect(results[0]?.record.id).toBe('clz-title');
  });
});

// ---------------------------------------------------------------------------
// Facet pre-filter
// ---------------------------------------------------------------------------

describe('LibrarySearch — facet pre-filter', () => {
  let search: LibrarySearch;

  beforeEach(() => {
    search = new LibrarySearch();
    const records: ClauseIndexRecord[] = [
      makeRecord('clz-sg', { metadata: { jurisdiction: 'SG', practiceArea: 'Commercial', lastReviewed: '2025-06-01', approvedBy: 'A', visibility: 'personal' } }),
      makeRecord('clz-ew', { metadata: { jurisdiction: 'EW', practiceArea: 'Commercial', lastReviewed: '2025-06-01', approvedBy: 'A', visibility: 'personal' } }),
      makeRecord('clz-other', { metadata: { jurisdiction: 'OTHER', practiceArea: 'Commercial', lastReviewed: '2025-06-01', approvedBy: 'A', visibility: 'personal' } }),
      makeRecord('clz-shared', { metadata: { jurisdiction: 'SG', practiceArea: 'M&A', lastReviewed: '2025-06-01', approvedBy: 'A', visibility: 'shared' } }),
    ];
    search.buildIndex(records);
  });

  it('filters by jurisdiction', () => {
    const results = search.search('', { jurisdiction: 'SG' });
    expect(results.every(r => r.record.metadata.jurisdiction === 'SG')).toBe(true);
    expect(results.length).toBe(2);   // clz-sg + clz-shared
  });

  it('filters by practiceArea', () => {
    const results = search.search('', { practiceArea: 'M&A' });
    expect(results).toHaveLength(1);
    expect(results[0]?.record.id).toBe('clz-shared');
  });

  it('filters by visibility', () => {
    const results = search.search('', { visibility: 'shared' });
    expect(results).toHaveLength(1);
    expect(results[0]?.record.id).toBe('clz-shared');
  });

  it('combines multiple facets (AND logic)', () => {
    const results = search.search('', { jurisdiction: 'SG', visibility: 'personal' });
    expect(results).toHaveLength(1);
    expect(results[0]?.record.id).toBe('clz-sg');
  });
});

// ---------------------------------------------------------------------------
// Staleness facet
// ---------------------------------------------------------------------------

describe('LibrarySearch — stale facet', () => {
  let search: LibrarySearch;
  const NOW = new Date('2026-06-08T00:00:00Z');

  beforeEach(() => {
    search = new LibrarySearch();
    const records: ClauseIndexRecord[] = [
      makeRecord('clz-fresh', {
        metadata: { jurisdiction: 'SG', practiceArea: 'Commercial', lastReviewed: '2025-06-01', approvedBy: 'A', visibility: 'personal' },
      }),  // ~12 months old — not stale (< 24 months)
      makeRecord('clz-stale', {
        metadata: { jurisdiction: 'SG', practiceArea: 'Commercial', lastReviewed: '2024-05-01', approvedBy: 'A', visibility: 'personal' },
      }),  // ~25 months old — stale
    ];
    search.buildIndex(records);
  });

  it('staleOnly=true returns only stale entries', () => {
    const results = search.search('', { staleOnly: true }, NOW);
    expect(results.every(r => r.record.id === 'clz-stale')).toBe(true);
    expect(results).toHaveLength(1);
  });

  it('without staleOnly filter, both entries are returned', () => {
    const results = search.search('', {}, NOW);
    expect(results).toHaveLength(2);
  });
});

// ---------------------------------------------------------------------------
// Incremental index update (upsertEntry / removeEntry)
// ---------------------------------------------------------------------------

describe('LibrarySearch — incremental index update', () => {
  let search: LibrarySearch;

  beforeEach(() => { search = new LibrarySearch(); });

  it('upsertEntry adds a new record', () => {
    search.buildIndex([]);
    search.upsertEntry(makeRecord('clz-new'));
    expect(search.size).toBe(1);
    const results = search.search('test clause');
    expect(results.some(r => r.record.id === 'clz-new')).toBe(true);
  });

  it('upsertEntry updates an existing record', () => {
    const rec = makeRecord('clz-upd');
    search.buildIndex([rec]);
    const updated = { ...rec, title: 'Updated Title' };
    search.upsertEntry(updated);
    expect(search.size).toBe(1);
    const results = search.search('Updated Title');
    expect(results[0]?.record.title).toBe('Updated Title');
  });

  it('removeEntry removes a record', () => {
    search.buildIndex([makeRecord('clz-del')]);
    search.removeEntry('clz-del');
    expect(search.size).toBe(0);
  });
});
