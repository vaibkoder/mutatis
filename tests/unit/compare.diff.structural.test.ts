import { describe, it, expect } from 'vitest';

import type { ParaRecord } from '@features/compare/model.js';
import { structuralDiff, computeLCS } from '@features/compare/diff/index.js';
import { fnv1a } from '@features/compare/preprocess/index.js';

const AUTHOR = 'Test';
const DATE = '2026-06-08T00:00:00.000Z';

/** Build a minimal ParaRecord whose hash is derived from `text`. */
function para(index: number, text: string): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: `<w:p>${text}</w:p>`,
    normText: text,
    normTextAggressive: text.toLowerCase(),
    numberingPrefix: '',
    hash: fnv1a(text),
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  };
}

function side(...texts: string[]): ParaRecord[] {
  return texts.map((t, i) => para(i, t));
}

describe('computeLCS', () => {
  it('returns matched index pairs in order', () => {
    const pairs = computeLCS(['a', 'b', 'c', 'd'], ['a', 'x', 'c', 'd'], (s) => s);
    expect(pairs).toEqual([
      [0, 0],
      [2, 2],
      [3, 3],
    ]);
  });

  it('handles empty inputs', () => {
    expect(computeLCS<string>([], ['a'], (s) => s)).toEqual([]);
    expect(computeLCS(['a'], [], (s) => s)).toEqual([]);
  });
});

describe('structuralDiff — T1 (insert mid-document, no cascade)', () => {
  it('produces exactly one insert and leaves every following paragraph equal', () => {
    const original = side('A', 'B', 'C', 'D', 'E');
    // Insert "B2" between B and C; everything else is verbatim.
    const revised = side('A', 'B', 'B2', 'C', 'D', 'E');

    const ops = structuralDiff(original, revised, AUTHOR, DATE);

    const inserts = ops.filter((o) => o.type === 'insert');
    const equals = ops.filter((o) => o.type === 'equal');
    const deletes = ops.filter((o) => o.type === 'delete');
    const replaces = ops.filter((o) => o.type === 'replace');

    expect(inserts).toHaveLength(1);
    expect(deletes).toHaveLength(0);
    expect(replaces).toHaveLength(0);
    // A, B, C, D, E all remained equal — no cascade misalignment.
    expect(equals).toHaveLength(5);

    const ins = inserts[0]!;
    expect(ins.revRange).toEqual({ start: 2, end: 2 });
    expect(ins.origRange).toBeNull();
    expect(ins.classification).toBe('structural');
  });
});

describe('structuralDiff — anchors, replace, delete', () => {
  it('emits a delete for an original-only paragraph', () => {
    const ops = structuralDiff(side('A', 'B', 'C'), side('A', 'C'), AUTHOR, DATE);
    const del = ops.filter((o) => o.type === 'delete');
    expect(del).toHaveLength(1);
    expect(del[0]!.origRange).toEqual({ start: 1, end: 1 });
    expect(del[0]!.revRange).toBeNull();
  });

  it('emits a replace for an index-aligned, hash-unequal pair between anchors', () => {
    // A and C are unique anchors; the middle paragraph changed text (B → B-edited).
    const ops = structuralDiff(
      side('A', 'B original', 'C'),
      side('A', 'B edited', 'C'),
      AUTHOR,
      DATE,
    );
    const replaces = ops.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(replaces[0]!.origRange).toEqual({ start: 1, end: 1 });
    expect(replaces[0]!.revRange).toEqual({ start: 1, end: 1 });
    // wordDiffs are filled later (A3/I1) — not set here.
    expect(replaces[0]!.wordDiffs).toBeUndefined();
  });

  it('assigns monotonic op ids and injects author/date', () => {
    const ops = structuralDiff(side('A', 'B'), side('A', 'B'), AUTHOR, DATE);
    expect(ops.map((o) => o.id)).toEqual(['op-0', 'op-1']);
    for (const op of ops) {
      expect(op.author).toBe(AUTHOR);
      expect(op.date).toBe(DATE);
    }
  });
});
