/**
 * B3 — PDF fuzzy similarity fallback. Two textually near-identical PDF paragraphs
 * that exact-hash anchoring can't pair (boundary noise) must anchor via fuzzy
 * match as a single `replace`, badged low-confidence, with NOTHING suppressed.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';

import type { DiffOp, ParaRecord } from '../../../src/features/compare/model.js';
import { fuzzyAnchorPdf, similarity } from '../../../src/features/compare/pdf/fuzzy-anchor.js';
import { runPipelinePdf } from '../../../src/features/compare/pipeline.js';
import { fnv1a } from '../../../src/features/compare/preprocess/hash.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

function rec(index: number, text: string): ParaRecord {
  return {
    index,
    partId: 'document',
    rawXml: '',
    normText: text,
    normTextAggressive: text.toLowerCase().replace(/[^a-z0-9 ]/g, ''),
    numberingPrefix: '',
    hash: fnv1a(text),
    styleId: null,
    isTableCell: false,
    isSignatureBlock: false,
    isHeaderFooter: false,
    isFootnote: false,
  };
}

describe('B3 — similarity metric', () => {
  it('scores near-identical high and disjoint low', () => {
    expect(similarity('the seller shall deliver the goods', 'the seller shall deliver the goods')).toBe(1);
    expect(
      similarity('the seller shall deliver the goods', 'the seller shall deliver the items'),
    ).toBeGreaterThan(0.85);
    expect(similarity('governing law clause', 'payment of consideration')).toBeLessThan(0.85);
  });
});

/** Build a delete op referencing original[i]. */
function delOp(id: string, i: number): DiffOp {
  return {
    id,
    type: 'delete',
    origRange: { start: i, end: i },
    revRange: null,
    partId: 'document',
    classification: 'structural',
    classes: ['structural'],
    author: 'A',
    date: 'D',
  };
}

/** Build an insert op referencing revised[i]. */
function insOp(id: string, i: number): DiffOp {
  return {
    id,
    type: 'insert',
    origRange: null,
    revRange: { start: i, end: i },
    partId: 'document',
    classification: 'structural',
    classes: ['structural'],
    author: 'A',
    date: 'D',
  };
}

describe('B3 — fuzzyAnchorPdf', () => {
  it('re-pairs near-identical unanchored del+ins into one badged replace', () => {
    const original = [rec(0, 'The Seller shall deliver the Goods within thirty days of Completion.')];
    const revised = [rec(0, 'The Seller shall deliver the Goods within sixty days of Completion.')];
    const ops: DiffOp[] = [delOp('d', 0), insOp('i', 0)];

    const anchored = fuzzyAnchorPdf(ops, original, revised);
    const replaces = anchored.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(replaces[0]!.fuzzyAnchored).toBe(true);
    // Never suppresses: the insert is merged in, not orphaned; range spans both.
    expect(anchored.filter((o) => o.type === 'insert' || o.type === 'delete')).toHaveLength(0);
    expect(replaces[0]!.origRange).not.toBeNull();
    expect(replaces[0]!.revRange).not.toBeNull();
  });

  it('leaves genuinely dissimilar del+ins untouched (never guesses)', () => {
    const original = [rec(0, 'Governing law and jurisdiction provisions of the Deed.')];
    const revised = [rec(0, 'Schedule of fixtures and fittings at the Property.')];
    const ops: DiffOp[] = [delOp('d', 0), insOp('i', 0)];
    const anchored = fuzzyAnchorPdf(ops, original, revised);
    expect(anchored.some((o) => o.type === 'delete')).toBe(true);
    expect(anchored.some((o) => o.type === 'insert')).toBe(true);
    expect(anchored.some((o) => o.type === 'replace')).toBe(false);
  });

  it('runs inside runPipelinePdf — a moved+edited paragraph anchors as a replace', async () => {
    // [X, A1, Y] vs [X, Y, A2]: X/Y anchor, A1/A2 land in different segments →
    // structuralDiff emits delete(A1)+insert(A2); fuzzy (pre-move) re-pairs them.
    const original = [
      rec(0, 'Clause heading alpha that anchors the start.'),
      rec(1, 'The Seller shall deliver the Goods within thirty days of Completion.'),
      rec(2, 'Clause heading omega that anchors the end.'),
    ];
    const revised = [
      rec(0, 'Clause heading alpha that anchors the start.'),
      rec(1, 'Clause heading omega that anchors the end.'),
      rec(2, 'The Seller shall deliver the Goods within sixty days of Completion.'),
    ];
    const result = await runPipelinePdf({
      originalParas: original,
      revisedParas: revised,
      config: CONFIG,
      author: 'A',
      date: 'D',
      fuzzyAnchor: fuzzyAnchorPdf,
    });
    const fuzzy = result.ops.filter((o) => o.fuzzyAnchored);
    expect(fuzzy).toHaveLength(1);
    expect(fuzzy[0]!.type).toBe('replace');
    expect(fuzzy[0]!.wordDiffs?.length ?? 0).toBeGreaterThan(0); // the real change stays visible
    // Nothing suppressed: no orphaned delete/insert for that paragraph.
    expect(result.ops.filter((o) => o.type === 'delete' || o.type === 'insert')).toHaveLength(0);
  });
});
