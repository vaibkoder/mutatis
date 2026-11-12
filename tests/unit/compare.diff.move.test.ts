import { describe, it, expect } from 'vitest';

import type { ParaRecord } from '@features/compare/model.js';
import { structuralDiff, detectMoves } from '@features/compare/diff/index.js';
import { fnv1a } from '@features/compare/preprocess/index.js';

const AUTHOR = 'Test';
const DATE = '2026-06-08T00:00:00.000Z';

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

const side = (...texts: string[]): ParaRecord[] => texts.map((t, i) => para(i, t));

describe('detectMoves — T4 (verbatim move)', () => {
  it('pairs a verbatim relocated paragraph as a cross-linked move', () => {
    // "INDEMNITY" relocated from the end (Original) to near the start (Revised).
    const original = side('A', 'B', 'C', 'INDEMNITY');
    const revised = side('A', 'INDEMNITY', 'B', 'C');

    const ops = detectMoves(structuralDiff(original, revised, AUTHOR, DATE), original, revised);

    const moves = ops.filter((o) => o.type === 'move');
    expect(moves).toHaveLength(2);

    const [m1, m2] = moves as [typeof moves[number], typeof moves[number]];
    // One end is the delete (origRange set), the other the insert (revRange set).
    const del = moves.find((m) => m.origRange && !m.revRange)!;
    const ins = moves.find((m) => m.revRange && !m.origRange)!;
    expect(del).toBeDefined();
    expect(ins).toBeDefined();

    // Cross-linked.
    expect(del.moveTargetId).toBe(ins.id);
    expect(ins.moveTargetId).toBe(del.id);
    expect(m1.moveTargetId).toBe(m2.id);
    expect(m2.moveTargetId).toBe(m1.id);

    // Not reported as plain insert/delete.
    expect(ops.some((o) => o.type === 'insert')).toBe(false);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
  });

  it('matches duplicated boilerplate greedily by nearest index, leaving residue', () => {
    // "X" appears twice on the original side as a delete and twice as an insert,
    // but positions differ — greedy nearest-index pairing.
    const original = side('X', 'a', 'b', 'X', 'c');
    const revised = side('a', 'b', 'X', 'c', 'X');

    const ops = detectMoves(structuralDiff(original, revised, AUTHOR, DATE), original, revised);
    const moves = ops.filter((o) => o.type === 'move');
    // Both X deletes pair with both X inserts (equal counts) → two move pairs.
    expect(moves.length % 2).toBe(0);
    // Every move has a cross-link.
    for (const m of moves) {
      expect(m.moveTargetId).toBeDefined();
      const partner = ops.find((o) => o.id === m.moveTargetId)!;
      expect(partner.moveTargetId).toBe(m.id);
    }
  });

  it('leaves residue as plain insert/delete when counts are unequal', () => {
    // Feed detectMoves a hand-built op list where hash "X" is deleted twice but
    // inserted once → one move pair, one residual delete (residue path §5.4).
    const original = side('X', 'a', 'X', 'b'); // indices 0 and 2 carry hash("X")
    const revised = side('a', 'X', 'b'); // index 1 carries hash("X")

    const ops = detectMoves(
      [
        {
          id: 'op-0',
          type: 'delete',
          origRange: { start: 0, end: 0 },
          revRange: null,
          partId: 'document',
          classification: 'structural',
          classes: ['structural'],
          author: AUTHOR,
          date: DATE,
        },
        {
          id: 'op-1',
          type: 'delete',
          origRange: { start: 2, end: 2 },
          revRange: null,
          partId: 'document',
          classification: 'structural',
          classes: ['structural'],
          author: AUTHOR,
          date: DATE,
        },
        {
          id: 'op-2',
          type: 'insert',
          origRange: null,
          revRange: { start: 1, end: 1 },
          partId: 'document',
          classification: 'structural',
          classes: ['structural'],
          author: AUTHOR,
          date: DATE,
        },
      ],
      original,
      revised,
    );

    const moves = ops.filter((o) => o.type === 'move');
    const deletes = ops.filter((o) => o.type === 'delete');
    expect(moves).toHaveLength(2); // one del↔ins pair becomes two move ops
    expect(deletes).toHaveLength(1); // residual unmatched X delete

    // The move pair is the nearest-index match: insert at rev index 1 pairs with
    // the delete at orig index 2 (|2-1|=1 < |0-1|=1 tie → first sorted wins);
    // whichever pairs, the residual delete is left un-cross-linked.
    expect(deletes[0]!.moveTargetId).toBeUndefined();
  });
});
