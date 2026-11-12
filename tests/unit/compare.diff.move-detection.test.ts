import { describe, it, expect } from 'vitest';

import type { ParaRecord, DiffOp } from '@features/compare/model.js';
import { structuralDiff, detectMoves } from '@features/compare/diff/index.js';
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

const side = (...texts: string[]): ParaRecord[] => texts.map((t, i) => para(i, t));

/** Run the producer stages exactly as the pipeline does: structural → moves. */
function diff(original: ParaRecord[], revised: ParaRecord[]): DiffOp[] {
  return detectMoves(structuralDiff(original, revised, AUTHOR, DATE), original, revised);
}

/** The text of the paragraph an op refers to (original side for del/move-from,
 *  revised side for ins/move-to/equal/replace). */
function textOf(op: DiffOp, original: ParaRecord[], revised: ParaRecord[]): string {
  if (op.revRange) return revised[op.revRange.start]!.normText;
  if (op.origRange) return original[op.origRange.start]!.normText;
  return '';
}

/**
 * Regression: a single relocated paragraph must be isolated as the move; the
 * long common run on either side of it stays equal. The pre-fix anchor chooser
 * (a naive greedy "keep increasing" filter rather than a real LIS) pinned the
 * moved paragraph as the sole anchor and redlined the entire common run — the
 * exact inversion. Spec §5.1/§9.1 (anchor matching), §5.4 (verbatim move),
 * ADR-0008 (never suppress — the change set must be *correct*, not hidden).
 */
describe('Compare move-detection — single relocated paragraph isolates as the move', () => {
  const P = (n: number): string => `Paragraph ${n} text content.`;

  it('paragraph 1 moved to the bottom: 2–10 stay equal, 1 is the move', () => {
    const original = side(...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(P));
    const revised = side(...[2, 3, 4, 5, 6, 7, 8, 9, 10, 1].map(P));

    const ops = diff(original, revised);

    // Paragraphs 2–10 are unchanged.
    const equals = ops.filter((o) => o.type === 'equal');
    const equalTexts = new Set(equals.map((o) => textOf(o, original, revised)));
    for (let n = 2; n <= 10; n++) {
      expect(equalTexts.has(P(n))).toBe(true);
    }
    expect(equals).toHaveLength(9);

    // Paragraph 1 is the single move (a move-from + move-to pair), not 2–10.
    const moves = ops.filter((o) => o.type === 'move');
    expect(moves).toHaveLength(2);
    for (const m of moves) {
      expect(textOf(m, original, revised)).toBe(P(1));
    }

    // Nothing else is redlined: no stray insert/delete/replace.
    expect(ops.some((o) => o.type === 'insert')).toBe(false);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
  });

  it('paragraph 10 moved to the top: 1–9 stay equal, 10 is the move', () => {
    const original = side(...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(P));
    const revised = side(...[10, 1, 2, 3, 4, 5, 6, 7, 8, 9].map(P));

    const ops = diff(original, revised);

    const equals = ops.filter((o) => o.type === 'equal');
    const equalTexts = new Set(equals.map((o) => textOf(o, original, revised)));
    for (let n = 1; n <= 9; n++) {
      expect(equalTexts.has(P(n))).toBe(true);
    }
    expect(equals).toHaveLength(9);

    const moves = ops.filter((o) => o.type === 'move');
    expect(moves).toHaveLength(2);
    for (const m of moves) {
      expect(textOf(m, original, revised)).toBe(P(10));
    }
    expect(ops.some((o) => o.type === 'insert')).toBe(false);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
  });

  it('a middle paragraph moved downward: the rest stay equal, it is the move', () => {
    // Move paragraph 5 to sit after paragraph 8.
    const original = side(...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(P));
    const revised = side(...[1, 2, 3, 4, 6, 7, 8, 5, 9, 10].map(P));

    const ops = diff(original, revised);

    const equals = ops.filter((o) => o.type === 'equal');
    const equalTexts = new Set(equals.map((o) => textOf(o, original, revised)));
    for (const n of [1, 2, 3, 4, 6, 7, 8, 9, 10]) {
      expect(equalTexts.has(P(n))).toBe(true);
    }
    expect(equals).toHaveLength(9);

    const moves = ops.filter((o) => o.type === 'move');
    expect(moves).toHaveLength(2);
    for (const m of moves) {
      expect(textOf(m, original, revised)).toBe(P(5));
    }
    expect(ops.some((o) => o.type === 'insert')).toBe(false);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
  });
});
