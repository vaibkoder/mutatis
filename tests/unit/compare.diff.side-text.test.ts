/**
 * Feature 08 — Compare §5.3 · the yardstick itself (ADR-0008).
 *
 * `sideText` is what the never-suppress self-check in every length-preserving
 * word-diff pass measures. Those guards compare before-and-after using the same
 * function, so a SYMMETRICALLY wrong `sideText` is invisible to all three — the
 * guard cannot test its own yardstick. These are therefore direct assertions on
 * the function, not an ablation through `refineReplaceBlocks`,
 * `slideEditBoundaries` or `detectTransposition`.
 *
 * Measured, not assumed. Two corruptions were injected and the FULL suite run:
 *   - `side` flipped, so each call returns the other side. Every guard still
 *     passes (before and after are the same wrong text), and the whole suite fails
 *     in exactly ONE file: this one. (Measured when this test was written, at
 *     2433 tests / 208 files; the suite has grown since, and the figure is dated
 *     rather than restated — what matters is the ONE file, not the total.)
 *   - the `eq` filter dropped. This one the guards DO catch, because the passes
 *     move text between `eq` and the sided ops, so the sided-only reconstruction
 *     really does change across a legal re-partition. Kept as a case anyway: it
 *     is cheap and it pins the yardstick's shape.
 */
import { describe, it, expect } from 'vitest';

import { sideText } from '@features/compare/diff/index.js';
import type { WordDiff } from '@features/compare/model.js';

const d = (op: WordDiff['op'], text: string, atomic?: WordDiff['atomic']): WordDiff =>
  atomic === undefined ? { op, text } : { op, text, atomic };

describe('sideText — the never-suppress yardstick', () => {
  it('reconstructs the ORIGINAL as eq + del, in stream order', () => {
    const diffs = [
      d('eq', 'The Vendor shall '),
      d('del', 'indemnify'),
      d('ins', 'defend'),
      d('eq', ' the Purchaser.'),
    ];
    expect(sideText(diffs, 'del')).toBe('The Vendor shall indemnify the Purchaser.');
  });

  it('reconstructs the REVISED as eq + ins, in stream order', () => {
    const diffs = [
      d('eq', 'The Vendor shall '),
      d('del', 'indemnify'),
      d('ins', 'defend'),
      d('eq', ' the Purchaser.'),
    ];
    expect(sideText(diffs, 'ins')).toBe('The Vendor shall defend the Purchaser.');
  });

  it('INCLUDES the equalities', () => {
    // Unlike the side flip below, the three guards do catch this one — a legal
    // re-partition moves text between `eq` and the sided ops, so the sided-only
    // reconstruction changes. Pinned here anyway, where the failure is legible.
    const diffs = [d('eq', 'A'), d('del', 'B'), d('ins', 'C'), d('eq', 'D')];
    expect(sideText(diffs, 'del')).toBe('ABD');
    expect(sideText(diffs, 'ins')).toBe('ACD');
  });

  it('EXCLUDES the other side — the corruption no guard can see', () => {
    // Flip `side` inside `sideText` and all three self-checks still pass: they
    // compare this function against itself, so a symmetric error is invisible to
    // every one of them. Measured — under that flip the full suite fails only in
    // this file.
    const diffs = [d('del', 'old'), d('ins', 'new')];
    expect(sideText(diffs, 'del')).toBe('old');
    expect(sideText(diffs, 'ins')).toBe('new');
    expect(sideText(diffs, 'del')).not.toContain('new');
    expect(sideText(diffs, 'ins')).not.toContain('old');
  });

  it('preserves STREAM ORDER — it is a concatenation, not a set', () => {
    // A del run placed after the ins run beside it must still reconstruct the
    // original in reading order. This is what makes the mark-order pass safe.
    const delFirst = [d('eq', 'x '), d('del', 'a'), d('ins', 'b'), d('eq', ' y')];
    const insFirst = [d('eq', 'x '), d('ins', 'b'), d('del', 'a'), d('eq', ' y')];
    expect(sideText(delFirst, 'del')).toBe(sideText(insFirst, 'del'));
    expect(sideText(delFirst, 'ins')).toBe(sideText(insFirst, 'ins'));
    expect(sideText(insFirst, 'del')).toBe('x a y');
    expect(sideText(insFirst, 'ins')).toBe('x b y');
  });

  it('does not de-duplicate repeated text', () => {
    const diffs = [d('eq', 'and '), d('eq', 'and '), d('del', 'and')];
    expect(sideText(diffs, 'del')).toBe('and and and');
  });

  it('reads atomic entries exactly like ordinary ones', () => {
    const diffs = [
      d('eq', 'up to '),
      d('del', 'S$1,250,000', 'amount'),
      d('ins', 'S$2,500,000', 'amount'),
      d('eq', ' in aggregate.'),
    ];
    expect(sideText(diffs, 'del')).toBe('up to S$1,250,000 in aggregate.');
    expect(sideText(diffs, 'ins')).toBe('up to S$2,500,000 in aggregate.');
  });

  it('returns the empty string for an empty diff', () => {
    expect(sideText([], 'del')).toBe('');
    expect(sideText([], 'ins')).toBe('');
  });
});
