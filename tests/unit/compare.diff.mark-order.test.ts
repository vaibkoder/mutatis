/**
 * Feature 08 — Compare §5.5 · mark order (ADR-0043).
 *
 * The insertion reads before the deletion it replaces. These are direct unit
 * tests over constructed `WordDiff[]`: the pass's own contract, its refusals,
 * and the never-suppress invariant it may not break (ADR-0008).
 *
 * Measured at HEAD: reduced to `return diffs`, EIGHT of the eleven fail. The three
 * that pass are REFUSAL tests — they assert the pass leaves a shape alone, so a
 * no-op satisfies them by construction and no ablation of this pass can break
 * them. They pin the opposite corruption (a pass too greedy to respect run
 * boundaries, or an order that already reads correctly), and the "swaps WHOLE
 * runs" case is the one that fails under that. Both transcripts are in the plan's
 * Evidence.
 *
 * This comment read "six of the nine" until review measured it: that
 * reading was taken before the last two tests were added, and adding them changed
 * both figures. Re-derive a count like this when the file grows; do not re-run the
 * old sentence.
 * The SEAM barrier and the two pipeline wirings have their own files:
 * `tests/unit/compare/mark-order.seam.test.ts` and
 * `tests/unit/compare/pipeline.mark-order.test.ts`.
 */
import { describe, it, expect } from 'vitest';

import { insertBeforeDelete, sideText } from '@features/compare/diff/index.js';
import type { WordDiff } from '@features/compare/model.js';

const d = (op: WordDiff['op'], text: string, atomic?: WordDiff['atomic']): WordDiff =>
  atomic === undefined ? { op, text } : { op, text, atomic };

/** `{+ins+}[-del-]` rendering, so a failure reads like the redline it describes. */
function render(diffs: readonly WordDiff[]): string {
  return diffs
    .map((x) => (x.op === 'ins' ? `{+${x.text}+}` : x.op === 'del' ? `[-${x.text}-]` : x.text))
    .join('');
}

describe('insertBeforeDelete (§5.5, ADR-0043)', () => {
  it('puts the insertion first in a one-to-one replace', () => {
    const out = insertBeforeDelete([
      d('eq', 'a '),
      d('del', 'exclusive'),
      d('ins', 'non-exclusive'),
      d('eq', ' licence'),
    ]);
    expect(render(out)).toBe('a {+non-exclusive+}[-exclusive-] licence');
  });

  it('swaps WHOLE runs, never reordering within a side', () => {
    // Two del entries (atomic keeps them separate) followed by two ins entries.
    const out = insertBeforeDelete([
      d('eq', 'up to '),
      d('del', 'S$1,000', 'amount'),
      d('del', ' and no more'),
      d('ins', 'S$2,000', 'amount'),
      d('ins', ' in aggregate'),
      d('eq', '.'),
    ]);
    expect(out.map((x) => x.op)).toEqual(['eq', 'ins', 'ins', 'del', 'del', 'eq']);
    // Each side's internal order is untouched.
    expect(out.filter((x) => x.op === 'del').map((x) => x.text)).toEqual([
      'S$1,000',
      ' and no more',
    ]);
    expect(out.filter((x) => x.op === 'ins').map((x) => x.text)).toEqual([
      'S$2,000',
      ' in aggregate',
    ]);
  });

  it('leaves a del run with no insertion beside it untouched', () => {
    const input = [d('eq', 'indemnify '), d('del', 'and hold harmless '), d('eq', 'the Purchaser')];
    const out = insertBeforeDelete(input);
    expect(render(out)).toBe('indemnify [-and hold harmless -]the Purchaser');
    // Nothing swapped ⇒ the input itself comes back, not a rebuilt copy.
    expect(out).toBe(input);
  });

  it('leaves an insertion that already precedes its deletion untouched', () => {
    // This is `detectTransposition`'s own output shape (ADR-0031) — the
    // precedent the whole pass is being aligned to.
    const input = [
      d('eq', 'The Vendor shall '),
      d('ins', 'defend, '),
      d('eq', 'indemnify'),
      d('del', ', defend'),
      d('eq', ' and hold harmless'),
    ];
    expect(insertBeforeDelete(input)).toBe(input);
  });

  it('re-coalesces a NEW same-op adjacency the swap creates', () => {
    // `{+A+}[-B-]{+C+}` → `{+A+}{+C+}[-B-]`: an ins/ins adjacency the input
    // could not contain, because the input is already coalesced.
    const out = insertBeforeDelete([
      d('eq', 'x '),
      d('ins', 'new '),
      d('del', 'old'),
      d('ins', 'words'),
      d('eq', ' y'),
    ]);
    expect(out.map((x) => x.op)).toEqual(['eq', 'ins', 'del', 'eq']);
    expect(render(out)).toBe('x {+new words+}[-old-] y');
  });

  it('never coalesces an ATOMIC entry away — its kind must survive the swap', () => {
    // `classify` has already run, but the AI export and the side-by-side panel
    // read `atomic` downstream.
    const out = insertBeforeDelete([
      d('ins', 'by '),
      d('del', 'on'),
      d('ins', '1 January 2026', 'date'),
      d('eq', '.'),
    ]);
    // Two ins ENTRIES, not one: the atomic date keeps its own. The renderer
    // draws adjacent same-op entries as one continuous mark (ADR-0031).
    expect(render(out)).toBe('{+by +}{+1 January 2026+}[-on-].');
    const dates = out.filter((x) => x.atomic === 'date');
    expect(dates).toHaveLength(1);
    expect(dates[0]).toEqual({ op: 'ins', text: '1 January 2026', atomic: 'date' });
    // The plain ins beside it must NOT have absorbed it.
    expect(out.find((x) => x.op === 'ins' && x.atomic === undefined)?.text).toBe('by ');
  });

  it('handles several replace blocks in one paragraph independently', () => {
    const out = insertBeforeDelete([
      d('del', 'A'),
      d('ins', 'B'),
      d('eq', ' and '),
      d('del', 'C'),
      d('ins', 'D'),
    ]);
    expect(render(out)).toBe('{+B+}[-A-] and {+D+}[-C-]');
  });

  it('reconstructs both sides identically — the never-suppress invariant', () => {
    const input = [
      d('eq', 'The Purchaser shall pay '),
      d('del', 'S$1,250,000'),
      d('ins', 'S$2,500,000'),
      d('eq', ' on '),
      d('del', 'Completion'),
      d('ins', 'the Completion Date'),
      d('eq', '.'),
    ];
    const out = insertBeforeDelete(input);
    // The order really did change — otherwise this test passes on a no-op.
    expect(render(out)).toBe(
      'The Purchaser shall pay {+S$2,500,000+}[-S$1,250,000-] on ' +
        '{+the Completion Date+}[-Completion-].',
    );
    expect(sideText(out, 'del')).toBe(sideText(input, 'del'));
    expect(sideText(out, 'ins')).toBe(sideText(input, 'ins'));
    expect(sideText(out, 'del')).toBe(
      'The Purchaser shall pay S$1,250,000 on Completion.',
    );
    expect(sideText(out, 'ins')).toBe(
      'The Purchaser shall pay S$2,500,000 on the Completion Date.',
    );
  });

  it('only ever RE-ORDERS the same two texts, over a generated corpus (ADR-0008)', () => {
    // Like its four siblings, this pass ends in a reconstruct-both-sides guard
    // that returns the input unchanged on any mismatch. In correct code that
    // branch is unreachable — a whole-run swap cannot change either side — so it
    // cannot be provoked without injecting a fault, and an "ablation" of it
    // therefore breaks nothing. What IS testable is the property it protects, so
    // that is pinned here over generated streams, exactly as
    // `compare.diff.refine.test.ts` does for `refineReplaceBlocks`.
    //
    // Deterministic LCG, high bits — no Math.random.
    const OPS: Array<WordDiff['op']> = ['eq', 'del', 'ins'];
    const KINDS: Array<WordDiff['atomic'] | undefined> = [undefined, undefined, 'amount', 'date'];
    let seed = 20260917;
    const next = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return Math.floor((seed >>> 16) % n);
    };

    let swapsSeen = 0;
    for (let trial = 0; trial < 500; trial += 1) {
      const input: WordDiff[] = [];
      const len = 2 + next(8);
      for (let k = 0; k < len; k += 1) {
        const op = OPS[next(3)]!;
        const atomic = KINDS[next(4)];
        const text = `t${String(next(6))} `;
        input.push(atomic === undefined ? { op, text } : { op, text, atomic });
      }
      const before = { orig: sideText(input, 'del'), rev: sideText(input, 'ins') };
      const out = insertBeforeDelete(input.map((x) => ({ ...x })));
      expect(sideText(out, 'del')).toBe(before.orig);
      expect(sideText(out, 'ins')).toBe(before.rev);
      if (render(out) !== render(input)) swapsSeen += 1;
    }
    // Measured, so the corpus is checkable rather than asserted: the generator
    // must actually produce swappable shapes or the loop above proves nothing.
    expect(swapsSeen).toBe(199);
  });

  it('is NOT idempotent on chained replace blocks — the pass runs exactly once', () => {
    // Recorded property, not an aspiration. `[-A-]{+B+}[-C-]{+D+}` with no
    // equality between the two blocks swaps to `{+B+}[-A-]{+D+}[-C-]`, which
    // leaves `[-A-]` abutting `{+D+}`; a SECOND invocation re-windows that pair
    // and merges two independent edits into one block. Both sides still
    // reconstruct, so the never-suppress guard cannot see it.
    //
    // Measured: the shape occurs in 0 of 3,500+ word-diff streams across the
    // suite — DMP's `diff_cleanupMerge` leaves an equality between changed
    // regions. So the answer is "call it once", which the two pipeline guards in
    // `tests/unit/compare/pipeline.mark-order.test.ts` pin. If that ever has to
    // change, the fix is provenance on a swapped run, not a second pass.
    const chained = [d('del', 'A'), d('ins', 'B'), d('del', 'C'), d('ins', 'D')];
    const once = insertBeforeDelete(chained);
    expect(render(once)).toBe('{+B+}[-A-]{+D+}[-C-]');
    const twice = insertBeforeDelete(once);
    expect(render(twice)).toBe('{+BD+}[-AC-]');
    // …and even then nothing has left the redline.
    expect(sideText(twice, 'del')).toBe(sideText(chained, 'del'));
    expect(sideText(twice, 'ins')).toBe(sideText(chained, 'ins'));
  });

  it('returns short inputs unchanged', () => {
    const one = [d('del', 'x')];
    expect(insertBeforeDelete(one)).toBe(one);
    expect(insertBeforeDelete([])).toEqual([]);
  });
});
