/**
 * The yardstick every length-preserving word-diff pass measures itself against.
 * Spec: spec/features/08-compare.md §5.3.
 *
 * `sideText` is what each never-suppress self-check in this directory compares
 * before and after (ADR-0008): a pass may re-partition the two texts, but
 * reconstructing either side must give back exactly what it was handed. It lived
 * in three byte-equivalent private copies (`transpose.ts`, `refine.ts`,
 * `anchor.ts`); one home is the rule, so this is it.
 *
 * NB `merge-split.ts` also has a function called `sideText`. It is a DIFFERENT
 * function — it takes `DiffOp[]` + `ParaRecord[]` and joins `normText` — and must
 * not be folded in here.
 *
 * `coalesce` was private to `word.ts` and joins this module because the mark-order
 * pass creates new same-op adjacencies of its own and must not grow a second copy.
 *
 * Because all three guards compare before-and-after using THIS function, a
 * symmetrically wrong `sideText` is invisible to every one of them. It therefore
 * carries its own direct unit test (`tests/unit/compare.diff.side-text.test.ts`):
 * a guard cannot test its own yardstick.
 */
import type { WordDiff } from '../model.js';

/** Reconstruct one side's full text from the diff (eq + that side's own op). */
export function sideText(diffs: readonly WordDiff[], side: 'ins' | 'del'): string {
  return diffs.filter((d) => d.op === 'eq' || d.op === side).map((d) => d.text).join('');
}

/**
 * Coalesce consecutive WordDiffs with the same op into one entry where NONE are
 * atomic (reduces verbosity). Atomic tokens stay as individual entries so their
 * `atomic` kind is preserved for classification.
 */
export function coalesce(diffs: readonly WordDiff[]): WordDiff[] {
  const out: WordDiff[] = [];
  for (const d of diffs) {
    const prev = out[out.length - 1];
    const dAtomic = d.atomic !== undefined;
    const prevAtomic = prev?.atomic !== undefined;
    if (prev && !prevAtomic && !dAtomic && prev.op === d.op) {
      prev.text += d.text;
    } else {
      out.push({ ...d });
    }
  }
  return out;
}
