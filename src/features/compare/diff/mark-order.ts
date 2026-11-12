/**
 * Mark order inside a changed paragraph: the insertion reads before the deletion.
 * Spec: spec/features/08-compare.md §5.5 · ADR-0043.
 *
 * diff-match-patch's `diff_cleanupMerge` normalises every changed region to
 * DELETE-then-INSERT, and nothing between `diff_main` and authoring reordered the
 * two sides — so the shipped redline read `[-exclusive-]{+non-exclusive+}`. That
 * order was an inherited library default, never a decision. The reader wants the
 * text that SURVIVES first: `{+non-exclusive+}[-exclusive-]`, which is what this
 * codebase's own `detectTransposition` already emits (ADR-0031).
 *
 * An earlier header claimed this is what commercial comparison tools do. That
 * claim was not substantiated, so this pass is a deliberate choice rather than a
 * market convention. ADR-0043 carries the evidence and the caveats.
 *
 * The pass swaps each maximal `del` run with the maximal `ins` run immediately
 * following it. It is the fifth length-preserving word-diff pass (ADR-0028/0030/
 * 0031/0034) and follows their shape exactly: whole runs only, never a reorder
 * WITHIN a side, self-verifying against `sideText`, and returning its input
 * unchanged on any mismatch (ADR-0008 — no change may leave the redline).
 *
 * **It must run exactly once per op.** The pass is NOT idempotent on a stream
 * holding two changed regions with no equality between them
 * (`[-A-]{+B+}[-C-]{+D+}`): the first swap leaves `[-A-]` abutting `{+D+}`, and a
 * second invocation would re-window that pair and merge two independent edits
 * into one `{+BD+}[-AC-]` block. Both sides still reconstruct, so the
 * never-suppress guard below would NOT catch it — it is a readability
 * regression, not a suppression. Measured across the suite: **4,535** streams leave
 * `wordDiff`, **327** of them reach this pass, and **0** carry the chained shape
 * (DMP's `diff_cleanupMerge` leaves an equality between changed regions and none of
 * the four upstream passes chain them), so
 * this is a property to respect at the two call sites, not a defect to design
 * around. `tests/unit/compare.diff.mark-order.test.ts` pins it.
 *
 * Runs after classification, on the `WordDiff` stream — not inside `word.ts`
 * (the classifiers would eat the new order for nothing) and not inside
 * `author.ts` (the side-by-side panel and the exports would then disagree with
 * the redline). After Stage 6 every consumer sees one order.
 */
import type { WordDiff } from '../model.js';

import { PARA_SEAM } from './merge-split.js';
import { sideText, coalesce } from './side-text.js';

/**
 * Re-order a word-diff stream so each insertion precedes the deletion it
 * replaces.
 *
 * @param diffs an ordered `WordDiff[]`, already coalesced (as `wordDiff` returns it)
 * @returns the re-ordered, re-coalesced stream — or `diffs` unchanged if either
 *          side failed to reconstruct
 */
export function insertBeforeDelete(diffs: WordDiff[]): WordDiff[] {
  if (diffs.length < 2) return diffs;

  const before = { orig: sideText(diffs, 'del'), rev: sideText(diffs, 'ins') };

  const out: WordDiff[] = [];
  let i = 0;
  let swapped = false;

  while (i < diffs.length) {
    if (diffs[i]!.op !== 'del') {
      out.push(diffs[i]!);
      i += 1;
      continue;
    }

    // The MAXIMAL del run starting here, then the maximal ins run abutting it.
    // Whole runs only: a swap of a slice would interleave one side with itself.
    let delEnd = i;
    while (delEnd < diffs.length && diffs[delEnd]!.op === 'del') delEnd += 1;
    let insEnd = delEnd;
    while (insEnd < diffs.length && diffs[insEnd]!.op === 'ins') insEnd += 1;

    // A deletion with nothing inserted beside it already reads correctly.
    if (insEnd === delEnd) {
      for (let t = i; t < delEnd; t += 1) out.push(diffs[t]!);
      i = delEnd;
      continue;
    }

    // ── the seam barrier ────────────────────────────────────────────────
    // A `PARA_SEAM` in this window is a PARAGRAPH MARK, not a word.
    // `splitAtSeams` (`output/author.ts`) cuts the stream at each seam and makes
    // the seam-bearing entry's op the revision on that mark, so a run that
    // crosses the seam in the swap crosses a paragraph boundary: the runs either
    // side land in different `<w:p>` elements than ADR-0032 authored them into.
    //
    // Measured, so the barrier's width is a judgement against a cost and not a
    // reflex: the mark-revision swap — a merge rendered as a split — needs seams
    // on BOTH sides of the window, and `detectMergesAndSplits → tryPair` only
    // ever produces 2-original→1-revised or 1→2, never 2→2, so an op can carry
    // at most ONE seam and that corruption is unreachable. What IS reachable,
    // and what the barrier prevents, is text moving between the emitted
    // paragraphs while accept-all and reject-all still round-trip — a silent
    // rewrite of ADR-0032's output for a presentation preference.
    //
    // Merge/split is also where the order matters least: the two marks sit on
    // either side of a paragraph mark, not side by side in one sentence.
    let seam = false;
    for (let t = i; t < insEnd && !seam; t += 1) {
      if (diffs[t]!.text.includes(PARA_SEAM)) seam = true;
    }
    if (seam) {
      for (let t = i; t < insEnd; t += 1) out.push(diffs[t]!);
      i = insEnd;
      continue;
    }

    for (let t = delEnd; t < insEnd; t += 1) out.push(diffs[t]!);
    for (let t = i; t < delEnd; t += 1) out.push(diffs[t]!);
    swapped = true;
    i = insEnd;
  }

  if (!swapped) return diffs;

  // A swap can put the moved ins run next to an ins run that preceded the del
  // (`{+A+}[-B-]{+C+}` → `{+A+}{+C+}[-B-]`), an adjacency the input could not
  // contain. `coalesce` never merges an atomic entry, so an amount, date or
  // clause reference keeps its own entry and its `atomic` kind — which the AI
  // export and the side-by-side panel read long after `classify` has run.
  const result = coalesce(out);

  // Never-suppress guard (ADR-0008): the pass may only re-order whole runs, so
  // each side must reconstruct exactly. If that ever failed, keep the input.
  if (sideText(result, 'del') !== before.orig || sideText(result, 'ins') !== before.rev) {
    return diffs;
  }
  return result;
}
