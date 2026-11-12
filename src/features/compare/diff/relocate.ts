/**
 * Level-2 word diff · relocation-preserving semantic cleanup.
 * Spec: spec/features/08-compare.md §5.3. Principle:
 * docs/compare-readability-principles.md P1, P2, P3, P7.
 *
 * A slot relocation — one span struck where it used to sit and re-inserted a
 * few words later, with the text in between untouched — is the single most
 * common shape a redline of a legal clause has to render, and `diff_main`
 * already gets it right:
 *
 *   The Purchaser [-shall -]notify{+ shall+} each Indemnified Person …
 *
 * `diff_cleanupSemantic` then destroys it. Its rule is "eliminate an equality
 * smaller than or equal to the edits on either side of it", measured with
 * `String.length` — and `diff/word.ts` runs it in SENTINEL space, where one
 * character is one TOKEN. The equality `notify` counts 1 against edits of 2
 * (`shall` + the space), so the equality goes, both edits grow by what they
 * swallowed, and the reader is handed:
 *
 *   The Purchaser {+notify shall+}[-shall notify-] each Indemnified Person …
 *
 * — `notify` struck and re-inserted for nothing, which is P1 ("never widen a
 * mark to swallow text that is word-for-word identical on both sides") and P2
 * ("never make the reader diff the redline") in one move.
 *
 * **Nothing downstream can undo it, by design.** `refineReplaceBlocks`
 * (ADR-0030) re-diffs each fused block, but its second acceptance rule refuses
 * any candidate that strikes a word in one mark and re-inserts it in another —
 * and every finer rendering of a relocation does exactly that. ADR-0030 states
 * the consequence outright: "a transposition stays coarse and is left to move
 * detection". `detectTransposition` (ADR-0031) is that move detection, and it
 * matches one shape, `del(A) ins(B) eq(mid) del(B) ins(A)`, in the COALESCED
 * diff — a shape the fusion has already erased. Measured over 3,675 generated
 * clause pairs, removing `detectTransposition` from the chain changes the sweep
 * by nothing at all.
 *
 * So the relocation is only visible in one place: BEFORE the cleanup. This pass
 * runs the cleanup everywhere except across a relocation — the diff is cut at
 * each protected triple and `diff_cleanupSemantic` is applied to each remaining
 * stretch on its own, so every other over-fusion is still cleaned up exactly as
 * before while the relocation survives with its equality intact.
 *
 * Guarded on STRUCTURE alone — no vocabulary, no semantics, no
 * equivalence scoring (ADR-0008):
 *
 *  - the window is a `del` run, ONE `eq` run and an `ins` run (either order):
 *    nothing else changed between the two ends, so "this moved across that" is
 *    the whole story;
 *  - the two runs are VERBATIM identical — the same tokens in the same order,
 *    ignoring only pure-whitespace tokens, compared as sentinels so the
 *    comparison includes each token's atomic kind and a reference value can
 *    never match an identical-looking ordinary word (ADR-0029). This is exact
 *    sequence equality, never a similarity score;
 *  - the relocated span carries readable content — a letter or a digit. A
 *    stray comma "relocating" is not something a reader can see move, and a
 *    span with nothing in it is the contentless-span artefact the readability
 *    metric had to be corrected for.
 *
 * The equality between the two ends is non-empty by construction: `runsOf`
 * drops empty entries, so a `del` and an `ins` with nothing between them are
 * adjacent runs and never match the window at all. That case is a replace
 * block, which `refineReplaceBlocks` owns.
 *
 * A run may take part in more than one window, and is deliberately not
 * de-duplicated. Protection only decides where the cleanup is applied, so
 * overlapping windows cost nothing — and a first-match-wins rule was measured
 * to make one rendering WORSE, leaving a second relocation's `del` unprotected
 * and fused with the edit after it. It moved nothing on the 3,675-pair sweep in
 * either direction, so it went rather than staying as untested weight.
 *
 * There is deliberately NO bound on how far the span moved. The rule DMP
 * applies is self-limiting: it only eliminates an equality no longer than the
 * edits flanking it, so where the untouched text between the two ends is long,
 * the fusion never happens and this pass has nothing to protect. A bound would
 * be a tuning knob with no failure mode behind it.
 *
 * Never suppresses (ADR-0008): the pass only decides which
 * stretches of one diff are handed to `diff_cleanupSemantic`, and the cleanup
 * is itself length-preserving. Both sides are reconstructed before and after,
 * and on any mismatch the ordinary whole-diff cleanup is returned instead.
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */
import DiffMatchPatch from 'diff-match-patch';

import { hasReadableContent } from './atomic-tokens.js';
import type { Token } from './atomic-tokens.js';

/** One diff-match-patch entry: `[op, sentinels]`, one sentinel char per token. */
export type DmpDiff = [number, string];

/** Resolve a sentinel char back to the token it stands for. */
export type TokenAt = (ch: string) => Token;

const DIFF_DELETE = -1;
const DIFF_INSERT = 1;
const DIFF_EQUAL = 0;

/** Collapse to maximal same-op runs, dropping empties. */
function runsOf(diffs: readonly DmpDiff[]): DmpDiff[] {
  const out: DmpDiff[] = [];
  for (const [op, sentinels] of diffs) {
    if (sentinels === '') continue;
    const prev = out[out.length - 1];
    if (prev && prev[0] === op) prev[1] += sentinels;
    else out.push([op, sentinels]);
  }
  return out;
}

/** The run's sentinels with pure-whitespace tokens removed. */
function contentSentinels(sentinels: string, tokenAt: TokenAt): string {
  let out = '';
  for (const ch of sentinels) if (!/^\s+$/.test(tokenAt(ch).text)) out += ch;
  return out;
}

/** Does this span carry anything a reader could see move — a letter or digit? */
function spanIsReadable(sentinels: string, tokenAt: TokenAt): boolean {
  for (const ch of sentinels) if (hasReadableContent(tokenAt(ch).text)) return true;
  return false;
}

/** Reconstruct one side of a diff, as sentinels (eq + that side's own op). */
function sideSentinels(diffs: readonly DmpDiff[], side: number): string {
  let out = '';
  for (const [op, sentinels] of diffs) if (op === DIFF_EQUAL || op === side) out += sentinels;
  return out;
}

/** Run DMP's own semantic cleanup over a whole diff (mutating a fresh copy). */
function cleanupWhole(diffs: readonly DmpDiff[]): DmpDiff[] {
  const copy: DmpDiff[] = diffs.map(([op, s]) => [op, s]);
  new DiffMatchPatch().diff_cleanupSemantic(copy);
  return copy;
}

/**
 * Indices of the runs that form a relocation window — a `del`/`ins` pair whose
 * spans are verbatim identical, separated by exactly one equality.
 */
function protectedRunIndices(runs: readonly DmpDiff[], tokenAt: TokenAt): Set<number> {
  const guarded = new Set<number>();
  for (let i = 0; i + 2 < runs.length; i++) {
    const [aOp, aSent] = runs[i]!;
    const [midOp] = runs[i + 1]!;
    const [bOp, bSent] = runs[i + 2]!;
    const isPair =
      (aOp === DIFF_DELETE && bOp === DIFF_INSERT) || (aOp === DIFF_INSERT && bOp === DIFF_DELETE);
    if (!isPair || midOp !== DIFF_EQUAL) continue;
    const aCore = contentSentinels(aSent, tokenAt);
    // Exact sequence equality, on SENTINELS — so the comparison carries each
    // token's atomic kind and a reference value can never match an
    // identical-looking ordinary word (ADR-0029). Never a similarity score.
    if (aCore !== contentSentinels(bSent, tokenAt)) continue;
    // Subsumes an empty span: `spanIsReadable('')` is false.
    if (!spanIsReadable(aCore, tokenAt)) continue;
    guarded.add(i);
    guarded.add(i + 1);
    guarded.add(i + 2);
  }
  return guarded;
}

/**
 * `diff_cleanupSemantic`, applied everywhere except across a relocation.
 *
 * @param diffs   the raw `diff_main` output, in sentinel space.
 * @param tokenAt resolves a sentinel char to its token (for whitespace and
 *                readable-content tests only — identity is compared on the
 *                sentinels themselves, so atomic kind is part of it).
 * @returns a cleaned diff; identical to `diff_cleanupSemantic`'s own output
 *          whenever no relocation is present.
 */
export function cleanupPreservingRelocations(diffs: DmpDiff[], tokenAt: TokenAt): DmpDiff[] {
  const runs = runsOf(diffs);
  const guarded = protectedRunIndices(runs, tokenAt);
  if (guarded.size === 0) return cleanupWhole(diffs);

  const out: DmpDiff[] = [];
  let pending: DmpDiff[] = [];
  const flush = (): void => {
    if (pending.length > 0) out.push(...cleanupWhole(pending));
    pending = [];
  };
  for (const [i, run] of runs.entries()) {
    if (guarded.has(i)) {
      flush();
      out.push(run);
    } else {
      pending.push(run);
    }
  }
  flush();

  // Never-suppress guard (ADR-0008): partitioning may only change WHERE the
  // cleanup is applied, never what either side says. Fall back to the ordinary
  // whole-diff cleanup if that ever failed.
  if (
    sideSentinels(out, DIFF_DELETE) !== sideSentinels(diffs, DIFF_DELETE) ||
    sideSentinels(out, DIFF_INSERT) !== sideSentinels(diffs, DIFF_INSERT)
  ) {
    return cleanupWhole(diffs);
  }
  return out;
}
