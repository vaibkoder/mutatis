/**
 * Level-2 word diff · transposition-aware anchoring. ADR-0031.
 * Spec: spec/features/08-compare.md §5.3. Principle: docs/compare-readability-principles.md P7
 * ("A move reads as a move ... at paragraph scale and at phrase scale").
 *
 * `slideEditBoundaries` (anchor.ts) only ever touches a SINGLE-sided edit
 * flanked by `eq` on both sides — by design, "a two-sided replace has no free
 * boundary to slide". An adjacent-word transposition ("indemnify, defend" →
 * "defend, indemnify") is exactly a two-sided replace TWICE over: DMP's Myers
 * diff, maximising the retained common subsequence, matches the separator
 * between the two words as `eq` and reports the transposition as two
 * independent one-word substitutions:
 *
 *   shall {+defend+}[-indemnify-], {+indemnify+}[-defend-] and hold …
 *
 * — "diff salad": the reader has to notice, unaided, that the two struck words
 * are exactly the two inserted words in the other order. This is the classic
 * LCS/Myers-diff blind spot for adjacent transpositions (the same failure mode
 * `git diff` and `difflib` show); it is not a boundary-placement problem, so no
 * amount of re-scoring `slideEditBoundaries`'s existing triple fixes it.
 *
 * `detectTransposition` recognises the shape directly — `del(A) ins(B) eq(mid)
 * del(B) ins(A)`, i.e. the same two spans changing places around a fixed
 * middle — and re-expresses it as ONE relocation instead of two replacements:
 *
 *   shall {+defend, +}indemnify[-, defend-] and hold …
 *
 * one span (B) moves forward across `mid`; the other (A) is left as the sole
 * survivor and is never marked at all. This is still an ordinary insert-here /
 * delete-there pair: P7 "at phrase scale" does not require Word's native
 * `w:moveFrom`/`w:moveTo`, which carries a range-start/range-end pair and a
 * shared `w:name` and buys nothing a reader can see inside one sentence.
 *
 * Guarded narrowly, on structure alone (no vocabulary, no semantics):
 *  - both swapped spans must be short (`MAX_SWAP_TOKENS`) — this is a WORD or
 *    short-phrase transposition fix; a long-phrase swap already reads fine as
 *    two clean phrase-for-phrase replacements and is left alone (measured in
 *    the probe script — rewriting it did not read better).
 *  - the connective between them must be short (`MAX_MID_CHARS`) — bounds how
 *    far apart the two spans can be and still read as "these traded places"
 *    rather than two coincidentally-matching, unrelated edits.
 *  - the connective must carry NO WORD — only whitespace and punctuation. The
 *    rewrite duplicates it (inserted with B, struck with B's old copy), and a
 *    WORD struck in one mark and re-inserted in another is exactly the P2
 *    failure `refine.ts` refuses to commit (ADR-0030 rule 2). `, ` is invisible
 *    duplication; ` and the ` would make the reader prove two marks identical,
 *    so `A and the B` → `B and the A` is deliberately left as two clean
 *    one-span replacements.
 *  - the quintuple must cover WHOLE runs — the entry before it may not be a
 *    `del` and the entry after may not be an `ins`, or the window would be a
 *    slice through the middle of a longer span (`coalesce` leaves adjacent
 *    same-op entries in place wherever one of them is atomic), and "A moved"
 *    would be claimed for only part of what actually moved.
 *  - the two spans must differ (a same-text "swap" is not a change at all, and
 *    cannot arise from a real diff, but is guarded defensively).
 *
 * Length-preserving and self-verifying, per ADR-0008 (never suppress): it
 * reconstructs both sides before and after and returns the input unchanged on
 * any mismatch, mirroring `slideEditBoundaries`'s guard exactly.
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { WordDiff } from '../model.js';

import { tokenize } from './atomic-tokens.js';
import { sideText } from './side-text.js';

/** Each swapped span: at most this many non-whitespace reader tokens. */
const MAX_SWAP_TOKENS = 4;
/**
 * The connective between the two spans: at most this many characters. It carries
 * no word (see `isNoise`), so its cost to the reader is measured in the distance
 * it puts between the two spans, not in tokens — `, ` is nothing; the dotted
 * leader of a schedule or signature block is too far to read as one swap.
 */
const MAX_MID_CHARS = 8;

/** True when the text carries no letter or digit — whitespace or punctuation. */
function isNoise(text: string): boolean {
  return !/[\p{L}\p{N}]/u.test(text);
}

/** Count of `tokenize` tokens in `text` that are not pure whitespace. */
function nonSpaceTokenCount(text: string, config: CompareConfig): number {
  let n = 0;
  for (const tok of tokenize(text, config)) if (!/^\s+$/.test(tok)) n++;
  return n;
}

/**
 * Rewrite every `del(A) ins(B) eq(mid) del(B) ins(A)` quintuple in `diffs`
 * (A/B verbatim-equal on both occurrences) into `ins(B+mid) eq(A) del(mid+B)`.
 * Runs on the already-coalesced WordDiff[] `wordDiff` produces (so a "span" is
 * whatever single blob DMP + coalesce already decided was one del/ins run —
 * this pass only ever re-labels existing spans, it never re-splits one).
 */
export function detectTransposition(diffs: WordDiff[], config: CompareConfig): WordDiff[] {
  if (diffs.length < 5) return diffs;

  const before = { orig: sideText(diffs, 'del'), rev: sideText(diffs, 'ins') };

  const out: WordDiff[] = [];
  let i = 0;
  while (i < diffs.length) {
    const quint = i + 4 < diffs.length ? diffs.slice(i, i + 5) : undefined;
    if (quint && wholeRuns(diffs, i) && isTransposition(quint, config)) {
      const [a1, b1, mid, b2] = quint as [WordDiff, WordDiff, WordDiff, WordDiff, WordDiff];
      // Each span keeps the atomic kind it arrived with, so a swap of two
      // amounts still classifies as a numerical-value change (§5.4, P10). The
      // connective travels as its own entry rather than being glued on: two
      // adjacent same-op entries render as one mark either way.
      out.push(withKind('ins', b1), { op: 'ins', text: mid.text });
      out.push(withKind('eq', a1));
      out.push({ op: 'del', text: mid.text }, withKind('del', b2));
      i += 5;
      continue;
    }
    out.push(diffs[i]!);
    i += 1;
  }

  // Never-suppress guard (ADR-0008): reconstruct both sides; bail to the
  // untouched diff if the rewrite ever changed what either side says.
  if (sideText(out, 'del') !== before.orig || sideText(out, 'ins') !== before.rev) return diffs;
  return out;
}

/** Re-label one entry with a new op, carrying its atomic kind across. */
function withKind(op: WordDiff['op'], from: WordDiff): WordDiff {
  return from.atomic === undefined ? { op, text: from.text } : { op, text: from.text, atomic: from.atomic };
}

/**
 * Is the five-entry window at `i` made of WHOLE del/ins runs? `a1` is the start
 * of its deletion only if what precedes it is not another `del`, and `a2` is the
 * end of its insertion only if what follows is not another `ins`; the three
 * middle entries are bounded by the window itself. Without this the pass could
 * fire on a slice through a longer span and report a partial move.
 */
function wholeRuns(diffs: readonly WordDiff[], i: number): boolean {
  return diffs[i - 1]?.op !== 'del' && diffs[i + 5]?.op !== 'ins';
}

function isTransposition(quint: readonly WordDiff[], config: CompareConfig): boolean {
  const [a1, b1, mid, b2, a2] = quint as [WordDiff, WordDiff, WordDiff, WordDiff, WordDiff];
  return (
    a1.op === 'del' &&
    b1.op === 'ins' &&
    mid.op === 'eq' &&
    b2.op === 'del' &&
    a2.op === 'ins' &&
    a1.text.length > 0 &&
    b1.text.length > 0 &&
    mid.text.length > 0 &&
    a1.text === a2.text &&
    b1.text === b2.text &&
    a1.text !== b1.text &&
    isNoise(mid.text) &&
    nonSpaceTokenCount(a1.text, config) <= MAX_SWAP_TOKENS &&
    nonSpaceTokenCount(b1.text, config) <= MAX_SWAP_TOKENS &&
    mid.text.length <= MAX_MID_CHARS
  );
}

