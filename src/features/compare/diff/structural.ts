/**
 * Level-1 structural / paragraph diff.
 * Spec: spec/features/08-compare.md §5.1, §9.1.
 *
 * Pipeline:
 *   1. Build hash→index maps for both sides.
 *   2. Anchor matching: hashes UNIQUE on both sides AND equal → anchor pairs;
 *      keep a monotonically-increasing chain (reordering-safe, §9.1).
 *   3. Segmented LCS between consecutive anchors (bounds cost; one insert does
 *      not cascade misalignment onto every following paragraph).
 *   4. Pair each gap between matches — by similarity when the caller passes a
 *      `GapPairing` (ADR-0039), otherwise by position.
 *   5. Emit ordered DiffOp[]: equal / replace / insert / delete.
 *
 * Pure: author/date are injected; no Date.now().
 */
import type { ParaRecord, DiffOp, ChangeClass } from '../model.js';

import type { GapPairing } from './gap-pairing.js';
import { computeLCS } from './lcs.js';

/**
 * Produce the ordered Level-1 DiffOp list.
 *
 * @param original  ParaRecord[] for the Original (base) side, in body order.
 * @param revised   ParaRecord[] for the Revised (compared) side, in body order.
 * @param author    injected w:author for each op.
 * @param date      injected ISO-8601 date for each op.
 * @param pairing   how to pair a gap's leftovers (ADR-0039). Omitted ⇒ by
 *                  position, which is what the PDF path and the parts diff use.
 */
export function structuralDiff(
  original: ParaRecord[],
  revised: ParaRecord[],
  author: string,
  date: string,
  pairing?: GapPairing,
): DiffOp[] {
  const anchors = findAnchors(original, revised);

  const ops: DiffOp[] = [];
  const counter = { next: 0 };

  // Walk anchor-bounded segments. Sentinels at both ends frame the first/last gaps.
  let prevO = 0;
  let prevR = 0;

  const emitSegment = (oStart: number, oEnd: number, rStart: number, rEnd: number): void => {
    diffSegment(
      original,
      revised,
      oStart,
      oEnd,
      rStart,
      rEnd,
      author,
      date,
      counter,
      ops,
      pairing,
    );
  };

  for (const [oIdx, rIdx] of anchors) {
    // Diff the gap strictly before this anchor.
    emitSegment(prevO, oIdx, prevR, rIdx);
    // The anchor itself is an equal pair (unique + hash-equal by construction).
    ops.push(
      makeOp('equal', original[oIdx], revised[rIdx], author, date, counter, ['substantiveText']),
    );
    prevO = oIdx + 1;
    prevR = rIdx + 1;
  }
  // Trailing gap after the last anchor.
  emitSegment(prevO, original.length, prevR, revised.length);

  return ops;
}

/**
 * Find anchor pairs: hashes that are unique on BOTH sides and equal across
 * sides. Keep the longest monotonically-increasing chain (a proper Longest
 * Increasing Subsequence on the revised indices) so reordered anchors don't
 * produce crossing alignments — the original-side list is already ascending
 * because we iterate original order, so an LIS on the revised index yields the
 * largest order-preserving anchor set and drops only the genuinely moved paras.
 */
function findAnchors(original: ParaRecord[], revised: ParaRecord[]): Array<[number, number]> {
  const origCount = new Map<string, number>();
  const origFirst = new Map<string, number>();
  for (const [i, para] of original.entries()) {
    const h = para.hash;
    origCount.set(h, (origCount.get(h) ?? 0) + 1);
    if (!origFirst.has(h)) origFirst.set(h, i);
  }

  const revCount = new Map<string, number>();
  const revFirst = new Map<string, number>();
  for (const [j, para] of revised.entries()) {
    const h = para.hash;
    revCount.set(h, (revCount.get(h) ?? 0) + 1);
    if (!revFirst.has(h)) revFirst.set(h, j);
  }

  // Candidate anchors in original order: unique on both sides, equal hash.
  const candidates: Array<[number, number]> = [];
  for (const [i, para] of original.entries()) {
    const h = para.hash;
    if (origCount.get(h) === 1 && revCount.get(h) === 1) {
      candidates.push([i, revFirst.get(h) as number]);
    }
  }

  // Keep the LONGEST strictly-increasing-on-revised-index chain. The candidate
  // list is already ascending in original index, so a proper Longest Increasing
  // Subsequence on the revised indices retains the largest possible set of
  // order-preserving anchors and drops only the genuinely reordered ones — those
  // fall back to insert/delete and are re-paired as moves at Level 3 (§5.4).
  //
  // A naive greedy "accept any pair whose revised index beats the last accepted
  // one" is NOT an LIS: a single moved paragraph whose revised index happens to
  // come first (e.g. para 1 relocated to the end → revised sequence [9,0,1,…,8])
  // pins the moved paragraph as the sole anchor and forces the entire common run
  // to redline — the exact inversion of the intended redline. The real LIS pins
  // the common run and isolates the moved paragraph instead.
  return longestIncreasingByRevised(candidates);
}

/**
 * Longest strictly-increasing subsequence of `candidates` by their revised index
 * (element `[1]`). Patience-sorting, O(n log n). Returns the chosen pairs in
 * ascending order (which is ascending in both original and revised index, since
 * the input is already ascending in original index). Revised indices are unique
 * across candidates (each is the position of a both-sides-unique hash), so strict
 * increase is well-defined.
 */
function longestIncreasingByRevised(
  candidates: Array<[number, number]>,
): Array<[number, number]> {
  const n = candidates.length;
  if (n === 0) return [];

  const revAt = (i: number): number => (candidates[i] as [number, number])[1];

  // tails[k] = candidate index of the smallest tail value of an increasing
  // subsequence of length k+1; prev[i] links each candidate to its predecessor.
  const tails: number[] = [];
  const prev: number[] = new Array<number>(n).fill(-1);

  for (let i = 0; i < n; i++) {
    const r = revAt(i);
    // First tails slot whose tail value is >= r (strict-increase insertion point).
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (revAt(tails[mid] as number) < r) lo = mid + 1;
      else hi = mid;
    }
    if (lo > 0) prev[i] = tails[lo - 1] as number;
    if (lo === tails.length) tails.push(i);
    else tails[lo] = i;
  }

  // Reconstruct by walking predecessors back from the last tail.
  const chain: Array<[number, number]> = [];
  for (let k: number = tails[tails.length - 1] as number; k !== -1; k = prev[k] as number) {
    chain.push(candidates[k] as [number, number]);
  }
  chain.reverse();
  return chain;
}

/**
 * Diff a single anchor-bounded segment [oStart,oEnd) × [rStart,rEnd) via LCS on
 * paragraph hashes, emitting equal/replace/insert/delete ops in order.
 */
function diffSegment(
  original: ParaRecord[],
  revised: ParaRecord[],
  oStart: number,
  oEnd: number,
  rStart: number,
  rEnd: number,
  author: string,
  date: string,
  counter: { next: number },
  ops: DiffOp[],
  pairing: GapPairing | undefined,
): void {
  const oSlice = original.slice(oStart, oEnd);
  const rSlice = revised.slice(rStart, rEnd);

  if (oSlice.length === 0 && rSlice.length === 0) return;

  // Pure insert / pure delete shortcuts.
  if (oSlice.length === 0) {
    for (let j = 0; j < rSlice.length; j++) {
      ops.push(makeOp('insert', null, rSlice[j], author, date, counter, ['structural']));
    }
    return;
  }
  if (rSlice.length === 0) {
    for (let i = 0; i < oSlice.length; i++) {
      ops.push(makeOp('delete', oSlice[i], null, author, date, counter, ['structural']));
    }
    return;
  }

  const matches = computeLCS(oSlice, rSlice, (p) => p.hash);
  // matches are local indices into the slices; walk both cursors. Each gap
  // between consecutive matches (and the leading/trailing gaps) is paired into
  // `replace` ops (unequal hash — §5.1) with surplus on either side falling out
  // as insert/delete.
  let oi = 0;
  let ri = 0;
  for (const [mo, mr] of matches) {
    emitGap(oSlice, rSlice, oi, mo, ri, mr, author, date, counter, ops, pairing);
    // The matched pair (hash-equal) → equal.
    ops.push(
      makeOp('equal', oSlice[mo], rSlice[mr], author, date, counter, ['substantiveText']),
    );
    oi = mo + 1;
    ri = mr + 1;
  }
  // Trailing gap after the last match.
  emitGap(oSlice, rSlice, oi, oSlice.length, ri, rSlice.length, author, date, counter, ops, pairing);
}

/**
 * Emit a gap [oStart,oEnd) × [rStart,rEnd) between two LCS matches (ADR-0039).
 *
 * Without a `pairing`, positional exactly as before. With one:
 *
 *  1. **Lock** each positional pair that renders inline today — decided by
 *     running ADR-0032's merge/split pass and ADR-0026's demotion over the
 *     positional ops — so it keeps its partner.
 *  2. **Align** each stretch between locked pairs by similarity, order-preserving
 *     (a stretch holding a table unit stays positional; a record whose text is
 *     on the other side verbatim is held for Level 3's move detection).
 *  3. **Fall back** to positional for whatever the alignment leaves.
 *
 * Locked pairs sit on the gap's diagonal, so the stretch between two of them is
 * square and nothing can be paired across one.
 */
function emitGap(
  oSlice: ParaRecord[],
  rSlice: ParaRecord[],
  oStart: number,
  oEnd: number,
  rStart: number,
  rEnd: number,
  author: string,
  date: string,
  counter: { next: number },
  ops: DiffOp[],
  pairing: GapPairing | undefined,
): void {
  if (pairing === undefined) {
    emitPositional(oSlice, rSlice, oStart, oEnd, rStart, rEnd, author, date, counter, ops);
    return;
  }

  const locked = pairing.lockPositional(oSlice.slice(oStart, oEnd), rSlice.slice(rStart, rEnd));
  let o = oStart;
  let r = rStart;
  for (let k = 0; k < locked.length; k++) {
    const lockedO = oStart + k;
    const lockedR = rStart + k;
    if (!locked[k]) continue;
    emitAligned(oSlice, rSlice, o, lockedO, r, lockedR, author, date, counter, ops, pairing);
    ops.push(
      makeOp('replace', oSlice[lockedO], rSlice[lockedR], author, date, counter, ['substantiveText']),
    );
    o = lockedO + 1;
    r = lockedR + 1;
  }
  emitAligned(oSlice, rSlice, o, oEnd, r, rEnd, author, date, counter, ops, pairing);
}

/**
 * Rules 2 and 3 of {@link emitGap} over one stretch with no locked pair in it:
 * the similarity alignment's pairs as `replace`, and positional pairing for
 * each stretch the alignment leaves between them.
 */
function emitAligned(
  oSlice: ParaRecord[],
  rSlice: ParaRecord[],
  oStart: number,
  oEnd: number,
  rStart: number,
  rEnd: number,
  author: string,
  date: string,
  counter: { next: number },
  ops: DiffOp[],
  pairing: GapPairing,
): void {
  const pairs = pairing.align(oSlice.slice(oStart, oEnd), rSlice.slice(rStart, rEnd));
  if (pairs.length === 0) {
    // Nothing re-paired: the stretch is exactly what positional pairing made it.
    emitPositional(oSlice, rSlice, oStart, oEnd, rStart, rEnd, author, date, counter, ops);
    return;
  }
  let o = oStart;
  let r = rStart;
  for (const [a, b] of pairs) {
    emitFallback(oSlice, rSlice, o, oStart + a, r, rStart + b, author, date, counter, ops, pairing);
    ops.push(
      makeOp('replace', oSlice[oStart + a], rSlice[rStart + b], author, date, counter, ['substantiveText']),
    );
    o = oStart + a + 1;
    r = rStart + b + 1;
  }
  emitFallback(oSlice, rSlice, o, oEnd, r, rEnd, author, date, counter, ops, pairing);
}

/**
 * Rule 3 beside re-aligned pairs: positional, except that a record held for a
 * move (`GapPairing.heldForMove`) is never paired. It is emitted as the
 * `delete` / `insert` it would have become, so Level 3 still pairs it by hash.
 * With no held record this is exactly {@link emitPositional}.
 */
function emitFallback(
  oSlice: ParaRecord[],
  rSlice: ParaRecord[],
  oStart: number,
  oEnd: number,
  rStart: number,
  rEnd: number,
  author: string,
  date: string,
  counter: { next: number },
  ops: DiffOp[],
  pairing: GapPairing,
): void {
  const oFree: number[] = [];
  const rFree: number[] = [];
  for (let i = oStart; i < oEnd; i++) if (!pairing.heldForMove(oSlice[i]!, 'original')) oFree.push(i);
  for (let j = rStart; j < rEnd; j++) if (!pairing.heldForMove(rSlice[j]!, 'revised')) rFree.push(j);
  if (oFree.length === oEnd - oStart && rFree.length === rEnd - rStart) {
    emitPositional(oSlice, rSlice, oStart, oEnd, rStart, rEnd, author, date, counter, ops);
    return;
  }
  // Walk both sides in DOCUMENT order: before each pair, the records it
  // passes over on each side (held ones) go out as delete / insert. Emitting
  // the pairs first and the held records after — the first version — put a
  // held record behind a later paragraph of its own side, so the redline no
  // longer rebuilt its source in order (08f task 9).
  const paired = Math.min(oFree.length, rFree.length);
  let i = oStart;
  let j = rStart;
  for (let k = 0; k < paired; k++) {
    const pairO = oFree[k]!;
    const pairR = rFree[k]!;
    for (; i < pairO; i++) ops.push(makeOp('delete', oSlice[i], null, author, date, counter, ['structural']));
    for (; j < pairR; j++) ops.push(makeOp('insert', null, rSlice[j], author, date, counter, ['structural']));
    ops.push(makeOp('replace', oSlice[pairO], rSlice[pairR], author, date, counter, ['substantiveText']));
    i = pairO + 1;
    j = pairR + 1;
  }
  for (; i < oEnd; i++) ops.push(makeOp('delete', oSlice[i], null, author, date, counter, ['structural']));
  for (; j < rEnd; j++) ops.push(makeOp('insert', null, rSlice[j], author, date, counter, ['structural']));
}

/**
 * Positional pairing of [oStart,oEnd) × [rStart,rEnd): index-aligned leftover
 * pairs (one original + one revised, no hash match) are `replace` (Level-2 word
 * diff filled later by A3/I1); surplus on either side is insert/delete. The
 * whole of Level 1's gap handling before ADR-0039, and still all of it on the
 * PDF path and in the parts diff.
 */
function emitPositional(
  oSlice: ParaRecord[],
  rSlice: ParaRecord[],
  oStart: number,
  oEnd: number,
  rStart: number,
  rEnd: number,
  author: string,
  date: string,
  counter: { next: number },
  ops: DiffOp[],
): void {
  const oRemain = oEnd - oStart;
  const rRemain = rEnd - rStart;
  const paired = Math.min(oRemain, rRemain);
  const oiStart = oStart;
  const riStart = rStart;

  for (let k = 0; k < paired; k++) {
    ops.push(
      makeOp(
        'replace',
        oSlice[oiStart + k],
        rSlice[riStart + k],
        author,
        date,
        counter,
        ['substantiveText'],
      ),
    );
  }
  for (let k = paired; k < oRemain; k++) {
    ops.push(makeOp('delete', oSlice[oiStart + k], null, author, date, counter, ['structural']));
  }
  for (let k = paired; k < rRemain; k++) {
    ops.push(makeOp('insert', null, rSlice[riStart + k], author, date, counter, ['structural']));
  }
}

/** Construct a DiffOp with a monotonic string id and placeholder classification. */
function makeOp(
  type: DiffOp['type'],
  orig: ParaRecord | null | undefined,
  rev: ParaRecord | null | undefined,
  author: string,
  date: string,
  counter: { next: number },
  classes: [ChangeClass, ...ChangeClass[]],
): DiffOp {
  const id = `op-${counter.next++}`;
  const partId = (rev ?? orig)?.partId ?? 'document';
  return {
    id,
    type,
    origRange: orig ? { start: orig.index, end: orig.index } : null,
    revRange: rev ? { start: rev.index, end: rev.index } : null,
    partId,
    classification: classes[0],
    classes,
    author,
    date,
  };
}
