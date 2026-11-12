/**
 * Level-1 gap pairing by similarity — ADR-0039.
 * Spec: spec/features/08-compare.md §5.1.
 *
 * Between two LCS matches, `structuralDiff` is left with a GAP: originals and
 * reviseds with no hash match. It used to pair them by position, so one
 * inserted or deleted paragraph in the gap shifted every pairing after it,
 * and ADR-0026 then demoted the mismatched pairs to whole-paragraph churn.
 *
 * This module supplies the two decisions `emitGap` now asks, and a third that
 * Stage 2.25 asks after it:
 *
 *  1. **Lock** (`lockPositional`) — which POSITIONAL pairs render inline today?
 *     Answered by running the real passes that decide it — ADR-0032's
 *     merge/split re-pairing, then ADR-0026's demotion — over the gap's
 *     positional ops, never by re-deriving their rules. A locked pair keeps its
 *     partner. Asking the demotion gate about a pair ON ITS OWN was the first
 *     version, and it was wrong: a merge whose head was edited fails the gate
 *     alone and renders inline only because the merge pass scores the JOINED
 *     text (08f task 7).
 *  2. **Align** (`align`) — in each stretch between locked pairs, the
 *     ORDER-PRESERVING set of pairs whose Dice clears the demotion threshold,
 *     with maximal total Dice. Pairs never cross, and a gap is bounded by
 *     matched paragraphs on both sides, so nothing is ever paired across
 *     unchanged text — that would be move detection, v2 for Word-vs-Word (SCOPE).
 *     A stretch holding a table unit (ADR-0027) is left POSITIONAL, exactly as
 *     before. A paragraph whose text appears verbatim on the other side is
 *     HELD (`heldForMove`): never a candidate, and left unpaired by the
 *     fallback wherever the stretch is re-aligned. Holding it rather than
 *     skipping its whole stretch matters on tables, where a cell value such as
 *     a severity level recurs and moves while its neighbours are edited.
 *     Alignment is bounded per stretch (`MAX_ALIGN_WORK`, `MAX_ALIGN_CELLS`)
 *     and across the comparison (`MAX_ALIGN_WORK_TOTAL`); over any of them a
 *     stretch stays positional.
 *  3. **Absorb** (`mayAbsorb`) — may ADR-0032's merge/split pass take this
 *     record as an orphan? Not a held one, unless the lock probe's own merge
 *     pass absorbed it. Holding a record back from the alignment and the
 *     fallback was not enough: re-pairing could put it straight after a replace
 *     it never followed, Stage 2.25 absorbed it, and its move was lost (08f
 *     task 11). With all three, Level 3's hash-based move detection finds every
 *     held record exactly as it would have.
 *
 * Pure: no Date.now(), no I/O; caches live in the object `gapPairingFor` returns,
 * one per pipeline run.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { DiffOp, ParaRecord } from '../model.js';

import { detectMergesAndSplits } from './merge-split.js';
import { contentTokens, demoteDissimilarReplaces, diceFromOverlap } from './replace-threshold.js';

/** The pairing decisions `structuralDiff` asks of a gap. */
export interface GapPairing {
  /**
   * Rule 1, for one whole gap: `locked[k]` is true when the k-th original and
   * the k-th revised — the positional pair — render inline today. Locked pairs
   * are never re-paired.
   */
  lockPositional(original: readonly ParaRecord[], revised: readonly ParaRecord[]): boolean[];
  /**
   * Rule 2, for one stretch with no locked pair in it: the pairs to make, as
   * ascending local `[originalIndex, revisedIndex]`. `[]` means positional.
   */
  align(original: readonly ParaRecord[], revised: readonly ParaRecord[]): Array<[number, number]>;
  /**
   * Is this record's text present verbatim on the other side? Such a record is
   * never paired by the alignment or by the fallback beside it (it is headed for
   * Level 3's move detection, or for whatever positional pairing gave it).
   */
  heldForMove(record: ParaRecord, side: 'original' | 'revised'): boolean;
  /**
   * May Stage 2.25 (`detectMergesAndSplits`) absorb this record as a merge or
   * split orphan? Not a held record — unless positional pairing's own merge
   * pass absorbed it, which the lock probe records. Re-pairing can put a held
   * record directly after a replace it never followed before; absorbed there,
   * its verbatim twin loses its partner and the move becomes a plain insertion
   * (08f task 11). Meaningful only after `structuralDiff` has run with this
   * pairing, since the lock probes are what record the positional merges.
   */
  mayAbsorb(record: ParaRecord, side: 'original' | 'revised'): boolean;
  /**
   * A pair's weight as an alignment candidate: its Dice when that clears the
   * demotion threshold (the same value `paragraphSimilarity` gives), else 0.
   * Contentless paragraphs score 0.
   */
  score(original: ParaRecord, revised: ParaRecord): number;
}

/**
 * Work budget for one stretch's alignment, in cost units: cells × (fixed
 * per-cell cost + mean distinct tokens per paragraph). Above it the stretch
 * keeps positional pairing (an accepted limit, ADR-0039).
 *
 * Fitted, not guessed (08f task 10): `align`'s time per cell measured
 * ~90 ns at 1 token, ~256 ns at 40, ~566 ns at 100, ~1,860 ns at 600 —
 * linear at ~3 ns per distinct token with a fixed cost worth
 * {@link ALIGN_CELL_OVERHEAD} tokens (the score closure, the held-record and
 * cache lookups, the bound check). At 10⁸ units that is ~0.3 s here; a machine
 * twice as slow (a reviewer measured one) ~0.6 s. Two budgets came before
 * this one and both were wrong: a bare cell count could not see paragraph
 * length (24 s at 600 tokens, task 7), and cells × tokens alone could not see
 * the fixed cost (7 to 19 s for many very short paragraphs, task 9).
 */
export const MAX_ALIGN_WORK = 100_000_000;

/**
 * Work budget for the whole COMPARISON, in the same units: {@link MAX_ALIGN_WORK}
 * bounds one stretch, and a document of many large gaps paid it once per gap —
 * ten 999 × 1,000 gaps of 40-token paragraphs took 6.3 s of Stage 2 against
 * 0.15 s positional (08f task 11). Once the stretches aligned so far have spent
 * this, every later stretch keeps positional pairing. Deterministic: the gaps
 * are visited in document order.
 */
export const MAX_ALIGN_WORK_TOTAL = 300_000_000;

/** The fixed cost of scoring one cell, in distinct-token units (see {@link MAX_ALIGN_WORK}). */
export const ALIGN_CELL_OVERHEAD = 32;

/**
 * Cell cap for one stretch's alignment, beside {@link MAX_ALIGN_WORK}. The work
 * budget is token-weighted, so a stretch of very SHORT paragraphs passes it with
 * millions of cells — 8,600 × 8,600 two-token paragraphs took 7.2 s and ~0.9 GB
 * (08f task 9), because `alignGap` holds two `Float64Array`s of m·n cells. At
 * this cap they hold ~16 MB together.
 */
export const MAX_ALIGN_CELLS = 1_000_000;

/** A record's content tokens as interned ids, sorted, with a count per id. */
interface IdBag {
  /** Total tokens, counting repeats — Dice's denominator. */
  size: number;
  ids: Int32Array;
  counts: Int32Array;
}

/**
 * Build the pairing for one comparison. `original` and `revised` are the WHOLE
 * documents' records, indexed by `ParaRecord.index` — the merge/split pass the
 * lock runs reads neighbours through them.
 */
export function gapPairingFor(
  config: CompareConfig,
  original: readonly ParaRecord[],
  revised: readonly ParaRecord[],
): GapPairing {
  const threshold = config.replaceSimilarityThreshold;
  const originalHashes = new Set(original.map((p) => p.hash));
  const revisedHashes = new Set(revised.map((p) => p.hash));

  // Token bags, interned and cached per record. The cache is load-bearing:
  // the alignment scores every cell of a stretch, and re-tokenising per cell
  // took a 950-paragraph unanchored document from ~1.5 s to 86 s.
  const ids = new Map<string, number>();
  const bags = new WeakMap<ParaRecord, IdBag>();
  const bagOfRecord = (rec: ParaRecord): IdBag => {
    let bag = bags.get(rec);
    if (bag === undefined) {
      const counted = new Map<number, number>();
      const tokens = contentTokens(rec.normText, config);
      for (const t of tokens) {
        let id = ids.get(t);
        if (id === undefined) {
          id = ids.size;
          ids.set(t, id);
        }
        counted.set(id, (counted.get(id) ?? 0) + 1);
      }
      const sorted = [...counted.keys()].sort((a, b) => a - b);
      bag = {
        size: tokens.length,
        ids: Int32Array.from(sorted),
        counts: Int32Array.from(sorted, (id) => counted.get(id)!),
      };
      bags.set(rec, bag);
    }
    return bag;
  };

  const score = (o: ParaRecord, r: ParaRecord): number => {
    const a = bagOfRecord(o);
    const b = bagOfRecord(r);
    // Two contentless paragraphs (`* * *` against `- - -`) are identical to
    // Dice (two empty bags score 1) and would out-bid a real counterpart they
    // cross. Never candidates.
    if (a.size === 0 || b.size === 0) return 0;
    // Dice can be no higher than 2·min/(a+b); skip what cannot reach the bar.
    if ((2 * Math.min(a.size, b.size)) / (a.size + b.size) < threshold) return 0;
    const dice = diceFromOverlap(sortedOverlap(a, b), a.size, b.size);
    return dice >= threshold && dice > 0 ? dice : 0;
  };

  const heldForMove = (rec: ParaRecord, side: 'original' | 'revised'): boolean =>
    (side === 'original' ? revisedHashes : originalHashes).has(rec.hash);

  // Records positional pairing's own merge pass absorbed (filled by the lock
  // probes), and the alignment work spent so far across the comparison.
  const absorbedPositionally = new Set<ParaRecord>();
  let workSpent = 0;

  return {
    score,
    heldForMove,

    mayAbsorb(rec, side) {
      return !heldForMove(rec, side) || absorbedPositionally.has(rec);
    },

    lockPositional(o, r) {
      const diagonal = Math.min(o.length, r.length);
      const locked = new Array<boolean>(diagonal).fill(false);
      if (diagonal === 0) return locked;

      // The gap's positional ops, exactly as `emitPositional` emits them, run
      // through the two passes that decide what renders inline before Level 2.
      const ops: DiffOp[] = [];
      for (let k = 0; k < diagonal; k++) ops.push(probeOp('replace', o[k]!, r[k]!));
      for (let k = diagonal; k < o.length; k++) ops.push(probeOp('delete', o[k]!, null));
      for (let k = diagonal; k < r.length; k++) ops.push(probeOp('insert', null, r[k]!));
      const settled = demoteDissimilarReplaces(
        detectMergesAndSplits(ops, original as ParaRecord[], revised as ParaRecord[], config),
        original as ParaRecord[],
        revised as ParaRecord[],
        config,
      );

      // A surviving replace — merged/split or not — starts at its positional
      // pair's original. A merge's absorbed orphan is the first surplus record,
      // emitted straight after the locked pair, so it needs no lock of its own.
      const diagonalOf = new Map<number, number>();
      for (let k = 0; k < diagonal; k++) diagonalOf.set(o[k]!.index, k);
      for (const op of settled) {
        if (op.type !== 'replace' || op.origRange === null || op.revRange === null) continue;
        const k = diagonalOf.get(op.origRange.start);
        if (k !== undefined) locked[k] = true;
        for (let i = op.origRange.start + 1; i <= op.origRange.end; i++) absorbedPositionally.add(original[i]!);
        for (let j = op.revRange.start + 1; j <= op.revRange.end; j++) absorbedPositionally.add(revised[j]!);
      }
      return locked;
    },

    align(o, r) {
      if (o.length === 0 || r.length === 0) return [];
      if (o.some((p) => p.table !== undefined) || r.some((p) => p.table !== undefined)) return [];
      const meanDistinct =
        (o.reduce((n, p) => n + bagOfRecord(p).ids.length, 0) / o.length +
          r.reduce((n, p) => n + bagOfRecord(p).ids.length, 0) / r.length) /
        2;
      if (o.length * r.length > MAX_ALIGN_CELLS) return [];
      const work = o.length * r.length * (ALIGN_CELL_OVERHEAD + meanDistinct);
      if (work > MAX_ALIGN_WORK) return [];
      if (workSpent + work > MAX_ALIGN_WORK_TOTAL) return [];
      workSpent += work;
      return alignGap(o, r, (x, y) =>
        heldForMove(x, 'original') || heldForMove(y, 'revised') ? 0 : score(x, y),
      );
    },
  };
}

/**
 * Multiset intersection size of two sorted id bags: Σ min(count) over shared ids.
 *
 * The same quantity `bagDice` computes over string maps, by a second algorithm,
 * kept because it is measured 3–4.5× faster on the all-pairs scoring this module
 * does (999 × 999 cells of 100-token paragraphs: 3,154 ms with string maps,
 * 971 ms here; 600 × 600 at 140 tokens: 1,729 ms against 384 ms — 08f task 9).
 * The formula is shared (`diceFromOverlap`), and a property test pins that the
 * two agree on every pair the alignment scores.
 */
function sortedOverlap(a: IdBag, b: IdBag): number {
  let i = 0;
  let j = 0;
  let overlap = 0;
  while (i < a.ids.length && j < b.ids.length) {
    const x = a.ids[i]!;
    const y = b.ids[j]!;
    if (x === y) {
      overlap += Math.min(a.counts[i]!, b.counts[j]!);
      i += 1;
      j += 1;
    } else if (x < y) {
      i += 1;
    } else {
      j += 1;
    }
  }
  return overlap;
}

/** A positional op for the lock's probe run — never emitted. */
function probeOp(type: 'replace' | 'delete' | 'insert', o: ParaRecord | null, r: ParaRecord | null): DiffOp {
  return {
    id: 'probe',
    type,
    origRange: o ? { start: o.index, end: o.index } : null,
    revRange: r ? { start: r.index, end: r.index } : null,
    partId: (r ?? o)!.partId,
    classification: 'structural',
    classes: ['structural'],
    author: '',
    date: '',
  };
}

/**
 * The order-preserving set of pairs with maximal total `score`, as ascending
 * `[originalIndex, revisedIndex]` pairs local to the two arrays. Only pairs
 * with a positive score are ever chosen.
 *
 * A Needleman–Wunsch-shaped dynamic programme with zero gap cost: O(m·n) time
 * and space. On a tie it takes the pair, then skips an ORIGINAL, so the result
 * is deterministic. The caller bounds the work ({@link MAX_ALIGN_WORK}).
 */
export function alignGap(
  original: readonly ParaRecord[],
  revised: readonly ParaRecord[],
  score: (o: ParaRecord, r: ParaRecord) => number,
): Array<[number, number]> {
  const m = original.length;
  const n = revised.length;
  if (m === 0 || n === 0) return [];

  const weight = new Float64Array(m * n);
  let anyCandidate = false;
  for (let i = 0; i < m; i++) {
    for (let j = 0; j < n; j++) {
      const w = score(original[i]!, revised[j]!);
      weight[i * n + j] = w;
      if (w > 0) anyCandidate = true;
    }
  }
  if (!anyCandidate) return [];

  // best[i][j] = the maximal total over original[i..] × revised[j..].
  const width = n + 1;
  const best = new Float64Array((m + 1) * width);
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      const w = weight[i * n + j]!;
      const take = w > 0 ? w + best[(i + 1) * width + j + 1]! : 0;
      best[i * width + j] = Math.max(take, best[(i + 1) * width + j]!, best[i * width + j + 1]!);
    }
  }

  // Walk the table forward, re-deriving each step from the same arithmetic, so
  // the comparisons below are exact equalities on identical expressions.
  const pairs: Array<[number, number]> = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    const here = best[i * width + j]!;
    if (here === 0) break;
    const w = weight[i * n + j]!;
    if (w > 0 && w + best[(i + 1) * width + j + 1]! === here) {
      pairs.push([i, j]);
      i += 1;
      j += 1;
    } else if (best[(i + 1) * width + j]! === here) {
      i += 1;
    } else {
      j += 1;
    }
  }
  return pairs;
}
