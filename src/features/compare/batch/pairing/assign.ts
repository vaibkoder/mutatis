/**
 * Turning a score matrix into a proposed pairing (ADR-0037; spec §6.4).
 *
 * Two bundles go in; a set of proposed pairs plus the documents that pair with
 * nothing comes out. Nothing here decides anything final — the result is a
 * proposal the user confirms (§6.6), and a document that falls out as unmatched
 * is *shown* as unmatched rather than quietly dropped.
 *
 * ## The padding, and why the no-match cost is what it is
 *
 * A plain M×N assignment forces every document to pair with something. That is
 * wrong whenever a bundle gained a side letter, lost a schedule, or replaced one
 * agreement with a different one — the classic case where a forced pairing
 * produces a complete, well-formed, entirely worthless whole-document rewrite.
 *
 * So the matrix is padded to (M+N)×(M+N): M real rows plus N dummy rows, N real
 * columns plus M dummy columns. Every real document can therefore be assigned
 * to a dummy — "unmatched" is always an available option, even when the bundles
 * are the same size.
 *
 * The cost of taking that option is derived, not chosen. Leaving a real A and a
 * real B both unmatched costs `2·NO_MATCH_COST` (each pays once); pairing them
 * costs `1 − score`. The solver prefers the pair exactly when
 *
 *     1 − score < 2·NO_MATCH_COST
 *
 * and the configured `noMatchScore` is defined as the score below which a document
 * should fall out, so we need that inequality to be `score > NO_MATCH_SCORE`.
 * Solving gives `NO_MATCH_COST = (1 − NO_MATCH_SCORE) / 2`. Setting it to the
 * more obvious `1 − NO_MATCH_SCORE` would double the real threshold and pair
 * documents that share almost nothing — which is the one failure mode here that
 * does not announce itself.
 *
 * Dummy-to-dummy costs nothing: those cells only exist to make the matrix
 * square, and pricing them would tilt the optimum by an amount that depends on
 * the bundle sizes.
 *
 * Pure — no Office JS, no DOM, no network.
 */
import { DEFAULT_PAIRING_SETTINGS, type PairingSettings } from '../config.js';
import { hungarian } from './hungarian.js';
import { scorePair, type PairScore } from './score.js';
import type { PairingSketch } from './sketch.js';

/**
 * See the module comment for the derivation. Not `1 − noMatchScore`.
 *
 * A function of the settings rather than a constant, because the threshold it
 * is derived from is configurable: a profile that moves `noMatchScore` and left
 * this behind would price the dummies for a threshold nobody is using.
 */
export function noMatchCost(settings: PairingSettings = DEFAULT_PAIRING_SETTINGS): number {
  return (1 - settings.noMatchScore) / 2;
}

/** One proposed pairing, carrying the scores that produced it. */
export interface ProposedPair {
  a: string;
  b: string;
  score: PairScore;
}

/** The assignment: what pairs with what, and what pairs with nothing. */
export interface Assignment {
  pairs: readonly ProposedPair[];
  /** Side-A documents with no counterpart — removed, or replaced. */
  unmatchedA: readonly string[];
  /** Side-B documents with no counterpart — added, or replacing. */
  unmatchedB: readonly string[];
  /** Every candidate score, kept so the grid can show rivals (§6.5). */
  scores: readonly PairScore[];
}

/** Stable ordering by `docId`, so the result is a function of the input SET. */
function byDocId(a: PairingSketch, b: PairingSketch): number {
  return a.docId < b.docId ? -1 : a.docId > b.docId ? 1 : 0;
}

/**
 * Propose a pairing between two bundles.
 *
 * Both sides are sorted by `docId` before the matrix is built, and the output
 * is emitted in that order. The optimum is not always unique — two assignments
 * can cost exactly the same — and without a canonical input order the solver
 * would be free to return a different one of them for a different input order.
 * Acceptance M4 requires that shuffling both bundles changes nothing, so the
 * order is pinned here rather than hoped for.
 */
export function assignPairs(
  sideA: readonly PairingSketch[],
  sideB: readonly PairingSketch[],
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): Assignment {
  const as = [...sideA].sort(byDocId);
  const bs = [...sideB].sort(byDocId);

  if (as.length === 0 || bs.length === 0) {
    return {
      pairs: [],
      unmatchedA: as.map(a => a.docId),
      unmatchedB: bs.map(b => b.docId),
      scores: [],
    };
  }

  const scores: PairScore[][] = as.map(a => bs.map(b => scorePair(a, b, settings)));
  const scoreAt = (i: number, j: number): PairScore =>
    (scores[i] as PairScore[])[j] as PairScore;

  // ── Phase 1: pairs the CONTENT corroborates ──────────────────────────────
  //
  // A cell with no content evidence is priced out of this phase entirely, so
  // it cannot take a document away from a pairing that content supports. That
  // displacement is not hypothetical: scored on the same scale, a `.pdf` whose
  // filename matched perfectly beat every content-verified pair in the matrix
  // and left the real counterpart labelled "no counterpart in the other
  // bundle".
  const phase1 = solve(
    as.length,
    bs.length,
    (i, j) => (scoreAt(i, j).contentAvailable ? 1 - scoreAt(i, j).score : null),
    noMatchCost(settings),
  );

  // ── Phase 2: what is left, on the filename alone ─────────────────────────
  //
  // Only documents phase 1 could not place, and only cells where the content
  // signal was UNAVAILABLE rather than merely weak. A content-available pair
  // that lost in phase 1 lost on evidence, and re-running it here on its name
  // would be the force-matching the no-match dummy exists to prevent.
  //
  // Being a separate phase rather than a lower score is what makes both halves
  // true at once: filename-only evidence is enough to beat *nothing* — so a
  // cross-format pair still forms and is still surfaced with its reason
  // (acceptance M8) whatever its names look like — and never enough to beat
  // *something*, because by the time this phase runs there is nothing left for
  // it to beat. A single scale cannot do both: pitched high it displaces real
  // pairs, pitched low it forms none.
  const phase2 = solve(
    phase1.unmatchedRows.length,
    phase1.unmatchedCols.length,
    (i, j) => {
      const ai = phase1.unmatchedRows[i] as number;
      const bj = phase1.unmatchedCols[j] as number;
      return scoreAt(ai, bj).contentAvailable ? null : 1 - scoreAt(ai, bj).filename;
    },
    (1 - settings.noContentMatchScore) / 2,
  );

  const pairs: ProposedPair[] = [
    ...phase1.pairs.map(([i, j]) => ({
      a: (as[i] as PairingSketch).docId,
      b: (bs[j] as PairingSketch).docId,
      score: scoreAt(i, j),
    })),
    ...phase2.pairs.map(([i, j]) => {
      const ai = phase1.unmatchedRows[i] as number;
      const bj = phase1.unmatchedCols[j] as number;
      return {
        a: (as[ai] as PairingSketch).docId,
        b: (bs[bj] as PairingSketch).docId,
        score: scoreAt(ai, bj),
      };
    }),
  ];

  // Emitted in `docId` order, not in phase order: which phase placed a pair is
  // an implementation detail, and M4 requires the output to be a function of
  // the two SETS.
  pairs.sort((x, y) => (x.a < y.a ? -1 : x.a > y.a ? 1 : 0));

  return {
    pairs,
    unmatchedA: phase2.unmatchedRows
      .map(i => (as[phase1.unmatchedRows[i] as number] as PairingSketch).docId)
      .sort(),
    unmatchedB: phase2.unmatchedCols
      .map(j => (bs[phase1.unmatchedCols[j] as number] as PairingSketch).docId)
      .sort(),
    scores: scores.flat(),
  };
}

/**
 * One Hungarian solve over an m×n rectangle, padded so "unmatched" is always
 * available.
 *
 * `costOf` returns `null` for a cell this phase must never pair. Such a cell is
 * priced strictly above the cost of leaving both its documents unmatched, which
 * is what "never" means to a cost-minimising solver — there is no sentinel
 * value to leak, and no branch the optimum can take.
 *
 * The `+ 1` is the *strictly*, and it is deliberately not merely `2 * dummyCost`:
 * at exactly the dummy price the two options tie, and a tie is resolved by
 * whatever the solver's iteration order happens to do. Today `hungarian` breaks
 * it toward the dummy, so no test can distinguish the two — recorded here rather
 * than papered over with a test that would only be asserting an accident of the
 * solver. What must not depend on that accident is a pairing: a tie broken the
 * other way force-matches two documents this phase has just said must never
 * pair, which is the failure mode that produces a well-formed redline of one
 * agreement against an unrelated one.
 *
 * Indices are positions in the rectangle, not `docId`s: the caller owns the
 * mapping, because in the second phase they index a subset of the first.
 */
function solve(
  m: number,
  n: number,
  costOf: (i: number, j: number) => number | null,
  dummyCost: number,
): { pairs: [number, number][]; unmatchedRows: number[]; unmatchedCols: number[] } {
  if (m === 0 || n === 0) {
    return {
      pairs: [],
      unmatchedRows: Array.from({ length: m }, (_, i) => i),
      unmatchedCols: Array.from({ length: n }, (_, j) => j),
    };
  }

  const never = 2 * dummyCost + 1;
  const size = m + n;
  const cost: number[][] = [];
  for (let i = 0; i < size; i += 1) {
    const row = new Array<number>(size);
    for (let j = 0; j < size; j += 1) {
      const realRow = i < m;
      const realCol = j < n;
      if (realRow && realCol) row[j] = costOf(i, j) ?? never;
      else if (realRow || realCol) row[j] = dummyCost;
      else row[j] = 0;
    }
    cost.push(row);
  }

  const rowToCol = hungarian(cost);
  const pairs: [number, number][] = [];
  const matchedCols = new Set<number>();
  const unmatchedRows: number[] = [];

  for (let i = 0; i < m; i += 1) {
    const j = rowToCol[i] as number;
    if (j < n) {
      pairs.push([i, j]);
      matchedCols.add(j);
    } else {
      unmatchedRows.push(i);
    }
  }

  const unmatchedCols: number[] = [];
  for (let j = 0; j < n; j += 1) if (!matchedCols.has(j)) unmatchedCols.push(j);

  return { pairs, unmatchedRows, unmatchedCols };
}
