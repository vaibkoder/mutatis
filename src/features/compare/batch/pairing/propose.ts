/**
 * Building the pairing proposal (ADR-0037; spec §6.5, §6.6).
 *
 * Assignment says *what pairs with what*; this says *how confident that is and
 * what the alternatives were*, and produces the rows the grid renders.
 *
 * The band rule has two conditions and both matter:
 *
 *     auto-pair  ⟺  score ≥ autoPairScore
 *                   AND  score − bestRival ≥ autoPairMargin
 *
 * The score alone is not enough. Two consecutive drafts of one agreement both
 * score ~0.97 against a third — that is what near-identical *means* — so a
 * score-only rule pre-ticks a coin-toss between them. The margin is what
 * catches it: a high score with a close rival is exactly the case a human has
 * to look at, and it is acceptance M2.
 *
 * Nothing here filters. Every document appears once; every non-`auto-pair` row
 * carries its rivals with their scores. A band changes presentation only.
 *
 * Pure — no Office JS, no DOM, no network.
 */
import type { CompareConfig } from '@shared/config/index.js';
import { CROSS_FORMAT_REASON, UNSUPPORTED_FORMAT_REASON } from '../constants.js';
import {
  DEFAULT_PAIRING_SETTINGS,
  resolvePairingSettings,
  type PairingSettings,
} from '../config.js';
import {
  describeDocxCapBreach,
  describeUnpairedDocxCapBreach,
  exceedsDocxByteCap,
} from '../limits.js';
import { assignPairs, type ProposedPair } from './assign.js';
import { rankCandidates, type PairScore } from './score.js';
import type {
  PairingBand,
  PairingBlockKind,
  PairingDocument,
  PairingProposal,
  PairingRow,
} from './model.js';
import { tallyPairing } from './model.js';
import type { PairingSketch } from './sketch.js';

/** How many rivals a row carries. Two: the runner-up, and the one behind it. */
export const RIVALS_SHOWN = 2;

const CAP_DISCLOSURE =
  'Some documents are larger than this build can compare. They are shown with ' +
  'the cap named, but their comparisons will not run.';

const CROSS_FORMAT_DISCLOSURE =
  'Some pairs mix formats. They are paired and shown, but Word-against-PDF ' +
  'comparison is out of scope, so those comparisons will not run.';

const UNSUPPORTED_FORMAT_DISCLOSURE =
  'Some pairs are not Word documents. They are paired and shown, but only Word ' +
  'documents can be compared in a batch, so those comparisons will not run.';

const UNREADABLE_DISCLOSURE =
  'Some documents could not be read. They are listed with the reason, and the ' +
  'rest of the bundle was paired without them.';

/** Why a proposed pair cannot be compared, or `null` if it can. */
function blockFor(
  a: PairingSketch,
  b: PairingSketch,
): { blocked: PairingBlockKind; reason: string } | null {
  // Format first: it is the more fundamental refusal, and naming a size cap on
  // a pair that was never comparable anyway would be misleading.
  //
  // The two format refusals are kept apart. A Word against a PDF really is
  // cross-format. Two PDFs are not — they agree perfectly, on a format this
  // build does not batch — and telling the user otherwise sends them looking
  // for a mismatch that is not there.
  if (a.sourceKind !== b.sourceKind) {
    return { blocked: 'cross-format', reason: CROSS_FORMAT_REASON };
  }
  if (a.sourceKind !== 'docx') {
    return { blocked: 'unsupported-format', reason: UNSUPPORTED_FORMAT_REASON };
  }

  const over = [a, b].filter(s => exceedsDocxByteCap(s.byteLength));
  if (over.length > 0) {
    return { blocked: 'cap', reason: describeDocxCapBreach(over) };
  }

  return null;
}

export interface ProposePairingInput {
  sideA: readonly PairingSketch[];
  sideB: readonly PairingSketch[];
  /** ISO-8601, injected by the caller. */
  generatedAt: string;
  /**
   * The profile's compare config, from which the pairing settings are resolved.
   *
   * Optional so a test can propose without assembling one, but production
   * callers pass it: spec §7 requires `compare.batch.pairing.*` to reach the
   * pairing maths, and a review found that although those keys were resolved,
   * nothing consumed the result — the weights, thresholds and epsilon in a
   * profile did nothing whatsoever.
   */
  config?: CompareConfig;
  /**
   * An explicit A→B pairing to use instead of solving for one.
   *
   * Set when the user has re-pointed a row: their choice is used verbatim, and
   * the assignment solver is not consulted at all. A solver that could "improve"
   * a pairing the user just made would be overriding the confirmation this
   * whole feature is built around.
   */
  pairing?: ReadonlyMap<string, string>;
}

/**
 * Turn an explicit A→B map into the same shape `assignPairs` returns.
 *
 * Takes the READABLE documents, for the same reason `assignPairs` does: an
 * unreadable one is not a pairing candidate and gets its own row, so letting it
 * fall through here as well would give it two.
 */
function applyExplicitPairing(
  sideA: readonly PairingSketch[],
  sideB: readonly PairingSketch[],
  pairing: ReadonlyMap<string, string>,
  scores: readonly PairScore[],
): { pairs: readonly ProposedPair[]; unmatchedA: readonly string[]; unmatchedB: readonly string[] } {
  const scoreOf = new Map(scores.map(s => [`${s.a}::${s.b}`, s]));
  const pairs: ProposedPair[] = [];
  const takenB = new Set<string>();
  const unmatchedA: string[] = [];

  for (const a of [...sideA].sort((x, y) => (x.docId < y.docId ? -1 : 1))) {
    const b = pairing.get(a.docId);
    const score = b === undefined ? undefined : scoreOf.get(`${a.docId}::${b}`);
    if (b === undefined || score === undefined || takenB.has(b)) {
      unmatchedA.push(a.docId);
      continue;
    }
    takenB.add(b);
    pairs.push({ a: a.docId, b, score });
  }

  const unmatchedB = [...sideB]
    .sort((x, y) => (x.docId < y.docId ? -1 : 1))
    .map(s => s.docId)
    .filter(id => !takenB.has(id));

  return { pairs, unmatchedA, unmatchedB };
}

/**
 * Decide the band for one proposed pair.
 *
 * `bestRival` is the highest score any OTHER candidate for this document
 * reached — a number, deliberately not a list. It was a list, and the list the
 * call site had to hand was the two rivals the row displays; a review found the
 * margin being measured off that. `rankCandidates` re-orders the leading tie
 * band by the cast, so when three or more candidates are level the true best
 * can be pushed out of the displayed two and the margin measured from a lower
 * number — the cast deciding whether a pairing arrives pre-ticked, which plan
 * 17b guardrail 3 forbids. A number cannot be sliced.
 */
export function bandFor(
  score: PairScore,
  bestRival: number | null,
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): PairingBand {
  // A filename-only score is never enough to pre-tick, however high it is.
  // Filenames are how people track versions and also how people mislabel them,
  // and with no content signal there is nothing to catch the mislabelling.
  if (!score.contentAvailable) return 'propose-with-rivals';

  // No rivals at all means nothing to be confused with, so the margin is as
  // wide as it can be.
  const margin = bestRival === null ? 1 : score.score - bestRival;
  return score.score >= settings.autoPairScore && margin >= settings.autoPairMargin
    ? 'auto-pair'
    : 'propose-with-rivals';
}

/**
 * Propose a pairing between two bundles.
 *
 * The result is a proposal and nothing more: no comparison has run, and none
 * may until the user confirms (spec §6.6). That ordering is the whole point of
 * the feature — a mis-pairing produces a complete, plausible, worthless redline
 * that a reader has no cue to distrust.
 */
export function proposePairing(input: ProposePairingInput): PairingProposal {
  const settings = input.config === undefined
    ? DEFAULT_PAIRING_SETTINGS
    : resolvePairingSettings(input.config);

  // A document that could not be read is not a pairing candidate: there is
  // nothing to pair on and nothing to compare. It is kept out of the ASSIGNMENT
  // only — it still gets a document entry and its own row below, because the
  // grid's rule is that every document the user handed us appears exactly once.
  const readableA = input.sideA.filter(s => s.unreadable === undefined);
  const readableB = input.sideB.filter(s => s.unreadable === undefined);
  const unreadable = [...input.sideA, ...input.sideB].filter(s => s.unreadable !== undefined);

  const solved = assignPairs(readableA, readableB, settings);
  const { scores } = solved;

  // A user-supplied pairing is used verbatim; only an unsolicited one is solved.
  const { pairs, unmatchedA, unmatchedB } = input.pairing === undefined
    ? solved
    : applyExplicitPairing(readableA, readableB, input.pairing, scores);

  const sketchOf = new Map<string, PairingSketch>();
  for (const s of [...input.sideA, ...input.sideB]) sketchOf.set(s.docId, s);
  const label = (id: string): string => sketchOf.get(id)?.fileName ?? id;
  const disclosures: string[] = [];

  /**
   * Every candidate for one side-A document, excluding the one it took —
   * ranked, and NOT yet cut to the two the row shows.
   *
   * The cut happens at the row, and only for display. The band is computed
   * from the full list: the displayed two are cast-ordered, so cutting first
   * would let the cast pick which rival the margin is measured against.
   */
  const rivalsFor = (a: string, taken: string): PairScore[] =>
    rankCandidates(scores.filter(s => s.a === a && s.b !== taken), settings);

  /** The strongest rival's score, or `null` when there is no rival at all. */
  const bestRivalOf = (ranked: readonly PairScore[]): number | null =>
    ranked.reduce<number | null>((n, r) => (n === null || r.score > n ? r.score : n), null);

  const documents: PairingDocument[] = [...input.sideA, ...input.sideB]
    .map(s => ({
      docId: s.docId,
      side: s.side,
      fileName: s.fileName,
      sourceKind: s.sourceKind,
      ...(s.unreadable === undefined ? {} : { unreadable: s.unreadable }),
    }))
    .sort((x, y) => (x.docId < y.docId ? -1 : x.docId > y.docId ? 1 : 0));

  /**
   * Why a document ended up with no counterpart.
   *
   * "Added, or replacing" is a claim about the bundle, and it is one the
   * algorithm cannot make about a document whose content it could not read: a
   * `.pdf`'s text is deliberately not extracted for pairing, so all it had to
   * go on was the name. Saying "no counterpart" there asserts an absence it has
   * no evidence for — the same defect as calling two PDFs "cross-format", one
   * layer along.
   */
  const unmatchedReason = (docId: string, side: 'a' | 'b'): string => {
    const sk = sketchOf.get(docId);
    // The cap first. An over-cap document is deliberately never read (§4.4), so
    // it reaches here looking exactly like an unreadable one — and telling a
    // user their perfectly good file could not be read sends them looking for
    // corruption that is not there.
    if (sk !== undefined && exceedsDocxByteCap(sk.byteLength)) {
      if (!disclosures.includes(CAP_DISCLOSURE)) disclosures.push(CAP_DISCLOSURE);
      return describeUnpairedDocxCapBreach(sk);
    }
    if (sk !== undefined && sk.paraHashes.size === 0) {
      return 'Its content could not be read for pairing, and no filename in the other '
        + 'bundle was close enough to propose a counterpart.';
    }
    return side === 'a'
      ? 'No counterpart in the other bundle — removed, or replaced.'
      : 'No counterpart in the other bundle — added, or replacing.';
  };

  const rows: PairingRow[] = [];

  for (const pair of pairs) {
    const ranked = rivalsFor(pair.a, pair.b);
    const rivals = ranked.slice(0, RIVALS_SHOWN);
    // A pair that cannot be compared is still PAIRED, still scored and still
    // shown with its rivals (acceptance M8). Dropping it would hide a document
    // the user handed us behind a limitation of this build — the pairing
    // analogue of hiding a change.
    const block = blockFor(
      sketchOf.get(pair.a) as PairingSketch,
      sketchOf.get(pair.b) as PairingSketch,
    );
    if (block?.blocked === 'cross-format' && !disclosures.includes(CROSS_FORMAT_DISCLOSURE)) {
      disclosures.push(CROSS_FORMAT_DISCLOSURE);
    }
    if (
      block?.blocked === 'unsupported-format'
      && !disclosures.includes(UNSUPPORTED_FORMAT_DISCLOSURE)
    ) {
      disclosures.push(UNSUPPORTED_FORMAT_DISCLOSURE);
    }
    if (block?.blocked === 'cap' && !disclosures.includes(CAP_DISCLOSURE)) {
      disclosures.push(CAP_DISCLOSURE);
    }

    rows.push({
      id: `${pair.a}::${pair.b}`,
      a: pair.a,
      b: pair.b,
      label: `${label(pair.a)} ▸ ${label(pair.b)}`,
      band: bandFor(pair.score, bestRivalOf(ranked), settings),
      score: pair.score,
      rivals,
      runnable: block === null,
      ...(block === null ? {} : { blocked: block.blocked, reason: block.reason }),
    });
  }

  // An unmatched document is a real answer — a schedule removed, a side letter
  // added — so it gets its own row rather than being left out of the grid. Its
  // rivals are shown too: they are the pairings the solver considered and
  // rejected, which is what makes "unmatched" checkable rather than asserted.
  for (const a of unmatchedA) {
    rows.push({
      id: `${a}::`,
      a,
      b: null,
      label: label(a),
      band: 'unmatched',
      score: null,
      rivals: rankCandidates(scores.filter(s => s.a === a), settings).slice(0, RIVALS_SHOWN),
      runnable: false,
      reason: unmatchedReason(a, 'a'),
    });
  }

  for (const sk of unreadable) {
    rows.push({
      id: sk.side === 'a' ? `${sk.docId}::` : `::${sk.docId}`,
      a: sk.side === 'a' ? sk.docId : null,
      b: sk.side === 'b' ? sk.docId : null,
      label: sk.fileName,
      band: 'unmatched',
      score: null,
      rivals: [],
      runnable: false,
      reason: sk.unreadable as string,
    });
    if (!disclosures.includes(UNREADABLE_DISCLOSURE)) disclosures.push(UNREADABLE_DISCLOSURE);
  }

  for (const b of unmatchedB) {
    rows.push({
      id: `::${b}`,
      a: null,
      b,
      label: label(b),
      band: 'unmatched',
      score: null,
      rivals: rankCandidates(scores.filter(s => s.b === b), settings).slice(0, RIVALS_SHOWN),
      runnable: false,
      reason: unmatchedReason(b, 'b'),
    });
  }

  return {
    kind: 'compare-pairing',
    generatedAt: input.generatedAt,
    rows,
    documents,
    totals: tallyPairing(rows),
    disclosures,
  };
}


/**
 * Re-point one row at a different side-B document, or at nothing.
 *
 * Pairing is one-to-one, so a re-point is never a local edit: taking a document
 * that another row already holds must free that other row, and the document
 * this row was holding must become unmatched rather than vanish. Doing this by
 * mutating a single row is how a grid ends up showing one document twice and
 * another not at all — and since every row still looks plausible, nothing would
 * announce it.
 *
 * Rebuilt from the resulting pair set rather than patched in place, so the
 * bands, rivals, totals and blocked reasons are all recomputed from the same
 * code that produced them originally.
 */
export function repointPairing(
  input: ProposePairingInput,
  current: PairingProposal,
  rowId: string,
  b: string | null,
): PairingProposal {
  const row = current.rows.find(r => r.id === rowId);
  if (row === undefined || row.a === null) return current;

  // The pairing as it currently stands, as an A→B map.
  const chosen = new Map<string, string>();
  for (const r of current.rows) {
    if (r.a !== null && r.b !== null) chosen.set(r.a, r.b);
  }

  if (b === null) chosen.delete(row.a);
  else {
    // Whoever else held `b` loses it — one-to-one, enforced here rather than
    // hoped for.
    for (const [a, held] of [...chosen]) if (held === b && a !== row.a) chosen.delete(a);
    chosen.set(row.a, b);
  }

  return proposePairing({ ...input, pairing: chosen });
}
