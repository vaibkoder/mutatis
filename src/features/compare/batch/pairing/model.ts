/**
 * The pairing proposal (ADR-0037; spec §6.5, §6.6) — what the grid renders and
 * what the user confirms.
 *
 * Plain data, and shaped by one rule inherited from `batch/index-model.ts`:
 * **there is no filter parameter anywhere in this module.** A confidence band
 * is a badge. It changes how a row is presented — pre-ticked, flagged for
 * review, or shown as unmatched — and it never removes a row from the grid or a
 * rival from a row. ADR-0017 already rejected a default-on "hide low-confidence
 * differences" filter for the alignment case; this is the same rejection one
 * level up, and for a sharper reason: a pairing the user never saw is a
 * comparison they cannot audit, and a mis-pairing produces a redline that looks
 * entirely normal.
 *
 * Every document from both bundles appears in `rows` exactly once — as one half
 * of a proposed pair, or as an unmatched row of its own. That is the pairing
 * analogue of the batch index's "a pair the batch was asked to run always
 * produces a row".
 *
 * Pure — no Office JS, no DOM, no network.
 */
import type { PairScore } from './score.js';

/**
 * How much scrutiny a proposed pairing needs.
 *
 * - `auto-pair` — high score AND a clear lead over the best rival. Pre-ticked;
 *   still shown, still overridable.
 * - `propose-with-rivals` — a pairing was found, but either it is not strong or
 *   a rival is close behind. Shown with its rivals and their scores, awaiting a
 *   look. Named as the spec and ADR-0037 name it: the rivals are not an extra,
 *   they are what distinguishes this band from `auto-pair`.
 * - `unmatched` — no counterpart worth pairing with. An added or removed
 *   document, which is a real answer and not a failure.
 */
export type PairingBand = 'auto-pair' | 'propose-with-rivals' | 'unmatched';

/**
 * Why a confirmed pair still cannot be compared.
 *
 * - `cross-format` — the two sides are different formats (Word against PDF).
 *   ADR-0017's Phase-E alignment gate was never measured.
 * - `unsupported-format` — the two sides AGREE on a format this build does not
 *   batch, today PDF against PDF. Distinguished from `cross-format` because
 *   telling a user their two PDFs are "cross-format" is simply untrue, and a
 *   reason that is untrue is worse than no reason: it sends them looking for a
 *   mismatch that does not exist.
 * - `cap` — the pair is comparable but one side is over the per-document cap.
 */
export type PairingBlockKind = 'cross-format' | 'unsupported-format' | 'cap';

/** One line of the pairing grid. */
export interface PairingRow {
  /** Stable within the proposal; the grid keys on this, never on a file name. */
  id: string;
  /** Side-A document, or `null` for a side-B document with no counterpart. */
  a: string | null;
  /** Side-B document, or `null` for a side-A document with no counterpart. */
  b: string | null;
  /** Display label — `"A ▸ B"`, or the lone document's own name. */
  label: string;
  band: PairingBand;
  /** The proposed pair's own score. `null` on an unmatched row. */
  score: PairScore | null;
  /**
   * The next-best candidates for this document, best first, WITH their scores.
   * Shown whenever the band is not `auto-pair` — a proposal the user cannot see
   * the alternatives to is one they can only rubber-stamp.
   */
  rivals: readonly PairScore[];
  /**
   * Whether this pair can actually be compared once confirmed. A `false` here
   * is always accompanied by `reason`, and the row is still shown with its
   * pairing — surfaced, never dropped (spec §6.6, acceptance M8).
   */
  runnable: boolean;
  blocked?: PairingBlockKind;
  reason?: string;
}

/**
 * One document in the proposal, by identity rather than by display string.
 *
 * The grid needs names for the re-point control and for rival lists. Carrying
 * them here rather than parsing them back out of a row's `"A ▸ B"` label keeps
 * the label a presentation detail: a file name containing the separator would
 * otherwise silently corrupt every name the grid shows.
 */
export interface PairingDocument {
  docId: string;
  side: 'a' | 'b';
  fileName: string;
  sourceKind: 'docx' | 'pdf' | 'other';
  /**
   * Why this document could not be read, when it could not be.
   *
   * Carried here because the grid needs it: such a document is not a pairing
   * candidate, so it must not appear in a re-point control. A review found it
   * offered — with a percentage — and choosing it silently destroyed a correct
   * pairing, because the assignment never sees these documents and the choice
   * was discarded after the row holding the counterpart had already gone.
   */
  unreadable?: string;
}

/** Derived counts. Never authored by hand, so they cannot disagree with the rows. */
export interface PairingTotals {
  documentsA: number;
  documentsB: number;
  pairs: number;
  autoPair: number;
  needsReview: number;
  unmatchedA: number;
  unmatchedB: number;
  /** Pairs identified but not comparable (cross-format, or over a cap). */
  unrunnable: number;
}

/** The whole proposal — the thing the user confirms before anything runs. */
export interface PairingProposal {
  kind: 'compare-pairing';
  /** Injected by the caller — never `Date.now()` inside the builder. */
  generatedAt: string;
  rows: readonly PairingRow[];
  /** Every document from both bundles, in `docId` order. */
  documents: readonly PairingDocument[];
  totals: PairingTotals;
  /** Blind-spot notices, carried exactly as the batch index carries them. */
  disclosures: readonly string[];
}

/** Tally the rows. Derived from the rows themselves. */
export function tallyPairing(rows: readonly PairingRow[]): PairingTotals {
  const totals: PairingTotals = {
    documentsA: 0,
    documentsB: 0,
    pairs: 0,
    autoPair: 0,
    needsReview: 0,
    unmatchedA: 0,
    unmatchedB: 0,
    unrunnable: 0,
  };

  for (const row of rows) {
    if (row.a !== null) totals.documentsA += 1;
    if (row.b !== null) totals.documentsB += 1;

    if (row.band === 'unmatched') {
      if (row.a !== null) totals.unmatchedA += 1;
      if (row.b !== null) totals.unmatchedB += 1;
      continue;
    }

    totals.pairs += 1;
    if (row.band === 'auto-pair') totals.autoPair += 1;
    else totals.needsReview += 1;
    if (!row.runnable) totals.unrunnable += 1;
  }

  return totals;
}
