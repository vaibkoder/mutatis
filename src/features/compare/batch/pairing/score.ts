/**
 * Pair scoring (ADR-0037; spec §6.3, §6.4) — turning three signals into one
 * number a user can argue with.
 *
 * The weighted sum is content and filename only:
 *
 *     score = 0.80 · contentJaccard + 0.20 · filenameSimilarity
 *
 * **The cast signal is deliberately not in it.** Inside one deal bundle the
 * SPA, the disclosure letter and the side letter share the same parties and
 * most defined terms, so cast overlap is near 1 for every candidate pair and
 * discriminates almost nothing. A signal that is near-constant across the whole
 * matrix cannot inform a choice, but it can still *move* one: folded into the
 * sum at any weight it would shift scores across the auto-pair boundary while
 * carrying no information about which pairing is right. So it is consulted only
 * where the two signals that do discriminate have already come out level —
 * within the configured tie epsilon — and there it decides nothing except an
 * ordering the user is shown and can override.
 *
 * Every field of a `PairScore` is kept, not just the total. The grid shows the
 * parts because "0.86" is not a reason and "0.91 of the paragraphs match, and
 * the filenames agree" is (ADR-0008: no opaque judgement about a document).
 *
 * ## Absent evidence is not evidence of absence
 *
 * A document with no paragraph hashes — a `.pdf`, whose text is deliberately
 * not extracted for pairing (`sketch.ts`), or a `.docx` with nothing long
 * enough to be sketch-worthy — yields a content Jaccard of 0. Feeding that 0
 * into the weighted sum would score it as *positively dissimilar*, capping the
 * pair at 0.2 and pushing it below the no-match threshold: a Word and a PDF of
 * the very same agreement could never be paired at all, and the grid would show
 * two unmatched documents rather than the cross-format pair the user is entitled
 * to see (acceptance M8).
 *
 * So when the content signal is *unavailable* the score falls back onto the
 * filename, and records that it did.
 *
 * **But it falls back into a ceiling, not onto the full range.** The first
 * version of this renormalised outright — filename 1.0 meant score 1.0 — and a
 * review found what that does on the very bundle it was written for: a `.pdf`
 * whose name matched perfectly outscored every content-verified pair and
 * displaced the genuine `.docx` counterpart, which was then shown as "no
 * counterpart in the other bundle". Nothing was hidden — both rows were on the
 * grid with their rivals — but the proposal was wrong, and wrong in the
 * confident direction.
 *
 * The fallback is therefore capped at
 * `PAIRING_NO_MATCH_SCORE + PAIRING_NO_CONTENT_HEADROOM`, so wherever a
 * filename-only score is COMPARED against a corroborated one it loses: in a
 * rival list, and in `bandFor`'s margin — where an uncapped PDF rival would
 * have stopped a perfect pair auto-pairing merely by being in the bundle.
 *
 * Whether such a pair FORMS is a separate question, and it is answered
 * structurally rather than by this number: `assign.ts` runs a second assignment
 * phase over what the first could not place (spec §6.4). Trying to answer both
 * with one number is what failed twice — pitched high, the score displaced real
 * pairs; pitched low enough not to, it stopped cross-format pairs forming at all
 * unless their filenames were near-identical, which is a different way of
 * failing acceptance M8.
 *
 * Such a pair can never auto-pair (`propose.ts`) either: filenames are how
 * people track versions and also how people mislabel them, and one weak signal
 * is never enough to pre-tick.
 *
 * Pure — no Office JS, no DOM, no network.
 */
import { DEFAULT_PAIRING_SETTINGS, type PairingSettings } from '../config.js';
import { castOverlap, filenameSimilarity, jaccard } from './signals.js';
import type { PairingSketch } from './sketch.js';

/** One candidate pairing, with its workings left in. */
export interface PairScore {
  /** `docId` of the side-A document. */
  a: string;
  /** `docId` of the side-B document. */
  b: string;
  /** Jaccard over sketch paragraph hashes — the dominant signal. */
  content: number;
  /** Dice over version-token-stripped filename trigrams. */
  filename: number;
  /** Defined-term + party overlap. Tie-break only; NOT in `score`. */
  cast: number;
  /**
   * Whether both sides had paragraph hashes to compare at all.
   *
   * `false` means `content` is 0 because there was nothing to measure, not
   * because the documents differ — and the score is filename-only. Carried
   * rather than inferred so the grid can say which it is.
   */
  contentAvailable: boolean;
  /** The weighted sum. 0–1. */
  score: number;
}

/** Score one candidate pairing. Symmetric in its two arguments. */
export function scorePair(
  a: PairingSketch,
  b: PairingSketch,
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): PairScore {
  const contentAvailable = a.paraHashes.size > 0 && b.paraHashes.size > 0;
  const content = jaccard(a.paraHashes, b.paraHashes);
  const filename = filenameSimilarity(a.fileName, b.fileName);
  return {
    a: a.docId,
    b: b.docId,
    content,
    filename,
    cast: castOverlap(a, b),
    contentAvailable,
    // Capped fallback, not zero-weighted and not a full renormalisation:
    // see the module comment.
    score: contentAvailable
      ? settings.weightContent * content + settings.weightFilename * filename
      : settings.noContentCeiling * filename,
  };
}

/** Every candidate pairing between two bundles, unranked and unfiltered. */
export function scoreAll(
  sideA: readonly PairingSketch[],
  sideB: readonly PairingSketch[],
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): PairScore[] {
  const out: PairScore[] = [];
  for (const a of sideA) for (const b of sideB) out.push(scorePair(a, b, settings));
  return out;
}

/**
 * Whether two candidates are level on the signals that discriminate.
 *
 * Exported because the tie rule is a claim the tests need to make directly:
 * "the cast was consulted here, and only here".
 */
export function isTie(
  left: PairScore,
  right: PairScore,
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): boolean {
  return Math.abs(left.score - right.score) < settings.tieEpsilon;
}

/** Score desc, then `docId` — a total order, independent of input order. */
function byScore(left: PairScore, right: PairScore): number {
  if (right.score !== left.score) return right.score - left.score;
  return left.b < right.b ? -1 : left.b > right.b ? 1 : 0;
}

/** Cast desc, then `docId`. Applied only inside the leading tie band. */
function byCast(left: PairScore, right: PairScore): number {
  if (right.cast !== left.cast) return right.cast - left.cast;
  return left.b < right.b ? -1 : left.b > right.b ? 1 : 0;
}

/**
 * Rank candidates best-first, breaking ties — and only ties — with the cast.
 *
 * "Tie" is defined against the BEST score, not pairwise between neighbours.
 * Pairwise would not be transitive: at a tie epsilon of 0.02, scores of
 * 0.10, 0.11 and 0.12 tie as adjacent pairs but not at the ends, and a
 * comparator built on that sorts differently for different input permutations —
 * which is exactly the order-dependence M4 exists to forbid. So the leading
 * band (everything within epsilon of the top score) is identified first and
 * ordered by cast; everything below it is ordered by score alone. Both halves
 * are total orders, so the result is a function of the SET of candidates.
 *
 * Nothing is dropped. This orders the candidates the grid shows; it never
 * decides which of them the user sees, because a confidence signal that removes
 * a candidate from view is the thing ADR-0017 already rejected for alignment
 * confidence and ADR-0037 rejects again here.
 */
export function rankCandidates(
  candidates: readonly PairScore[],
  settings: PairingSettings = DEFAULT_PAIRING_SETTINGS,
): PairScore[] {
  if (candidates.length === 0) return [];

  const ordered = [...candidates].sort(byScore);
  const best = ordered[0] as PairScore;
  const band = ordered.filter(c => best.score - c.score < settings.tieEpsilon);
  if (band.length < 2) return ordered;

  const rest = ordered.slice(band.length);
  return [...band.sort(byCast), ...rest];
}
