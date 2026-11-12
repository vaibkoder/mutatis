/**
 * Named constants for Compare at scale (ADR-0037; spec/features/17-compare-batch.md).
 *
 * Everything tunable lives here so nothing downstream carries a magic literal —
 * the same rule feature 10 §7 sets for the PDF and text paths. The pairing
 * constants are declared now and consumed by many-to-many;
 * declaring them together keeps the two features' knobs in one readable place.
 *
 * Pure — no Office JS, no DOM, no network.
 */

// ---------------------------------------------------------------------------
// Worker pool (§4.2)
// ---------------------------------------------------------------------------

/**
 * Upper bound on pooled workers, whatever the machine reports. Compare is
 * memory-bound before it is CPU-bound — each in-flight pair holds two
 * document-sized `ParaRecord[]` — so more workers past this buys latency at the
 * price of the thing that actually breaks first (§4.3).
 */
export const BATCH_POOL_MAX = 4;

/** Fallback when `navigator.hardwareConcurrency` is unavailable. */
export const BATCH_POOL_DEFAULT_CONCURRENCY = 2;

/**
 * Pool size for a host reporting `hardwareConcurrency` cores. One core is left
 * for the UI thread: a batch that pegs every core makes its own progress bar
 * stutter, which reads as a hang.
 */
export function resolvePoolSize(
  hardwareConcurrency: number | undefined,
  override = 0,
): number {
  if (override > 0) return Math.min(override, BATCH_POOL_MAX);
  const cores = hardwareConcurrency ?? BATCH_POOL_DEFAULT_CONCURRENCY;
  return Math.max(1, Math.min(cores - 1, BATCH_POOL_MAX));
}

// ---------------------------------------------------------------------------
// Pairing (§6) — declared here, consumed by many-to-many
// ---------------------------------------------------------------------------

/** Weight of the content signal (Jaccard over paragraph hashes) — dominant. */
export const PAIRING_WEIGHT_CONTENT = 0.8;

/** Weight of the filename signal (Dice over version-token-stripped names). */
export const PAIRING_WEIGHT_FILENAME = 0.2;

/**
 * Two candidates whose scores differ by less than this are a tie, and only then
 * is the defined-term / party-cast signal consulted. That signal is deliberately
 * NOT in the weighted sum: inside one deal bundle the SPA, the disclosure letter
 * and the side letter share the same cast and most defined terms, so it barely
 * discriminates and must never move a pair across a band boundary.
 */
export const PAIRING_TIE_EPSILON = 0.02;

/** A pair at or above this score, with enough margin, is proposed pre-ticked. */
export const PAIRING_AUTO_PAIR_SCORE = 0.85;

/** Required lead over the best rival before a pair is treated as unambiguous. */
export const PAIRING_AUTO_PAIR_MARGIN = 0.15;

/**
 * The price of going unmatched. A document whose best real partner scores below
 * this is assigned to a padding dummy and falls out as added/removed — which is
 * the right answer when one bundle gained a side letter.
 */
export const PAIRING_NO_MATCH_SCORE = 0.25;

/**
 * How far above the no-match threshold a filename-only score may reach.
 *
 * A pair with no content evidence at all — a `.pdf`, or a document with nothing
 * long enough to sketch — is scored on its filename alone. That score has to
 * clear `PAIRING_NO_MATCH_SCORE`, or such a pair could never form and the grid
 * could never show it (acceptance M8). But it must NOT be able to outrank a
 * pairing that content actually corroborates, and an unbounded renormalisation
 * did exactly that: a PDF whose name matched perfectly scored 1.0, won the
 * global assignment, and displaced the genuine `.docx` counterpart — which was
 * then labelled "no counterpart in the other bundle", a statement that was
 * simply false.
 *
 * A small headroom gives filename-only evidence the meaning it should have:
 * enough to beat *nothing*, never enough to beat *something*. The ceiling is
 * `PAIRING_NO_MATCH_SCORE + PAIRING_NO_CONTENT_HEADROOM`, which a name has to
 * be a near-perfect match to reach.
 */
export const PAIRING_NO_CONTENT_HEADROOM = 0.05;

/**
 * Filename similarity below which two documents with NO content evidence stay
 * unmatched rather than pairing on their names.
 *
 * This is the threshold for the second assignment phase (§6.4), where the only
 * candidates left are documents the first phase could not place and whose
 * content could not be read at all. They compete with nothing but each other,
 * so the bar is a plain question about the names — is this plausibly the same
 * document? — and not a comparison against corroborated evidence.
 *
 * Set LOW on purpose — this is "is there any name evidence at all?", not "is
 * this a good match?". Two things justify the low bar:
 *
 *  - Everything in this phase is already unpairable by content, so a pairing
 *    here takes nothing away from a better-evidenced one, and it can never be
 *    pre-ticked (`bandFor` refuses to auto-pair without content).
 *  - The alternative is worse than a weak proposal. Reporting a document as
 *    "no counterpart in the other bundle — added, or replacing" when a
 *    plausible candidate is sitting in the other bundle is an assertion the
 *    algorithm cannot support, and the spec's own motivating example is exactly
 *    that shape: `Project_Falcon_SPA_v7_exec.docx` against
 *    `SPA (clean) 2026-08-30.pdf` normalises to "project falcon spa" against
 *    "spa" — a Dice of 0.26, weak but not nothing.
 *
 * An unrelated name still scores 0 ("project falcon spa" against "board minutes
 * march"), so the bar discriminates where it needs to.
 */
export const PAIRING_NO_CONTENT_MATCH_SCORE = 0.2;

/**
 * Paragraphs shorter than this (in content tokens) are excluded from a pairing
 * sketch. Empty paragraphs, one-word headings and numbering stubs are identical
 * across every document in a bundle and would pull every Jaccard score toward
 * every other. Matches the absorbed-paragraph guard in ADR-0032.
 */
export const PAIRING_MIN_SKETCH_TOKENS = 4;

/**
 * v1 compares `.docx` against `.docx` only.
 *
 * A mixed pair is surfaced as a real row — identified, paired, scored — and
 * marked not-comparable with this reason. Never silently skipped: ADR-0017's
 * Phase-E alignment gate for Word-vs-PDF was never measured, so shipping the
 * comparison through a batch would be shipping it ungated, while dropping the
 * pair from the grid would hide a document the user handed us.
 *
 * Lives here rather than in `runner.ts` because both the runner and the pairing
 * proposal need it, and pairing must not import the runner.
 */
export const CROSS_FORMAT_REASON =
  'Cross-format comparison is out of scope, so this pair was identified but not compared.';

/**
 * Why a pair whose two sides share a format this build cannot batch is shown
 * but not run — today, PDF against PDF.
 *
 * A separate sentence from `CROSS_FORMAT_REASON` because the two say different
 * things and only one of them is true of two PDFs. The earlier wording called
 * such a pair "cross-format", which is a statement about a mismatch that does
 * not exist and would send a reader looking for one. PDF-vs-PDF comparison
 * itself is built (feature 12) — what is not in v1 is batching it (spec §2.2).
 */
export const UNSUPPORTED_FORMAT_REASON =
  'Only Word documents can be compared in a batch, so this pair was identified but not compared.';
