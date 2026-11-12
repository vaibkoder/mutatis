/**
 * Hard input caps for Compare at scale (ADR-0037; spec §4.4).
 *
 * Compare's only caps today are the PDF ones (`io/pdf/limits.ts`) and the text
 * paste cap. The `.docx` path has **no size cap at all** — survivable at one
 * pair, and not at N, because a batch multiplies every exposure by the number
 * of variants.
 *
 * Same convention as the PDF caps: named constant, typed error, and a guard
 * that fails BEFORE any partial work. No timeouts — a timeout is
 * non-deterministic and leaves a half-built, untrustworthy batch.
 *
 * A cap REFUSES; it never silently truncates. Quietly dropping the variants
 * past a ceiling would be suppression by another name (ADR-0008) — the user
 * would get a complete-looking batch that is missing documents.
 *
 * Pure — no Office JS, no DOM, no network.
 */

/** Maximum variants in a one-to-many batch. */
export const MAX_BATCH_VARIANTS = 100;

/** Maximum documents per side in a many-to-many bundle. */
export const MAX_BUNDLE_DOCUMENTS = 50;

/** Maximum raw byte size of any single `.docx`. 50 MiB. */
export const MAX_DOCX_BYTES = 50 * 1024 * 1024;

/** Maximum summed byte size across one batch's inputs. 1 GiB. */
export const MAX_BATCH_TOTAL_BYTES = 1024 * 1024 * 1024;

/**
 * Ceiling on cumulative retained artefacts. Past it the runner stops
 * materialising eagerly and flips later rows to on-demand — WITH a disclosure
 * (spec §4.3 step 6). Degrading silently is not an option.
 */
export const MAX_BATCH_ARTEFACT_BYTES = 512 * 1024 * 1024;

/** Which cap was exceeded. */
export type BatchCapKind =
  | 'variants'
  | 'documents'
  | 'docxBytes'
  | 'totalBytes'
  | 'artefactBytes';

/** Thrown when a hard batch input cap is exceeded. No partial batch is produced. */
export class BatchInputTooLargeError extends Error {
  constructor(
    readonly cap: BatchCapKind,
    readonly actual: number,
    readonly limit: number,
    /** The offending document, when the cap is per-document. */
    readonly documentLabel?: string,
  ) {
    const where = documentLabel !== undefined ? ` ("${documentLabel}")` : '';
    super(
      `This batch exceeds the ${cap} cap${where} (${actual} > ${limit}). ` +
        `It was not processed, and no variants were dropped to fit — ` +
        `remove documents and run again.`,
    );
    this.name = 'BatchInputTooLargeError';
  }
}

/** Throw if a one-to-many batch has too many variants. Call before any ingest. */
export function checkVariantCap(
  count: number,
  limit: number = MAX_BATCH_VARIANTS,
): void {
  if (count > limit) throw new BatchInputTooLargeError('variants', count, limit);
}

/** Throw if one side of a many-to-many bundle has too many documents. */
export function checkBundleCap(
  count: number,
  limit: number = MAX_BUNDLE_DOCUMENTS,
): void {
  if (count > limit) throw new BatchInputTooLargeError('documents', count, limit);
}

/** Throw if a single `.docx` is oversize. Call before reading its bytes. */
export function checkDocxByteCap(
  byteLength: number,
  documentLabel: string,
  limit: number = MAX_DOCX_BYTES,
): void {
  if (byteLength > limit) {
    throw new BatchInputTooLargeError('docxBytes', byteLength, limit, documentLabel);
  }
}

/** Throw if the summed input size of a batch is oversize. Call before any ingest. */
export function checkTotalByteCap(
  totalBytes: number,
  limit: number = MAX_BATCH_TOTAL_BYTES,
): void {
  if (totalBytes > limit) {
    throw new BatchInputTooLargeError('totalBytes', totalBytes, limit);
  }
}

/**
 * Whether another artefact of `nextBytes` still fits under the retention
 * ceiling. Unlike the others this does NOT throw: exceeding it is a graceful
 * degradation to on-demand artefacts (with a disclosure), not a refusal — the
 * comparisons themselves are all still run and all still reported.
 */
export function artefactBudgetAllows(
  retainedBytes: number,
  nextBytes: number,
  limit: number = MAX_BATCH_ARTEFACT_BYTES,
): boolean {
  return retainedBytes + nextBytes <= limit;
}

/**
 * Whether one document is over the per-document byte cap — as a question, not
 * a refusal.
 *
 * `checkDocxByteCap` throws, which is right when a batch is being started: an
 * oversize input is refused whole rather than quietly trimmed. But the pairing
 * grid is shown BEFORE anything runs, and a document that will be too big to
 * compare must still appear there, paired, with the cap named. Dropping it from
 * the grid would hide a document the user handed us; throwing would refuse the
 * whole bundle for one bad file.
 */
export function exceedsDocxByteCap(byteLength: number): boolean {
  return byteLength > MAX_DOCX_BYTES;
}

/**
 * The sentence shown for a pair blocked by the per-document cap.
 *
 * Here rather than in `pairing/propose.ts` because two layers now need it and
 * they must not drift: the grid says why a pair will not run, and the runner
 * says why it did not. If those two sentences disagreed, the second would read
 * as a new failure rather than the same one the user already accepted.
 */
export function describeDocxCapBreach(
  over: readonly { fileName: string; byteLength: number }[],
): string {
  return `${namesOverCap(over)}, so this pair was identified but will not be compared.`;
}

/**
 * The same breach, for a document that ended up with no counterpart at all.
 *
 * Split because the paired sentence's tail ("so this pair…") is a lie on an
 * unmatched row, and the reason it would otherwise fall back to — "its content
 * could not be read" — is a different lie: the document is perfectly readable,
 * it was simply never read, because reading it would have been work spent on a
 * comparison the cap had already ruled out (spec §4.4).
 */
export function describeUnpairedDocxCapBreach(
  document: { fileName: string; byteLength: number },
): string {
  return (
    `${namesOverCap([document])}, so it was not read and could not be paired on `
    + `its content.`
  );
}

function namesOverCap(over: readonly { fileName: string; byteLength: number }[]): string {
  const mib = (n: number): string => `${(n / (1024 * 1024)).toFixed(1)} MiB`;
  return (
    `${over.map(s => `"${s.fileName}" (${mib(s.byteLength)})`).join(' and ')} `
    + `exceeds the ${mib(MAX_DOCX_BYTES)} per-document cap`
  );
}
