/**
 * Hard input caps for the PDF compare path.
 *
 * PDFs are adversarial input: a small file can declare millions of pages or
 * decompress to gigabytes of text. We fail with a TYPED error at a pinned
 * ceiling rather than hang — there are NO timeouts (a timeout is non-deterministic
 * and would leave a half-built, untrustworthy diff). All thresholds are named
 * constants here, not magic literals scattered through the extractor.
 *
 * Pure — no Office JS, no pdfjs.
 */

/** Maximum raw PDF byte size accepted (per side). 100 MB. */
export const MAX_PDF_BYTES = 100 * 1024 * 1024;

/** Maximum page count accepted (per document). */
export const MAX_PDF_PAGES = 2000;

/** Maximum total extracted character count accepted (per document). */
export const MAX_PDF_CHARS = 20_000_000;

/** Which cap was exceeded. */
export type PdfCapKind = 'bytes' | 'pages' | 'chars';

/** Thrown when a hard input cap is exceeded. No partial diff is produced. */
export class PdfInputTooLargeError extends Error {
  constructor(
    readonly cap: PdfCapKind,
    readonly actual: number,
    readonly limit: number,
  ) {
    super(
      `PDF exceeds the ${cap} cap (${actual} > ${limit}). ` +
        `The document is too large to compare safely and was not processed. ` +
        `No partial comparison is produced.`,
    );
    this.name = 'PdfInputTooLargeError';
  }
}

/** Throw if the raw byte size exceeds {@link MAX_PDF_BYTES}. Call before opening. */
export function checkByteCap(byteLength: number): void {
  if (byteLength > MAX_PDF_BYTES) {
    throw new PdfInputTooLargeError('bytes', byteLength, MAX_PDF_BYTES);
  }
}

/** Throw if the page count exceeds {@link MAX_PDF_PAGES}. Call right after open. */
export function checkPageCap(numPages: number): void {
  if (numPages > MAX_PDF_PAGES) {
    throw new PdfInputTooLargeError('pages', numPages, MAX_PDF_PAGES);
  }
}

/** Throw if the running character total exceeds {@link MAX_PDF_CHARS}. */
export function checkCharCap(totalChars: number): void {
  if (totalChars > MAX_PDF_CHARS) {
    throw new PdfInputTooLargeError('chars', totalChars, MAX_PDF_CHARS);
  }
}
