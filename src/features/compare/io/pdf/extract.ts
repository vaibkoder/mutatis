/**
 * PDF text extraction + silent-corruption guards.
 *
 * Opens a PDF (via an injected, confidentiality-hardened {@link PdfOpener}) and
 * returns per-page text items (str + geometry + hasEOL) for the geometry
 * reconstruction pass (A3). The guards here are the highest-value reliability
 * finding in the plan — they stop a SILENT bad diff:
 *
 *  - P7  NoExtractableTextError — a page (or whole doc) with no extractable text
 *        is image-only/scanned; it is named and NEVER silently diffed as empty.
 *  - P9  SuspectedEncoding — a page whose extracted text is mostly out-of-printable
 *        glyphs (missing ToUnicode) is EXCLUDED from the diff stream and listed,
 *        rather than feeding garbage into the comparison.
 *  - P7b OCR-layer — a page with a real text layer that is also dominated by a
 *        full-page raster is flagged (badge, never suppressed). The image-dominance
 *        signal is supplied by the image-operator pass (D3); {@link flagOcrLayers}
 *        is the pure assessment fed that signal.
 *
 * Hard input caps (limits.ts) fail with a typed error and no partial diff.
 * Pure except for the awaited PDF.js calls; no Office JS, no Date.now().
 */
import { isTextItem, type PdfOpener, type PdfTextItem } from './pdfjs-types.js';
import { checkByteCap, checkPageCap, checkCharCap } from './limits.js';

/** P9 — a page over this out-of-printable ratio is treated as encoding-garbled. */
export const GARBLE_RATIO_THRESHOLD = 0.08;

/** Extracted, trustworthy text for one page (1-based). */
export interface PdfPageText {
  pageNumber: number;
  /** Raw text items in PDF.js array order (incl. empty-str EOL markers). */
  items: PdfTextItem[];
  /** Non-whitespace character count (0 ⇒ no extractable text). */
  charCount: number;
  /** Fraction of non-whitespace chars that are out of printable range. */
  outOfPrintableRatio: number;
}

export interface PdfExtractDiagnostics {
  totalPages: number;
  /** P7 — pages with no extractable text (image-only candidates). */
  emptyPages: number[];
  /** P9 — pages excluded from the diff for suspected encoding garble. */
  encodingSuspectPages: number[];
  /** P7b — pages included but flagged as a suspected OCR text layer. */
  ocrLayerSuspectPages: number[];
}

export interface PdfExtraction {
  /** Pages whose text is trustworthy and fed to reconstruction. */
  pages: PdfPageText[];
  diagnostics: PdfExtractDiagnostics;
}

/**
 * P7 — thrown when a PDF (or every page of it) has no extractable text. The
 * affected page numbers are named so the UI can say exactly which pages could
 * not be compared. This is an explicit can't-compare, NOT a silent drop.
 */
export class NoExtractableTextError extends Error {
  constructor(readonly pageNumbers: number[]) {
    super(
      `No extractable text — cannot compare (appears scanned/image-only; ` +
        `OCR is out of scope). Affected page(s): ${pageNumbers.join(', ')}.`,
    );
    this.name = 'NoExtractableTextError';
  }
}

/** Optional cooperative-cancellation signal (matches the pipeline's shape). */
export interface AbortLike {
  aborted: boolean;
}

/**
 * Extract trustworthy per-page text from raw PDF bytes.
 *
 * @throws PdfInputTooLargeError  on any hard cap (bytes/pages/chars).
 * @throws NoExtractableTextError when no page yields trustworthy text (P7/P9).
 */
export async function extractPdf(
  data: Uint8Array,
  opener: PdfOpener,
  abortSignal?: AbortLike,
): Promise<PdfExtraction> {
  checkByteCap(data.byteLength);
  const doc = await opener(data);
  try {
    checkPageCap(doc.numPages);

    const allPages: PdfPageText[] = [];
    let runningChars = 0;
    for (let p = 1; p <= doc.numPages; p++) {
      if (abortSignal?.aborted) throw new DOMException('PDF extraction cancelled', 'AbortError');
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      const items = content.items.filter(isTextItem);
      const text = items.map((i) => i.str).join('');
      runningChars += text.length;
      checkCharCap(runningChars);
      allPages.push({
        pageNumber: p,
        items,
        charCount: countNonWhitespace(text),
        outOfPrintableRatio: outOfPrintableRatio(text),
      });
    }

    const emptyPages = allPages.filter((pg) => pg.charCount === 0).map((pg) => pg.pageNumber);
    const encodingSuspectPages = allPages
      .filter((pg) => pg.charCount > 0 && pg.outOfPrintableRatio > GARBLE_RATIO_THRESHOLD)
      .map((pg) => pg.pageNumber);

    const suspect = new Set(encodingSuspectPages);
    const trustworthy = allPages.filter((pg) => pg.charCount > 0 && !suspect.has(pg.pageNumber));

    if (trustworthy.length === 0) {
      // Nothing safe to diff: image-only and/or wholly garbled.
      throw new NoExtractableTextError(allPages.map((pg) => pg.pageNumber));
    }

    return {
      pages: trustworthy,
      diagnostics: {
        totalPages: doc.numPages,
        emptyPages,
        encodingSuspectPages,
        ocrLayerSuspectPages: [], // populated by D3 via flagOcrLayers()
      },
    };
  } finally {
    await doc.destroy();
  }
}

/**
 * P7b assessment (pure) — which trustworthy pages are a suspected OCR text layer.
 * A page is OCR-suspect when it has real extractable text AND is dominated by a
 * full-page raster image (the image-dominance set comes from the D3 image pass).
 * Flag only, never a suppression: the page's text still diffs, just badged.
 */
export function flagOcrLayers(
  pages: PdfPageText[],
  imageDominantPageNumbers: ReadonlySet<number>,
): number[] {
  return pages
    .filter((pg) => pg.charCount > 0 && imageDominantPageNumbers.has(pg.pageNumber))
    .map((pg) => pg.pageNumber);
}

// ---------------------------------------------------------------------------
// Encoding-garble detection (P9)
// ---------------------------------------------------------------------------

/**
 * Fraction of NON-WHITESPACE characters that fall out of the printable range.
 * A high ratio signals a missing/broken ToUnicode CMap (the extracted glyphs are
 * control/private-use/replacement code points, i.e. garbage). Legitimate format
 * characters (soft hyphen U+00AD, zero-width joiners) are deliberately treated
 * as printable so a soft-hyphenated document is not mis-flagged.
 */
export function outOfPrintableRatio(text: string): number {
  let total = 0;
  let bad = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (isWhitespace(cp)) continue;
    total++;
    if (isOutOfPrintable(cp)) bad++;
  }
  return total === 0 ? 0 : bad / total;
}

function countNonWhitespace(text: string): number {
  let n = 0;
  for (const ch of text) {
    if (!isWhitespace(ch.codePointAt(0)!)) n++;
  }
  return n;
}

function isWhitespace(cp: number): boolean {
  return cp === 0x20 || cp === 0x09 || cp === 0x0a || cp === 0x0d || cp === 0xa0 || cp === 0x3000;
}

/** True for code points that should never appear in legitimately-decoded text. */
function isOutOfPrintable(cp: number): boolean {
  if (cp === 0x09 || cp === 0x0a || cp === 0x0d) return false; // tab/newline
  if (cp < 0x20) return true; // C0 control
  if (cp >= 0x7f && cp <= 0x9f) return true; // DEL + C1 control
  if (cp >= 0xd800 && cp <= 0xdfff) return true; // lone surrogate
  if (cp >= 0xe000 && cp <= 0xf8ff) return true; // BMP private use area
  if (cp === 0xfffd) return true; // replacement character
  if (cp >= 0xf0000 && cp <= 0x10ffff) return true; // supplementary private use
  return false;
}
