import { extractPdf, NoExtractableTextError } from '@features/compare/io/pdf/extract.js';
import { PdfInputTooLargeError } from '@features/compare/io/pdf/limits.js';
import { reconstructDocument } from '@features/compare/io/pdf/reconstruct.js';
import { RECONSTRUCT_CONFIG } from '@features/compare/io/pdf/reconstruct.config.js';
import type { PdfExtraction } from '@features/compare/io/pdf/extract.js';
import { flagHeadersFooters } from '@features/compare/io/pdf/parts.js';
import type { PdfOpener } from '@features/compare/io/pdf/pdfjs-types.js';
import type { Block } from '../paste/model.js';

/**
 * A PDF, read into paragraphs.
 *
 * A PDF stores where each piece of text sits on the page, not paragraphs, so the
 * engine's reader rebuilds them from geometry (line clustering, column gutters,
 * paragraph gaps). That is a reconstruction and the user is told so: `notes` is
 * shown above the redline for any comparison that involved a PDF.
 *
 * The engine's own guards are kept as they are: no extractable text (a scan) and
 * the size caps refuse outright rather than produce an empty or partial
 * comparison, and pages whose text is garbled are excluded and NAMED.
 */

export interface PdfReading {
  readonly blocks: Block[];
  readonly notes: string[];
}

const LIGATURES: ReadonlyArray<readonly [string, string]> = [
  ['ﬀ', 'ff'],
  ['ﬁ', 'fi'],
  ['ﬂ', 'fl'],
  ['ﬃ', 'ffi'],
  ['ﬄ', 'ffl'],
  ['ﬅ', 'st'],
  ['ﬆ', 'st'],
];

/**
 * Typesetters fuse letter pairs into one glyph, and PDF.js reports the fused
 * character. Compared with a Word file's "financial", "ﬁnancial" would be a
 * change on every such word.
 */
export function normaliseLigatures(text: string): string {
  let out = text.replace(/­/g, '');
  for (const [from, to] of LIGATURES) out = out.split(from).join(to);
  return out;
}

const PAGE_NUMBER = /^(\d{1,4}|page\s+\d+(\s+of\s+\d+)?|\d+\s*(\/|of)\s*\d+)$/i;


/**
 * The engine starts a new paragraph when the gap between two lines exceeds a fixed
 * multiple of the font size (1.4). That suits single-spaced text and fails on
 * the 1.5- and double-spaced documents lawyers actually send: every line becomes
 * its own paragraph. So the threshold follows the document: it is set a little
 * above the document's own line pitch, and never below the engine's default.
 * A paragraph break is then a gap clearly bigger than an ordinary line.
 */
export function paragraphGapFactor(extraction: PdfExtraction): number {
  const ratios: number[] = [];
  for (const page of extraction.pages) {
    const sizes = page.items
      .filter((i) => i.str.trim() !== '')
      .map((i) => Math.hypot(i.transform[0] ?? 0, i.transform[1] ?? 0));
    const fontSize = median(sizes);
    if (fontSize <= 0) continue;
    const ys = [...new Set(page.items.filter((i) => i.str.trim() !== '').map((i) => Math.round((i.transform[5] ?? 0) * 2) / 2))]
      .sort((a, b) => b - a);
    for (let i = 1; i < ys.length; i++) {
      const gap = (ys[i - 1] ?? 0) - (ys[i] ?? 0);
      if (gap > 0.9 * fontSize && gap < 2.6 * fontSize) ratios.push(gap / fontSize);
    }
  }
  if (ratios.length < 5) return RECONSTRUCT_CONFIG.PARAGRAPH_GAP_FACTOR;
  // The line pitch is the smallest gap that is common; larger gaps are paragraph
  // breaks and headings, so take a low percentile, not the median.
  return Math.min(3, Math.max(RECONSTRUCT_CONFIG.PARAGRAPH_GAP_FACTOR, 1.1 * percentile(ratios, 0.25)));
}

function percentile(values: readonly number[], p: number): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * p))] ?? 0;
}

const median = (values: readonly number[]): number => percentile(values, 0.5);

function pageList(pages: readonly number[]): string {
  const shown = pages.slice(0, 8).join(', ');
  return pages.length > 8 ? `${shown} and ${pages.length - 8} more` : shown;
}

export async function readPdf(bytes: Uint8Array, opener: PdfOpener): Promise<PdfReading> {
  const extraction = await extractPdf(bytes, opener);
  // The engine types its tuning constants as literals; widen the one we adapt.
  const config = {
    ...RECONSTRUCT_CONFIG,
    PARAGRAPH_GAP_FACTOR: paragraphGapFactor(extraction),
  } as unknown as typeof RECONSTRUCT_CONFIG;
  const paragraphs = reconstructDocument(extraction, config);
  const flagged = flagHeadersFooters(paragraphs);

  // A bare page number is dropped only when it is the first or last thing on its
  // page — where a header or footer sits. A lone number in the middle of a page
  // is body text (a table value, a list number) and stays.
  const firstLast = new Map<number, { first: number; last: number }>();
  flagged.forEach((p, i) => {
    const entry = firstLast.get(p.pageNumber);
    if (entry) entry.last = i;
    else firstLast.set(p.pageNumber, { first: i, last: i });
  });

  const blocks: Block[] = [];
  let pageNumbersLeftOut = 0;
  let lowConfidence = 0;
  flagged.forEach((p, i) => {
    const edge = firstLast.get(p.pageNumber);
    const atEdge = edge !== undefined && (edge.first === i || edge.last === i);
    if (atEdge && p.isHeaderFooter && PAGE_NUMBER.test(p.text.trim())) {
      pageNumbersLeftOut += 1;
      return;
    }
    if (p.lowConfidence) lowConfidence += 1;
    blocks.push({ kind: 'p', runs: [{ text: normaliseLigatures(p.text) }] });
  });

  const { diagnostics } = extraction;
  const notes = [
    'Read from a PDF. A PDF has no real paragraphs, so they are rebuilt from the page layout: ' +
      'paragraph breaks, columns, tables and numbering can differ from the original. ' +
      'Images and scanned pages are not compared, and bold, italic and fonts are not carried. ' +
      'Treat this as a guide to the text, not as a Word-quality redline.',
  ];
  if (pageNumbersLeftOut > 0) {
    notes.push(
      `${pageNumbersLeftOut} page ${pageNumbersLeftOut === 1 ? 'number' : 'numbers'} at the top or bottom of pages ` +
        'left out. Repeated header and footer text is kept.',
    );
  }
  if (diagnostics.emptyPages.length > 0) {
    notes.push(
      `Not compared — no selectable text (scanned or image-only), page ${pageList(diagnostics.emptyPages)}.`,
    );
  }
  if (diagnostics.encodingSuspectPages.length > 0) {
    notes.push(
      `Not compared — the text on page ${pageList(diagnostics.encodingSuspectPages)} could not be read reliably.`,
    );
  }
  if (lowConfidence > 0) {
    notes.push(
      `${lowConfidence} ${lowConfidence === 1 ? 'paragraph sits' : 'paragraphs sit'} on multi-column pages ` +
        'and may be out of reading order.',
    );
  }
  return { blocks, notes };
}

/** Plain-English reasons for a PDF that cannot be read. */
export function explainPdfError(error: unknown): string {
  if (error instanceof NoExtractableTextError) {
    const pages = error.pageNumbers;
    return (
      'no selectable text: it looks scanned or image-only' +
      (pages.length > 0 ? ` (page${pages.length === 1 ? '' : 's'} ${pageList(pages)})` : '') +
      '. Mutatis cannot read scanned PDFs; use the original Word file or run OCR first'
    );
  }
  if (error instanceof PdfInputTooLargeError) {
    return error.message.replace(/\s+No partial comparison is produced\.$/, '');
  }
  const name = (error as { name?: string } | null)?.name;
  if (name === 'PasswordException') return 'password-protected; remove the password first';
  if (name === 'InvalidPDFException') return 'not a valid PDF';
  return `could not be read as a PDF (${error instanceof Error ? error.message : String(error)})`;
}
