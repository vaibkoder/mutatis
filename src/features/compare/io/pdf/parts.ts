/**
 * PDF side loader — the PDF analogue of the Word path's
 * extractParts→preprocessSide: open + extract + reconstruct + adapt to
 * `ParaRecord[]`, ready to feed `runPipelinePdf` (B2). One call per side.
 *
 * It also runs a cheap cross-page header/footer heuristic (a short line repeated
 * on ≥2 pages, or a bare page number) so the pipeline's existing
 * exclude-headers/footers rules apply to PDF bodies too — bounded noise sources
 * ADR-0017 flagged as already-handled.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { ParaRecord } from '../../model.js';

import { extractPdf, type AbortLike, type PdfExtractDiagnostics } from './extract.js';
import type { PdfOpener } from './pdfjs-types.js';
import { reconstructDocument, type ReconstructedParagraph } from './reconstruct.js';
import { pdfParagraphsToRecords } from './pararecord.js';

/** A page number is at most this many words to count as header/footer chrome. */
const HEADER_FOOTER_MAX_WORDS = 8;

export interface PdfSideRecords {
  /** Paragraph records, ready for the diff pipeline. */
  paras: ParaRecord[];
  /** Extraction diagnostics (empty/encoding-suspect pages) for the UI. */
  diagnostics: PdfExtractDiagnostics;
  /** The reconstructed paragraphs (kept for the side-by-side view + alignment). */
  reconstructed: Array<ReconstructedParagraph & { isHeaderFooter: boolean }>;
}

/** Load one PDF side into ParaRecord[] (+ diagnostics). */
export async function extractPdfParaRecords(
  data: Uint8Array,
  opener: PdfOpener,
  config: CompareConfig,
  abortSignal?: AbortLike,
): Promise<PdfSideRecords> {
  const extraction = await extractPdf(data, opener, abortSignal);
  const reconstructed = reconstructDocument(extraction);
  const flagged = flagHeadersFooters(reconstructed);
  const paras: ParaRecord[] = pdfParagraphsToRecords(flagged, config);
  return { paras, diagnostics: extraction.diagnostics, reconstructed: flagged };
}

/**
 * Mark header/footer paragraphs: a short line that repeats across ≥2 pages, or a
 * bare page-number line. Conservative — only short, repeated chrome is flagged,
 * never body prose (which is rarely identical across pages).
 */
export function flagHeadersFooters(
  paras: ReadonlyArray<ReconstructedParagraph>,
): Array<ReconstructedParagraph & { isHeaderFooter: boolean }> {
  // Pages each short text appears on.
  const pagesByText = new Map<string, Set<number>>();
  for (const p of paras) {
    if (wordCount(p.text) > HEADER_FOOTER_MAX_WORDS) continue;
    const key = p.text.trim().toLowerCase();
    const set = pagesByText.get(key) ?? new Set<number>();
    set.add(p.pageNumber);
    pagesByText.set(key, set);
  }
  return paras.map((p) => {
    const key = p.text.trim().toLowerCase();
    const repeated = (pagesByText.get(key)?.size ?? 0) >= 2;
    const isHeaderFooter = repeated || isPageNumber(p.text);
    return { ...p, isHeaderFooter };
  });
}

function wordCount(text: string): number {
  const t = text.trim();
  return t === '' ? 0 : t.split(/\s+/).length;
}

/** A bare page number ("12") or "Page 3 of 10" style chrome. */
function isPageNumber(text: string): boolean {
  const t = text.trim();
  return /^\d{1,4}$/.test(t) || /^page\s+\d+(\s+of\s+\d+)?$/i.test(t);
}
