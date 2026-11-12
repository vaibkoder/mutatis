/**
 * Shared Compare blind-spot safety copy.
 *
 * SINGLE SOURCE for the "what is NOT compared" disclosure and the per-side
 * image-presence notice lines, imported by BOTH the in-pane PdfComparePanel AND
 * the export report-model. Extracting it here guarantees the downloadable
 * artefact and the on-screen panel can never drift (ADR-0020 §2; the disclosures
 * the panel shows must appear verbatim in every exported artefact).
 *
 * Pure data + pure helpers — no Preact, no Office JS.
 */
import type { ImagePresenceReport } from '../io/pdf/images.js';
import type { PdfImageNotice } from '../ui-types.js';

/** Heading for the always-visible blind-spot disclosure. */
export const PDF_NOT_COMPARED_TITLE = 'What is not compared';

/**
 * One disclosure bullet. `lead` (when present) renders as a bold lead sentence
 * in the panel; the export joins `lead + ' ' + rest` into one plain string.
 */
export interface DisclosureItem {
  /** Emphasised lead sentence (rendered bold on screen). */
  lead?: string;
  /** Remainder of the bullet. */
  rest: string;
}

/** The PDF-mode "what is NOT compared" bullets (safety-critical copy). */
export const PDF_NOT_COMPARED_ITEMS: readonly DisclosureItem[] = [
  {
    lead: 'Images, signatures, seals and stamps are not compared.',
    rest: 'Check execution pages and any signed or sealed content visually.',
  },
  { rest: 'Visual formatting (fonts, colours, spacing, layout) is not compared.' },
  { rest: 'Scanned or image-only pages are reported, not diffed — there is no text to read.' },
  { rest: 'Tables are matched approximately and flagged low-confidence; verify cell changes.' },
];

/** Flatten a disclosure bullet into a single plain string (for the export). */
export function disclosureText(item: DisclosureItem): string {
  return item.lead ? `${item.lead} ${item.rest}` : item.rest;
}

/** The PDF "what is NOT compared" bullets as plain strings (for the export). */
export function pdfNotComparedLines(): string[] {
  return PDF_NOT_COMPARED_ITEMS.map(disclosureText);
}

/** One side's image-notice line, or null when that side has no content images. */
export function imageNoticeLine(label: string, report: ImagePresenceReport): string | null {
  const pages = report.pagesWithContentImages;
  if (pages.length === 0) return null;
  const noun = pages.length === 1 ? 'page' : 'pages';
  return `${label}: ${pages.length} ${noun} contain image content not text-compared (page ${pages.join(', ')}).`;
}

/** Both sides' image-notice lines (empty array when neither side carries content images). */
export function imageNoticeLines(notice: PdfImageNotice): string[] {
  return [imageNoticeLine('Original', notice.original), imageNoticeLine('Revised', notice.revised)].filter(
    (l): l is string => l !== null,
  );
}
