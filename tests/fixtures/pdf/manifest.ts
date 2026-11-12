/**
 * Manifest of the committed PDF fixtures and the
 * P-test each one backs. Single source of truth shared by the A0 presence test
 * and the later P1–P13 integration tests, so a renamed/removed fixture fails
 * loudly rather than silently skipping a test.
 *
 * Regenerate the bytes with:
 *   node --experimental-strip-types tests/fixtures/pdf/generate.ts
 */
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const PDF_FIXTURE_DIR = dirname(fileURLToPath(import.meta.url));

/** Absolute path to a committed fixture by file name. */
export function pdfFixture(name: string): string {
  return join(PDF_FIXTURE_DIR, name);
}

/** Every committed fixture file, with the acceptance test(s) it supports. */
export const PDF_FIXTURES = {
  'single-column.base.pdf': 'base body for P1/P2/P4/P6',
  'single-column.insert.pdf': 'P1 — inserted paragraph mid-document',
  'single-column.fontchange.pdf': 'P2 — identical text, bold/red one paragraph',
  'single-column.shall-may.pdf': 'P4 — obligation language shall→may',
  'single-column.amount.pdf': 'P6 — amount S$1,250,000→S$1,500,000',
  'renumber.base.pdf': 'P3 — renumbering cascade (original)',
  'renumber.revised.pdf': 'P3 — renumbering cascade (revised)',
  'two-column.pdf': 'P5 — multi-column reading order',
  'image-only.pdf': 'P7 — no extractable text (vector only)',
  'ocr-layer.pdf': 'P7b — invisible OCR text layer over a raster',
  'near-garbled.pdf': 'P9 — high out-of-printable-range char ratio (approx.)',
  'soft-hyphen-dash.pdf': 'P10 — soft hyphen + en/em dash normalisation',
  'justified-kerned.pdf': 'P11 — fragmented justified line items',
  'fee-table.base.pdf': 'P12 — fee table (original)',
  'fee-table.revised.pdf': 'P12 — fee table, one changed cell',
  'signature-image.pdf': 'P13 — signature-image page',
  'repeated-logo.pdf': 'P13 — repeated letterhead logo (dedupe)',
  'word-pdf-pair.pdf': 'Phase E — Word-vs-PDF pair (PDF side)',
  'word-pdf-pair.docx': 'Phase E — Word-vs-PDF pair (Word side)',
} as const;

export type PdfFixtureName = keyof typeof PDF_FIXTURES;
