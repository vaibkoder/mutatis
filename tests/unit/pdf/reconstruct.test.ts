/**
 * A3 — geometry paragraph reconstruction, against the committed fixtures and a
 * synthetic rotated-item page.
 *
 * Covers: single-column → correct paragraph array + heading; two-column →
 * columns ordered AND no cross-column token leak (P5); justified/kerned → one
 * block per line (P11); rotated text filtered.
 */
import { describe, it, expect } from 'vitest';

import { extractPdf, type PdfPageText } from '../../../src/features/compare/io/pdf/extract.js';
import { reconstructDocument, reconstructPage } from '../../../src/features/compare/io/pdf/reconstruct.js';
import { nodePdfOpener, readFixtureBytes } from '../../fixtures/pdf/node-opener.js';
import { pdfFixture } from '../../fixtures/pdf/manifest.js';

async function paragraphsOf(name: string) {
  const data = readFixtureBytes(pdfFixture(name));
  const extraction = await extractPdf(data, nodePdfOpener);
  return reconstructDocument(extraction);
}

describe('A3 — single-column reconstruction', () => {
  it('produces one paragraph per clause, in order, with the title as a heading', async () => {
    const paras = await paragraphsOf('single-column.base.pdf');
    const texts = paras.map((p) => p.text);
    expect(texts[0]).toBe('SALE AND PURCHASE AGREEMENT');
    expect(paras[0]!.isHeading).toBe(true);
    expect(paras[0]!.lowConfidence).toBe(false);
    // Each clause is its own paragraph (load-bearing for paragraph-granular diff).
    expect(texts).toContain('1.2 The Seller shall deliver the Goods to the Buyer on the Completion Date.');
    expect(texts.some((t) => t.includes('S$1,250,000'))).toBe(true);
    expect(paras.length).toBeGreaterThanOrEqual(8);
    expect(paras.every((p) => p.column === 0)).toBe(true);
  });
});

describe('A3 — multi-column reconstruction (P5)', () => {
  it('orders columns left→right and leaks no token across the gutter', async () => {
    const paras = await paragraphsOf('two-column.pdf');
    const colA = paras.filter((p) => p.column === 0).map((p) => p.text).join(' ');
    const colB = paras.filter((p) => p.column === 1).map((p) => p.text).join(' ');

    // Two columns detected; everything flagged low-confidence.
    expect(paras.some((p) => p.column === 1)).toBe(true);
    expect(paras.every((p) => p.lowConfidence)).toBe(true);

    // No cross-column leak: column A text never contains column-B-only tokens.
    expect(colA).toContain('COLUMN A');
    expect(colA).not.toContain('COLUMN B');
    expect(colA).not.toContain('operative');
    expect(colB).toContain('COLUMN B');
    expect(colB).not.toContain('COLUMN A');
    expect(colB).not.toContain('recitals');

    // Column order: every column-0 paragraph precedes every column-1 paragraph.
    const firstB = paras.findIndex((p) => p.column === 1);
    const lastA = paras.map((p) => p.column).lastIndexOf(0);
    expect(lastA).toBeLessThan(firstB);
  });
});

describe('A3 — justified/kerned reconstruction (P11)', () => {
  it('collapses fragmented line items into one block per line', async () => {
    const paras = await paragraphsOf('justified-kerned.pdf');
    // The single justified line is ONE paragraph, not a swarm of fragments.
    expect(paras).toHaveLength(1);
    expect(paras[0]!.text).toContain('The parties hereby agree');
    expect(paras[0]!.text).toContain('performed promptly');
  });
});

describe('A3 — rotated-text filter', () => {
  it('skips items rotated beyond the threshold', () => {
    const page: PdfPageText = {
      pageNumber: 1,
      charCount: 20,
      outOfPrintableRatio: 0,
      items: [
        { str: 'Horizontal body text', hasEOL: true, transform: [11, 0, 0, 11, 72, 700], width: 110, height: 11 },
        // 45°-rotated watermark — a=b=11*cos/sin(45°)≈7.78.
        { str: 'DRAFT WATERMARK', hasEOL: true, transform: [7.78, 7.78, -7.78, 7.78, 300, 400], width: 90, height: 11 },
      ],
    };
    const paras = reconstructPage(page);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.text).toBe('Horizontal body text');
  });
});
