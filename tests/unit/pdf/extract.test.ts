/**
 * A2 — PDF extraction + corruption guards, against the committed fixtures and
 * synthetic inputs for the guard cores.
 *
 * Covers: text-PDF yields items (extraction works); image-only → P7
 * NoExtractableTextError; over-cap → typed PdfInputTooLargeError, no partial
 * diff; encoding garble ratio (P9) on synthetic text; OCR-layer assessment
 * (P7b) on a synthetic image-dominance signal.
 */
import { describe, it, expect } from 'vitest';

import {
  extractPdf,
  outOfPrintableRatio,
  flagOcrLayers,
  NoExtractableTextError,
  GARBLE_RATIO_THRESHOLD,
  type PdfPageText,
} from '../../../src/features/compare/io/pdf/extract.js';
import {
  checkByteCap,
  checkPageCap,
  PdfInputTooLargeError,
  MAX_PDF_PAGES,
} from '../../../src/features/compare/io/pdf/limits.js';
import type { PdfDocumentProxyLike, PdfOpener } from '../../../src/features/compare/io/pdf/pdfjs-types.js';
import { nodePdfOpener, readFixtureBytes } from '../../fixtures/pdf/node-opener.js';
import { pdfFixture } from '../../fixtures/pdf/manifest.js';

describe('A2 — PDF extraction (generated fixtures)', () => {
  it('extracts text items from a single-column text PDF', async () => {
    const data = readFixtureBytes(pdfFixture('single-column.base.pdf'));
    const result = await extractPdf(data, nodePdfOpener);
    expect(result.pages).toHaveLength(1);
    const page = result.pages[0]!;
    expect(page.items.length).toBeGreaterThan(0);
    const joined = page.items.map((i) => i.str).join(' ');
    expect(joined).toContain('SALE AND PURCHASE AGREEMENT');
    expect(joined).toContain('S$1,250,000');
    // Items carry geometry + hasEOL for reconstruction.
    const textItem = page.items.find((i) => i.str.length > 0)!;
    expect(textItem.transform).toHaveLength(6);
    expect(typeof textItem.hasEOL).toBe('boolean');
  });

  it('P7 — image-only PDF throws NoExtractableTextError naming the page', async () => {
    // pdfjs detaches the input buffer on use, so extract once and inspect the error.
    const data = readFixtureBytes(pdfFixture('image-only.pdf'));
    const err = await extractPdf(data, nodePdfOpener).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(NoExtractableTextError);
    expect((err as NoExtractableTextError).pageNumbers).toEqual([1]);
  });

  it('extracts the invisible OCR text layer (still readable, not dropped)', async () => {
    const data = readFixtureBytes(pdfFixture('ocr-layer.pdf'));
    const result = await extractPdf(data, nodePdfOpener);
    // Invisible (Tr 3) text is still extracted — the page is not silently empty.
    expect(result.pages.length).toBeGreaterThan(0);
    expect(result.pages[0]!.items.map((i) => i.str).join(' ')).toContain('OCR text layer');
  });
});

describe('A2 — hard input caps (P-cap)', () => {
  it('checkByteCap / checkPageCap throw typed PdfInputTooLargeError', () => {
    expect(() => checkByteCap(200 * 1024 * 1024)).toThrow(PdfInputTooLargeError);
    expect(() => checkPageCap(MAX_PDF_PAGES + 1)).toThrow(PdfInputTooLargeError);
  });

  it('extractPdf rejects an over-page-cap document with no partial result', async () => {
    const hugeDoc: PdfDocumentProxyLike = {
      numPages: MAX_PDF_PAGES + 1,
      getPage: () => Promise.reject(new Error('must not be called')),
      destroy: () => Promise.resolve(),
    };
    const stubOpener: PdfOpener = () => Promise.resolve(hugeDoc);
    await expect(extractPdf(new Uint8Array([1, 2, 3]), stubOpener)).rejects.toBeInstanceOf(
      PdfInputTooLargeError,
    );
  });
});

describe('A2 — encoding garble (P9) + OCR assessment (P7b)', () => {
  it('outOfPrintableRatio flags private-use/control glyphs, not soft hyphens', () => {
    expect(outOfPrintableRatio('The quick brown fox.')).toBe(0);
    // Soft hyphen (U+00AD), en/em dash, NBSP are legitimate — ratio stays 0.
    const legit = 'indem' + String.fromCodePoint(0x00ad) + 'nify ' + String.fromCodePoint(0x2014);
    expect(outOfPrintableRatio(legit)).toBe(0);
    // A run of private-use-area glyphs (missing ToUnicode) → high ratio.
    const garbled = String.fromCodePoint(0xe001, 0xe002, 0xe003, 0xe004) + ' x';
    expect(outOfPrintableRatio(garbled)).toBeGreaterThan(GARBLE_RATIO_THRESHOLD);
  });

  it('extractPdf excludes a wholly-garbled page → NoExtractableTextError', async () => {
    const garbledStr = String.fromCodePoint(0xe010, 0xe011, 0xe012, 0xe013, 0xe014);
    const garbledPage = {
      getTextContent: () =>
        Promise.resolve({
          items: [
            { str: garbledStr, hasEOL: true, transform: [11, 0, 0, 11, 72, 700], width: 60, height: 11 },
          ],
        }),
    };
    const doc: PdfDocumentProxyLike = {
      numPages: 1,
      getPage: () => Promise.resolve(garbledPage),
      destroy: () => Promise.resolve(),
    };
    await expect(extractPdf(new Uint8Array([1]), () => Promise.resolve(doc))).rejects.toBeInstanceOf(
      NoExtractableTextError,
    );
  });

  it('flagOcrLayers flags only text pages that are image-dominant (badge, not drop)', () => {
    const pages: PdfPageText[] = [
      { pageNumber: 1, items: [], charCount: 40, outOfPrintableRatio: 0 },
      { pageNumber: 2, items: [], charCount: 40, outOfPrintableRatio: 0 },
      { pageNumber: 3, items: [], charCount: 0, outOfPrintableRatio: 0 },
    ];
    const flagged = flagOcrLayers(pages, new Set([2, 3]));
    expect(flagged).toEqual([2]); // page 3 has no text; page 1 not image-dominant
  });
});
