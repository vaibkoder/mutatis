/**
 * D3 (P13) — per-page image-presence detection (render-free blind-spot notice).
 *
 * A one-off signature image must be listed; a repeated letterhead logo must be
 * deduped so it does NOT fire on every page. This is a disclosure, never a diff.
 */
import { describe, it, expect } from 'vitest';

import {
  detectImagePresence,
  summarise,
  IMAGE_PRESENCE_CONFIG,
} from '../../../src/features/compare/io/pdf/images.js';
import { nodePdfOpener, loadPdfImageOps, readFixtureBytes } from '../../fixtures/pdf/node-opener.js';
import { pdfFixture } from '../../fixtures/pdf/manifest.js';

describe('D3 — summarise (pure dedupe/dominance)', () => {
  it('dedupes a rectangle repeated across pages as boilerplate', () => {
    const rects = [
      // Same logo rect on three pages → boilerplate, not content.
      { page: 1, x: 72, y: 721, w: 80, h: 40, coverage: 0.006 },
      { page: 2, x: 72, y: 721, w: 80, h: 40, coverage: 0.006 },
      { page: 3, x: 72, y: 721, w: 80, h: 40, coverage: 0.006 },
      // A unique signature rect on page 3 → content.
      { page: 3, x: 72, y: 581, w: 160, h: 60, coverage: 0.02 },
    ];
    const report = summarise(rects);
    expect(report.pagesWithContentImages).toEqual([3]);
    expect(report.contentImageCount).toBe(1);
    expect(report.boilerplateImageCount).toBe(1);
  });

  it('flags a page-dominant raster (OCR-layer signal)', () => {
    const rects = [{ page: 1, x: 0, y: 0, w: 595, h: 842, coverage: 0.99 }];
    const report = summarise(rects);
    expect(report.imageDominantPages).toEqual([1]);
    expect(report.contentImageCount).toBe(1);
  });

  it('honours the boilerplate threshold from config', () => {
    expect(IMAGE_PRESENCE_CONFIG.BOILERPLATE_MIN_PAGES).toBe(2);
  });
});

describe('D3 (P13) — fixtures', () => {
  it('lists the signature-image page', async () => {
    const ops = await loadPdfImageOps();
    const report = await detectImagePresence(
      readFixtureBytes(pdfFixture('signature-image.pdf')),
      nodePdfOpener,
      ops,
    );
    expect(report.pagesWithContentImages).toEqual([1]);
    expect(report.contentImageCount).toBe(1);
  });

  it('does NOT flag the repeated letterhead logo on every page', async () => {
    const ops = await loadPdfImageOps();
    const report = await detectImagePresence(
      readFixtureBytes(pdfFixture('repeated-logo.pdf')),
      nodePdfOpener,
      ops,
    );
    // The logo repeats on all 3 pages at the same rect → boilerplate, deduped out.
    expect(report.pagesWithContentImages).toEqual([]);
    expect(report.contentImageCount).toBe(0);
    expect(report.boilerplateImageCount).toBeGreaterThanOrEqual(1);
  });

  it('reports no image content for a text-only table page', async () => {
    const ops = await loadPdfImageOps();
    const report = await detectImagePresence(
      readFixtureBytes(pdfFixture('fee-table.base.pdf')),
      nodePdfOpener,
      ops,
    );
    expect(report.pagesWithContentImages).toEqual([]);
    expect(report.contentImageCount).toBe(0);
  });
});
