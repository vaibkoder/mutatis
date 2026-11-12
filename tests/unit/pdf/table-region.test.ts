/**
 * B4 (P12) — PDF approximate table-region cell diff.
 *
 * A fee table whose only change is one cell (S$500,000 → S$650,000) must diff as
 * a single LOW-CONFIDENCE cell change at that cell, NOT whole-row prose churn —
 * and a region that is not confidently a grid must return null so the caller
 * falls back to flagged prose (never suppressed, ADR-0008).
 */
import { describe, it, expect } from 'vitest';

import type { PdfTextItem } from '../../../src/features/compare/io/pdf/pdfjs-types.js';
import {
  detectTableRegion,
  diffPdfTableRegion,
} from '../../../src/features/compare/pdf/table-region.js';
import { extractPdf } from '../../../src/features/compare/io/pdf/extract.js';
import { nodePdfOpener, readFixtureBytes } from '../../fixtures/pdf/node-opener.js';
import { pdfFixture } from '../../fixtures/pdf/manifest.js';

/** Synthesise a positioned text item (transform = [size,0,0,size,x,y]). */
function item(str: string, x: number, y: number, size = 11): PdfTextItem {
  return {
    str,
    hasEOL: false,
    transform: [size, 0, 0, size, x, y],
    width: str.length * size * 0.5,
    height: size,
  };
}

describe('B4 — detectTableRegion (pure geometry)', () => {
  it('grids aligned x-bands into rows × cols', () => {
    const items = [
      item('Milestone', 72, 700), item('Due Date', 272, 700), item('Fee', 432, 700),
      item('Signing', 72, 672), item('01 Jan 2026', 272, 672), item('S$250,000', 432, 672),
      item('Completion', 72, 644), item('01 Mar 2026', 272, 644), item('S$500,000', 432, 644),
    ];
    const grid = detectTableRegion(items);
    expect(grid).not.toBeNull();
    expect(grid!).toHaveLength(3);
    expect(grid!.every((r) => r.cells.length === 3)).toBe(true);
    expect(grid![0]!.cells.map((c) => c.text)).toEqual(['Milestone', 'Due Date', 'Fee']);
  });

  it('returns null for single-column prose (not a table)', () => {
    const items = [
      item('The Seller shall deliver the Goods within thirty days.', 72, 700),
      item('Completion is conditional on the conditions precedent.', 72, 680),
      item('This Agreement is governed by the laws of Singapore.', 72, 660),
    ];
    expect(detectTableRegion(items)).toBeNull();
  });

  it('returns null for a single row (too few rows)', () => {
    const items = [item('A', 72, 700), item('B', 272, 700), item('C', 432, 700)];
    expect(detectTableRegion(items)).toBeNull();
  });
});

describe('B4 — diffPdfTableRegion (synthetic, one changed cell)', () => {
  it('diffs only the changed cell as low-confidence, not the whole row', () => {
    const base = detectTableRegion([
      item('Milestone', 72, 700), item('Due Date', 272, 700), item('Fee', 432, 700),
      item('Completion', 72, 672), item('01 Mar 2026', 272, 672), item('S$500,000', 432, 672),
    ])!;
    const revised = detectTableRegion([
      item('Milestone', 72, 700), item('Due Date', 272, 700), item('Fee', 432, 700),
      item('Completion', 72, 672), item('01 Mar 2026', 272, 672), item('S$650,000', 432, 672),
    ])!;

    const ops = diffPdfTableRegion(base, revised, 'Reviewer', '2026-01-01');
    const changed = ops.filter((o) => o.type !== 'equal');
    expect(changed).toHaveLength(1);
    expect(changed[0]!.type).toBe('replace');
    expect(changed[0]!.partId).toBe('tableCell:r1c2'); // row 1 (Completion), col 2 (Fee)
    expect(changed[0]!.lowConfidence).toBe(true);
    expect(changed[0]!.id.startsWith('pdf-table-op-')).toBe(true);
    // Never suppressed: nothing deleted/inserted wholesale; row stays intact.
    expect(ops.some((o) => o.type === 'delete' || o.type === 'insert')).toBe(false);
  });
});

describe('B4 (P12) — fee-table fixtures end-to-end', () => {
  it('one changed cell → one low-confidence cell change, not row churn', async () => {
    const base = await extractPdf(readFixtureBytes(pdfFixture('fee-table.base.pdf')), nodePdfOpener);
    const revised = await extractPdf(
      readFixtureBytes(pdfFixture('fee-table.revised.pdf')),
      nodePdfOpener,
    );

    const baseGrid = detectTableRegion(base.pages[0]!.items);
    const revGrid = detectTableRegion(revised.pages[0]!.items);
    expect(baseGrid).not.toBeNull();
    expect(revGrid).not.toBeNull();
    expect(baseGrid!).toHaveLength(4); // header + 3 data rows
    expect(baseGrid!.every((r) => r.cells.length === 3)).toBe(true);

    const ops = diffPdfTableRegion(baseGrid!, revGrid!, 'Reviewer', '2026-01-01');
    const changed = ops.filter((o) => o.type !== 'equal');
    expect(changed).toHaveLength(1);
    expect(changed[0]!.type).toBe('replace');
    expect(changed[0]!.partId).toBe('tableCell:r2c2'); // Completion row, Fee column
    expect(changed[0]!.lowConfidence).toBe(true);
    expect(ops.some((o) => o.type === 'delete' || o.type === 'insert')).toBe(false);
  });
});
