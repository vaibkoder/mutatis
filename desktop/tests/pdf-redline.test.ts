import { describe, it, expect } from 'vitest';
import { applyEdits } from '../src/renderer/src/edits/apply-edits.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import { PDFDocument } from '@pdfme/pdf-lib';
import { renderPdfRedline } from '../src/renderer/src/edits/pdf-redline.js';
import { makeDocxBytes } from './helpers/make-docx.js';

/**
 * The PDF redline of the assistant's own changes.
 *
 * Worth a real test rather than a snapshot of a call: this path pulls in the
 * vendored fonts, HarfBuzz and a PDF writer, and it is the ONLY place in the
 * app that does. A broken dynamic import here fails silently in the UI (a
 * button that spins and then reports a module error), so the test asserts on
 * the bytes.
 */

const DOC = [
  'This Agreement is made on 1 January 2026.',
  'The Purchase Price is GBP 1,000,000, payable on Completion.',
  'The Seller shall indemnify the Buyer against all losses.',
];

async function pdfOfAnEdit() {
  const target = makeDocxBytes(DOC);
  const { bytes } = applyEdits(target, [
    { find: 'GBP 1,000,000', replace: 'GBP 1,250,000', reason: 'agreed uplift' },
  ]);
  const outcome = await runComparison({
    originalBytes: target,
    revisedBytes: bytes,
    originalName: 'target.docx',
    revisedName: "target.docx (assistant's edits)",
    author: 'Mutatis assistant',
    date: '2026-01-01T00:00:00.000Z',
  });
  return renderPdfRedline(outcome.model, {
    title: "Assistant's changes to target.docx",
    author: 'Mutatis assistant',
  });
}

describe('renderPdfRedline', () => {
  it('produces a real PDF', async () => {
    const pdf = await pdfOfAnEdit();
    expect(pdf.byteLength).toBeGreaterThan(1000);
    // %PDF-
    expect([...pdf.slice(0, 5)]).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('ends with a complete trailer, not a truncated stream', async () => {
    const pdf = await pdfOfAnEdit();
    const tail = new TextDecoder('latin1').decode(pdf.slice(-2048));
    expect(tail).toContain('%%EOF');
    expect(tail).toContain('startxref');
  });

  it('embeds fonts rather than relying on the reader having them', async () => {
    const pdf = await pdfOfAnEdit();
    // The engine writes with object streams on, so the font descriptors are
    // inside a compressed container and a raw scan finds nothing. Re-serialising
    // without them is the honest way to look.
    const reopened = await PDFDocument.load(pdf);
    const plain = await reopened.save({ useObjectStreams: false });
    expect(new TextDecoder('latin1').decode(plain)).toMatch(/FontFile[23]?/);
  });

  it('carries the assistant as the document author, and a real page', async () => {
    const pdf = await pdfOfAnEdit();
    const reopened = await PDFDocument.load(pdf);
    expect(reopened.getAuthor()).toBe('Mutatis assistant');
    expect(reopened.getTitle()).toContain('target.docx');
    expect(reopened.getPageCount()).toBeGreaterThan(0);
  });

  it('paginates to A4-ish pages, not one endless sheet', async () => {
    const pdf = await pdfOfAnEdit();
    const reopened = await PDFDocument.load(pdf);
    const { width, height } = reopened.getPage(0).getSize();
    expect(height).toBeGreaterThan(width);
    expect(height).toBeLessThan(1200);
  });
}, 60_000);
