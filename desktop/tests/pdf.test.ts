import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { NoExtractableTextError } from '@features/compare/io/pdf/extract.js';
import { explainPdfError, normaliseLigatures, readPdf } from '../src/renderer/src/pdf/read-pdf.js';
import { convertUploads } from '../src/renderer/src/paste/convert-uploads.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import { blocksText } from '../src/renderer/src/paste/model.js';
import type { LoadedDoc } from '../src/renderer/src/state/documents.js';
import { nodePdfOpener } from './helpers/node-pdf-opener.js';

const fixture = (name: string) => new Uint8Array(readFileSync(resolve(import.meta.dirname, 'fixtures/pdf', name)));
const read = (bytes: Uint8Array) => readPdf(bytes.slice(), nodePdfOpener);

const asDoc = (name: string, id: string): LoadedDoc => {
  const bytes = fixture(name);
  return {
    id, name, byteLength: bytes.length, kind: 'pdf', overCap: false,
    file: new File([bytes as BlobPart], name, { type: 'application/pdf' }),
  };
};

describe('reading a PDF', () => {
  it('rebuilds the paragraphs of a multi-page contract', async () => {
    const { blocks } = await read(fixture('contract-a.pdf'));
    const text = blocksText(blocks);
    expect(text).toContain('SHARE PURCHASE AGREEMENT');
    expect(text).toContain('The Purchaser shall pay the Seller S$5,000 on Completion');
    expect(text).toContain('within 30 days after the Completion Date');
  });

  it('leaves out bare page numbers at the page edge, keeps repeated header text, and says so', async () => {
    const { blocks, notes } = await read(fixture('contract-numbered.pdf'));
    const paragraphs = blocksText(blocks).split('\n');
    expect(paragraphs.filter((p) => /^\d$/.test(p.trim()))).toEqual([]);
    // "CONFIDENTIAL" repeats on every page. It is text, not a page number: kept.
    expect(paragraphs.filter((p) => p.trim() === 'CONFIDENTIAL')).toHaveLength(4);
    expect(notes.some((n) => /^4 page numbers/.test(n) && /header and footer text is kept/.test(n))).toBe(true);
  });

  it('keeps a footer that carries more than a number (a path, a date) as text', async () => {
    const { blocks, notes } = await read(fixture('contract-a.pdf'));
    expect(blocksText(blocks)).toContain('Share Purchase Agreement');
    expect(notes.some((n) => /page number/.test(n))).toBe(false);
  });

  it('keeps paragraphs whole at 1.5 line spacing instead of one paragraph per line', async () => {
    const { blocks } = await read(fixture('contract-spaced.pdf'));
    // 1 title + 18 clauses x (heading + body) = 37; a split-per-line read gives ~67.
    expect(blocks.length).toBeGreaterThanOrEqual(35);
    expect(blocks.length).toBeLessThanOrEqual(38);
    expect(blocksText(blocks).split('\n').every((l) => /^[A-Z0-9]/.test(l))).toBe(true);
  });

  it('reads a PDF whose font is a plain monospaced text face', async () => {
    const { blocks } = await read(fixture('base14.pdf'));
    expect(blocksText(blocks)).toContain('The Purchaser shall pay S$5,000 within 30 days.');
  });

  it('always carries the reconstruction notice', async () => {
    const { notes } = await read(fixture('contract-a.pdf'));
    expect(notes[0]).toMatch(/Read from a PDF/);
    expect(notes[0]).toMatch(/rebuilt from the page layout/);
  });

  it('refuses a scanned, image-only PDF instead of comparing nothing', async () => {
    await expect(read(fixture('scanned.pdf'))).rejects.toBeInstanceOf(NoExtractableTextError);
    const reason = await read(fixture('scanned.pdf')).catch(explainPdfError);
    expect(reason).toMatch(/scanned or image-only/);
    expect(reason).toMatch(/cannot read scanned PDFs/);
  });

  it('says plainly when the file is not a PDF', async () => {
    const reason = await readPdf(new TextEncoder().encode('not a pdf at all'), nodePdfOpener).catch(explainPdfError);
    expect(reason).toMatch(/not a valid PDF|could not be read as a PDF/);
  });

  it('undoes typesetter ligatures so they do not show as changes against Word text', () => {
    expect(normaliseLigatures('ﬁnancial ﬂow oﬃce soft­hyphen')).toBe('financial flow office softhyphen');
  });
});

describe('uploaded PDFs through the normal compare path', () => {
  const upload = (docs: LoadedDoc[]) =>
    convertUploads(docs, (bytes) => readPdf(bytes, nodePdfOpener));

  it('become a docx that keeps their name and carries the notes', async () => {
    const { docs, failed } = await upload([asDoc('contract-a.pdf', 'd1')]);
    expect(failed).toEqual([]);
    expect(docs[0]).toMatchObject({ id: 'd1', name: 'contract-a.pdf', kind: 'docx' });
    expect(docs[0]?.notes?.[0]).toMatch(/Read from a PDF/);
  });

  it('are refused at upload, with the reason, when scanned', async () => {
    const { docs, failed } = await upload([asDoc('scanned.pdf', 'd2')]);
    expect(docs).toEqual([]);
    expect(failed[0]?.name).toBe('scanned.pdf');
    expect(failed[0]?.reason).toMatch(/scanned/);
  });

  it('compare PDF against PDF and find the edits', async () => {
    const { docs } = await upload([asDoc('contract-a.pdf', 'a'), asDoc('contract-b.pdf', 'b')]);
    const out = await runComparison({
      originalBytes: await docs[0]!.file.arrayBuffer(),
      revisedBytes: await docs[1]!.file.arrayBuffer(),
    });
    expect(out.summary.total).toBeGreaterThan(0);
    const shown = JSON.stringify(out.model);
    expect(shown).toContain('S$6,000');
    expect(shown).toContain('S$5,000');
    expect(shown).toContain('Governing law');
  });
});
