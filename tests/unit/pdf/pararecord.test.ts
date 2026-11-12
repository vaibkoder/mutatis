/**
 * A4 — PDF → ParaRecord adapter, incl. the ligature/soft-hyphen hash parity that
 * Word-vs-PDF anchoring depends on (P10).
 *
 * Covers: PDF records match the model; a PDF paragraph with ligatures + soft
 * hyphens hashes IDENTICALLY to the same words extracted from a Word document
 * (P10); signature blocks flagged; header/footer repeated-line detection.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';

import { extractParaRecords } from '../../../src/features/compare/preprocess/pararecord.js';
import {
  canonicalisePdfText,
  pdfParagraphToRecord,
  pdfParagraphsToRecords,
} from '../../../src/features/compare/io/pdf/pararecord.js';
import { flagHeadersFooters } from '../../../src/features/compare/io/pdf/parts.js';
import type { ReconstructedParagraph } from '../../../src/features/compare/io/pdf/reconstruct.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

function para(text: string, over: Partial<ReconstructedParagraph> = {}): ReconstructedParagraph {
  return {
    text,
    pageNumber: 1,
    column: 0,
    yTop: 700,
    fontSize: 11,
    isHeading: false,
    lowConfidence: false,
    ...over,
  };
}

/** Wrap plain text as a minimal Word paragraph for the real Word extractor. */
function wordBody(...lines: string[]): string {
  const ps = lines
    .map((l) => `<w:p><w:r><w:t xml:space="preserve">${l}</w:t></w:r></w:p>`)
    .join('');
  return `<w:body>${ps}</w:body>`;
}

describe('A4 — canonicalisePdfText', () => {
  it('expands ligatures and drops soft hyphens', () => {
    const ligated = 'o' + String.fromCodePoint(0xfb03) + 'ce'; // o + ﬃ + ce = "office"
    const soft = 'indem' + String.fromCodePoint(0x00ad) + 'nity'; // soft hyphen
    expect(canonicalisePdfText(ligated)).toBe('office');
    expect(canonicalisePdfText(soft)).toBe('indemnity');
    expect(canonicalisePdfText('e' + String.fromCodePoint(0xfb00) + 'ort')).toBe('effort');
  });
});

describe('A4 — PDF → ParaRecord', () => {
  it('produces a record matching the model with inert Word fields', () => {
    const rec = pdfParagraphToRecord(para('2.1 The Buyer shall pay S$1,500,000.'), 0, 'document');
    expect(rec.rawXml).toBe('');
    expect(rec.styleId).toBeNull();
    expect(rec.isTableCell).toBe(false);
    expect(rec.numberingPrefix).toBe('2.1 ');
    expect(rec.normText).toBe('The Buyer shall pay S$1,500,000.');
    expect(rec.hash).toMatch(/^[0-9a-f]{8}$/);
  });

  it('P10 — ligature + soft-hyphen PDF text hashes identically to the Word equivalent', () => {
    // PDF side: real ligature glyphs + a soft hyphen at a former line break.
    const pdfText =
      'The o' +
      String.fromCodePoint(0xfb03) +
      'ce shall con' +
      String.fromCodePoint(0xfb01) +
      'rm the e' +
      String.fromCodePoint(0xfb00) +
      'ective indem' +
      String.fromCodePoint(0x00ad) +
      'nity.';
    const pdfRec = pdfParagraphToRecord(para(pdfText), 0, 'document');

    // Word side: the same words, plain (how Word stores them).
    const wordPlain = 'The office shall confirm the effective indemnity.';
    const wordRecs = extractParaRecords(wordBody(wordPlain), 'document', CONFIG);

    expect(pdfRec.normText).toBe(wordRecs[0]!.normText);
    expect(pdfRec.hash).toBe(wordRecs[0]!.hash); // anchor-match parity
  });

  it('flags signature/execution blocks', () => {
    const recs = pdfParagraphsToRecords(
      [para('Signed by the Seller in the presence of a witness.'), para('Ordinary clause text.')],
      CONFIG,
    );
    expect(recs[0]!.isSignatureBlock).toBe(true);
    expect(recs[1]!.isSignatureBlock).toBe(false);
  });
});

describe('A4 — header/footer detection', () => {
  it('flags a short line repeated across pages and bare page numbers', () => {
    const flagged = flagHeadersFooters([
      para('Project Falcon — Confidential', { pageNumber: 1 }),
      para('1.1 The body of page one differs entirely.', { pageNumber: 1 }),
      para('Project Falcon — Confidential', { pageNumber: 2 }),
      para('2.1 The body of page two differs entirely.', { pageNumber: 2 }),
      para('2', { pageNumber: 2 }),
    ]);
    const running = flagged.filter((p) => p.text.includes('Project Falcon'));
    expect(running.every((p) => p.isHeaderFooter)).toBe(true);
    expect(flagged.find((p) => p.text === '2')!.isHeaderFooter).toBe(true);
    expect(flagged.find((p) => p.text.startsWith('1.1'))!.isHeaderFooter).toBe(false);
  });
});
