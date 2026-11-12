/**
 * PDF → ParaRecord adapter.
 *
 * Maps reconstructed PDF paragraphs (A3) onto the SAME `ParaRecord` shape the
 * Word path produces, so the entire diff/classify pipeline is reused verbatim
 * (ADR-0017 reuse). The Word-specific fields are inert for PDF: `rawXml=''`,
 * `styleId=null`, `isTableCell=false` (PDF tables go through B4).
 *
 * The load-bearing piece is {@link canonicalisePdfText}: PDF extraction yields
 * typographic ligatures (ﬀﬁﬂﬃﬄﬅﬆ) and soft hyphens that Word stores as plain
 * letters. Canonicalising them BEFORE the shared normalise/hash makes identical
 * wording hash identically across a PDF and a Word document — the anchor-match
 * parity Word-vs-PDF depends on (P10).
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { ParaRecord } from '../../model.js';
import { fnv1a } from '../../preprocess/hash.js';
import {
  normaliseText,
  normaliseTextAggressive,
  stripNumberingPrefix,
} from '../../preprocess/normalise.js';
import { isSignatureBlockText } from '../../preprocess/pararecord.js';

import type { ReconstructedParagraph } from './reconstruct.js';

/** Ligature → ASCII expansion (U+FB00–FB06 + the long-s/round-s forms). */
const LIGATURES: ReadonlyArray<[string, string]> = [
  ['ﬀ', 'ff'],
  ['ﬁ', 'fi'],
  ['ﬂ', 'fl'],
  ['ﬃ', 'ffi'],
  ['ﬄ', 'ffl'],
  ['ﬅ', 'st'], // long s + t
  ['ﬆ', 'st'], // s + t
];

const SOFT_HYPHEN = /­/g; // U+00AD — invisible hyphenation point, drop it.

/**
 * Canonicalise PDF-extracted text so identical wording matches Word text:
 *  - expand typographic ligatures to their ASCII letters;
 *  - drop soft hyphens (and the discretionary-hyphen artefact at line breaks).
 *
 * Dash/quote/NBSP canonicalisation is intentionally NOT done here — it already
 * lives in the shared {@link normaliseText}, which both paths run next.
 */
export function canonicalisePdfText(text: string): string {
  let out = text;
  for (const [lig, ascii] of LIGATURES) {
    if (out.includes(lig)) out = out.split(lig).join(ascii);
  }
  out = out.replace(SOFT_HYPHEN, '');
  return out;
}

/**
 * Build a `ParaRecord` from one reconstructed PDF paragraph. The normalise →
 * strip-numbering → hash chain is byte-identical to the Word extractor's, so
 * matching text produces a matching hash.
 */
export function pdfParagraphToRecord(
  para: ReconstructedParagraph & { isHeaderFooter?: boolean },
  index: number,
  partId: string,
): ParaRecord {
  const canonical = canonicalisePdfText(para.text);
  const normWhole = normaliseText(canonical);
  const { prefix, body } = stripNumberingPrefix(normWhole);
  const normText = body;
  return {
    index,
    partId,
    rawXml: '', // PDF has no OOXML — the redline author is never reached for PDF.
    normText,
    normTextAggressive: normaliseTextAggressive(body),
    numberingPrefix: prefix,
    hash: fnv1a(normText),
    styleId: null,
    isTableCell: false,
    isSignatureBlock: isSignatureBlockText(normWhole),
    isHeaderFooter: para.isHeaderFooter ?? false,
    isFootnote: false,
  };
}

/**
 * Map a document's reconstructed paragraphs to `ParaRecord[]`. Header/footer
 * paragraphs (flagged upstream by repeated-line detection) carry the flag so the
 * pipeline's existing exclude-headers/footers rules apply unchanged.
 */
export function pdfParagraphsToRecords(
  paras: ReadonlyArray<ReconstructedParagraph & { isHeaderFooter?: boolean }>,
  config: CompareConfig,
  partId = 'document',
): ParaRecord[] {
  void config; // reserved — kept in signature to mirror the Word extractor.
  return paras.map((para, i) => pdfParagraphToRecord(para, i, partId));
}
