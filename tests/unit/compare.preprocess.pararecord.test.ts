import { describe, it, expect } from 'vitest';

import { getCompareConfig, UK_PROFILE } from '@shared/config/index.js';
import {
  extractParaRecords,
  preprocessSide,
  fnv1a,
} from '@features/compare/preprocess/index.js';

const CONFIG = getCompareConfig(UK_PROFILE);

const P = (inner: string): string => `<w:p>${inner}</w:p>`;
const run = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const body = (...paras: string[]): string =>
  `<w:document><w:body>${paras.join('')}</w:body></w:document>`;

describe('extractParaRecords (§4.1, §5.0)', () => {
  it('extracts run text, strips the numbering prefix, and hashes the prefix-free body', () => {
    const xml = body(P(`${run('4.2(a) ')}${run('The Seller shall pay.')}`));
    const recs = extractParaRecords(xml, 'document', CONFIG);
    expect(recs).toHaveLength(1);
    const r = recs[0]!;
    expect(r.numberingPrefix).toBe('4.2(a) ');
    expect(r.normText).toBe('The Seller shall pay.');
    expect(r.hash).toBe(fnv1a('The Seller shall pay.'));
  });

  it('reads the w:pStyle value into styleId; null when absent', () => {
    const styled = body(
      P(`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('Title')}`),
      P(run('Plain')),
    );
    const recs = extractParaRecords(styled, 'document', CONFIG);
    expect(recs[0]!.styleId).toBe('Heading1');
    expect(recs[1]!.styleId).toBeNull();
  });

  it('flags a paragraph inside a <w:tc> as a table cell', () => {
    const xml = body(
      `<w:tbl><w:tr><w:tc>${P(run('cell text'))}</w:tc></w:tr></w:tbl>`,
      P(run('outside')),
    );
    const recs = extractParaRecords(xml, 'document', CONFIG);
    const cell = recs.find((r) => r.normText === 'cell text')!;
    const outside = recs.find((r) => r.normText === 'outside')!;
    expect(cell.isTableCell).toBe(true);
    expect(outside.isTableCell).toBe(false);
  });

  it('flags signature-block paragraphs via §7.2 heuristics', () => {
    const xml = body(P(run('EXECUTED as a DEED by the Seller')), P(run('ordinary clause')));
    const recs = extractParaRecords(xml, 'document', CONFIG);
    expect(recs[0]!.isSignatureBlock).toBe(true);
    expect(recs[1]!.isSignatureBlock).toBe(false);
  });

  it('sets isHeaderFooter / isFootnote from partId', () => {
    const hdr = extractParaRecords(body(P(run('x'))), 'header1', CONFIG)[0]!;
    const ftn = extractParaRecords(body(P(run('x'))), 'footnotes', CONFIG)[0]!;
    const doc = extractParaRecords(body(P(run('x'))), 'document', CONFIG)[0]!;
    expect(hdr.isHeaderFooter).toBe(true);
    expect(ftn.isFootnote).toBe(true);
    expect(doc.isHeaderFooter).toBe(false);
    expect(doc.isFootnote).toBe(false);
  });
});

describe('preprocessSide pipeline — T7 (pre-existing tracked changes)', () => {
  it('flattens accept-all then extracts prefix-free, hash-equal records', () => {
    // Original: plain "The Seller shall pay."
    const original = body(P(`${run('4.2 ')}${run('The Seller shall pay.')}`));
    // Revised: same text but reached via an <w:ins> tracked change (added "shall pay.")
    const revised = body(
      P(
        `${run('4.2 ')}${run('The Seller ')}` +
          `<w:ins w:id="1" w:author="X"><w:r><w:t>shall pay.</w:t></w:r></w:ins>`,
      ),
    );

    const o = preprocessSide(original, 'document', 'acceptAll', CONFIG);
    const r = preprocessSide(revised, 'document', 'acceptAll', CONFIG);

    // ins content kept; del content (none here) removed; prefix stored; bodies hash-equal.
    expect(o.paras[0]!.numberingPrefix).toBe('4.2 ');
    expect(r.paras[0]!.numberingPrefix).toBe('4.2 ');
    expect(o.paras[0]!.normText).toBe('The Seller shall pay.');
    expect(r.paras[0]!.normText).toBe('The Seller shall pay.');
    expect(r.paras[0]!.hash).toBe(o.paras[0]!.hash);
  });

  it('drops <w:ins> content under reject-all', () => {
    const revised = body(
      P(`${run('Base ')}<w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins>`),
    );
    const r = preprocessSide(revised, 'document', 'rejectAll', CONFIG);
    expect(r.paras[0]!.normText).toBe('Base');
  });
});
