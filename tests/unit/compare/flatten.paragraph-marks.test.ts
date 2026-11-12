/**
 * a revision on a paragraph MARK is a paragraph join, and the
 * flattener resolves it as Word does. Accepting a deleted mark joins the
 * paragraph to the next one (the next one's properties win); rejecting keeps
 * both and drops the mark. An inserted mark is the mirror. The last paragraph
 * of a body or cell, and one followed by a table, has nothing to join and stays.
 *
 * This is how ADR-0032 authors a merge and ADR-0035 a whole deleted paragraph,
 * so before it a merge accepted read as two paragraphs, and an input document
 * that already carried a tracked paragraph join was compared as two.
 */
import { describe, it, expect } from 'vitest';

import { flattenTrackedChanges } from '@features/compare/preprocess/flatten.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REV = 'w:id="9" w:author="Counterparty" w:date="2026-01-01T00:00:00Z"';
const doc = (body: string): string => `<w:document ${W}><w:body>${body}</w:body></w:document>`;
const r = (t: string): string => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
/** A paragraph whose mark carries `kind`, with an optional style. */
const marked = (kind: 'ins' | 'del' | 'moveFrom' | 'moveTo', inner: string, style = 'Body1'): string =>
  `<w:p w:rsidR="00AA"><w:pPr><w:pStyle w:val="${style}"/><w:rPr><w:${kind} ${REV}/></w:rPr></w:pPr>${inner}</w:p>`;
const plain = (inner: string, style = 'Body2'): string =>
  `<w:p w:rsidR="00BB"><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${inner}</w:p>`;
const paras = (xml: string): string[] => xml.match(/<w:p[\s>][\s\S]*?<\/w:p>/g) ?? [];

describe('a revised paragraph mark is a paragraph join', () => {
  it('accept-all joins a paragraph whose mark was deleted to the next; the next one\'s properties win', () => {
    const xml = doc(marked('del', r('The Purchaser shall notify the Vendor')) + plain(r(', and the notice shall specify the claim.')));
    const out = flattenTrackedChanges(xml, 'acceptAll');
    expect(paras(out)).toHaveLength(1);
    expect(paras(out)[0]).toContain('w:rsidR="00BB"');
    expect(paras(out)[0]).toContain('w:val="Body2"');
    expect(paras(out)[0]).not.toContain('Body1');
    expect(out.indexOf('notify the Vendor')).toBeLessThan(out.indexOf(', and the notice'));
    expect(out).not.toMatch(/<w:del[\s/>]/);
  });

  it('reject-all keeps both paragraphs and drops the deleted mark', () => {
    const xml = doc(marked('del', r('Head.')) + plain(r('Tail.')));
    const out = flattenTrackedChanges(xml, 'rejectAll');
    expect(paras(out)).toHaveLength(2);
    expect(out).not.toMatch(/<w:del[\s/>]/);
    expect(out).toContain('Body1');
  });

  it('an inserted mark is the mirror: reject-all joins, accept-all keeps both', () => {
    const xml = doc(marked('ins', r('Head.')) + plain(r('Tail.')));
    expect(paras(flattenTrackedChanges(xml, 'rejectAll'))).toHaveLength(1);
    expect(paras(flattenTrackedChanges(xml, 'acceptAll'))).toHaveLength(2);
  });

  it('a moved-away or moved-here mark joins like a deleted or inserted one', () => {
    expect(paras(flattenTrackedChanges(doc(marked('moveFrom', r('A.')) + plain(r('B.'))), 'acceptAll'))).toHaveLength(1);
    expect(paras(flattenTrackedChanges(doc(marked('moveTo', r('A.')) + plain(r('B.'))), 'rejectAll'))).toHaveLength(1);
  });

  it('a whole deleted paragraph (runs struck AND mark struck, ADR-0035) disappears on accept-all', () => {
    const struck = `<w:p><w:pPr><w:rPr><w:del ${REV}/></w:rPr></w:pPr><w:del ${REV}><w:r><w:delText>Gone.</w:delText></w:r></w:del></w:p>`;
    const out = flattenTrackedChanges(doc(plain(r('Before.')) + struck + plain(r('After.'))), 'acceptAll');
    expect(paras(out)).toHaveLength(2);
    expect(out).not.toContain('Gone.');
  });

  it('joins a chain: two deleted marks in a row make one paragraph of three', () => {
    const out = flattenTrackedChanges(doc(marked('del', r('One ')) + marked('del', r('two ')) + plain(r('three.'))), 'acceptAll');
    expect(paras(out)).toHaveLength(1);
    expect(paras(out)[0]).toMatch(/One [\s\S]*two [\s\S]*three\./);
  });

  it('a bookmark between the two paragraphs does not stop the join', () => {
    const xml = doc(marked('del', r('Head ')) + '<w:bookmarkStart w:id="1" w:name="x"/><w:bookmarkEnd w:id="1"/>' + plain(r('tail.')));
    expect(paras(flattenTrackedChanges(xml, 'acceptAll'))).toHaveLength(1);
  });

  it('the LAST paragraph of a body or cell, or one before a table, stays — there is nothing to join', () => {
    const last = flattenTrackedChanges(doc(plain(r('First.')) + marked('del', r('Last.'))), 'acceptAll');
    expect(paras(last)).toHaveLength(2);
    expect(last).not.toMatch(/<w:del[\s/>]/);
    const cell = (inner: string): string => `<w:tbl><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;
    const inCell = flattenTrackedChanges(doc(cell(plain(r('A.')) + marked('del', r('B.'))) + plain(r('After.'))), 'acceptAll');
    expect(paras(inCell)).toHaveLength(3);
    const beforeTable = flattenTrackedChanges(doc(marked('del', r('Head.')) + cell(plain(r('Cell.'))) + plain(r('After.'))), 'acceptAll');
    expect(paras(beforeTable)).toHaveLength(3);
  });

  it('joins inside a table cell, never across the cell boundary', () => {
    const cell = `<w:tbl><w:tr><w:tc>${marked('del', r('Cell head '))}${plain(r('cell tail.'))}</w:tc></w:tr></w:tbl>`;
    const out = flattenTrackedChanges(doc(cell + plain(r('After.'))), 'acceptAll');
    expect(paras(out)).toHaveLength(2);
    expect(paras(out)[0]).toMatch(/Cell head [\s\S]*cell tail\./);
  });

  it('a revision inside w:pPrChange is the OLD properties, not the mark — no join', () => {
    const p = `<w:p><w:pPr><w:pStyle w:val="Body1"/><w:pPrChange ${REV}><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange></w:pPr>${r('Head.')}</w:p>`;
    expect(paras(flattenTrackedChanges(doc(p + plain(r('Tail.'))), 'acceptAll'))).toHaveLength(2);
  });
});
