/**
 * ADR-0046 — flattening a document's EXISTING tracked changes
 * must resolve a row or cell that is itself inserted or deleted. Those are
 * marked in the row's or cell's properties, not by wrapping, so before this the
 * flattener dropped the mark and the text and left an empty row or cell behind.
 */
import { describe, it, expect } from 'vitest';

import { flattenTrackedChanges } from '@features/compare/preprocess/flatten.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REV = 'w:id="9" w:author="Counterparty" w:date="2026-01-01T00:00:00Z"';
const cell = (t: string, tcPr = ''): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const insCell = (t: string, tcPr = ''): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:ins ${REV}><w:r><w:t>${t}</w:t></w:r></w:ins></w:p></w:tc>`;
const delCell = (t: string, tcPr = ''): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:del ${REV}><w:r><w:delText>${t}</w:delText></w:r></w:del></w:p></w:tc>`;
const doc = (rows: string): string => `<w:document ${W}><w:body><w:tbl>${rows}</w:tbl></w:body></w:document>`;
const rows = (xml: string): number => (xml.match(/<w:tr[\s>]/g) ?? []).length;
const cells = (xml: string): number => (xml.match(/<w:tc[\s>]/g) ?? []).length;

describe('flattening tracked row and cell revisions (ADR-0046)', () => {
  const insertedRow = doc(
    `<w:tr>${cell('Fee')}</w:tr>` + `<w:tr><w:trPr><w:ins ${REV}/></w:trPr>${insCell('Deposit')}</w:tr>`,
  );
  const deletedRow = doc(
    `<w:tr>${cell('Fee')}</w:tr>` + `<w:tr><w:trPr><w:del ${REV}/></w:trPr>${delCell('Deposit')}</w:tr>`,
  );

  it('accept-all keeps an inserted row and drops a deleted one', () => {
    const kept = flattenTrackedChanges(insertedRow, 'acceptAll');
    expect(rows(kept)).toBe(2);
    expect(kept).toContain('Deposit');
    expect(kept).not.toMatch(/<w:ins[\s/>]/);
    const gone = flattenTrackedChanges(deletedRow, 'acceptAll');
    expect(rows(gone)).toBe(1);
    expect(gone).not.toContain('Deposit');
  });

  it('reject-all drops an inserted row and keeps a deleted one as ordinary text', () => {
    const gone = flattenTrackedChanges(insertedRow, 'rejectAll');
    expect(rows(gone)).toBe(1);
    expect(gone).not.toContain('Deposit');
    const kept = flattenTrackedChanges(deletedRow, 'rejectAll');
    expect(rows(kept)).toBe(2);
    expect(kept).toContain('<w:t>Deposit</w:t>');
    expect(kept).not.toMatch(/<w:del[\s/>]/);
  });

  it('resolves an inserted and a deleted cell both ways, and strips the cell marks', () => {
    const table = doc(
      `<w:tr>${cell('Fee')}${insCell('New', `<w:cellIns ${REV}/>`)}${delCell('Old', `<w:cellDel ${REV}/>`)}</w:tr>`,
    );
    const accepted = flattenTrackedChanges(table, 'acceptAll');
    expect(cells(accepted)).toBe(2);
    expect(accepted).toContain('New');
    expect(accepted).not.toContain('Old');
    const rejected = flattenTrackedChanges(table, 'rejectAll');
    expect(cells(rejected)).toBe(2);
    expect(rejected).toContain('Old');
    expect(rejected).not.toContain('New');
    for (const xml of [accepted, rejected]) expect(xml).not.toMatch(/<w:cell(Ins|Del)[\s/>]/);
  });

  it('a table every row of which the policy removes goes whole, as Word removes it — no empty shell', () => {
    // A table inserted with tracked changes marks every row; rejecting them all
    // must not leave a `<w:tbl>` with properties and no rows behind.
    const insertedTable = doc(`<w:tr><w:trPr><w:ins ${REV}/></w:trPr>${insCell('Deposit')}</w:tr>`);
    expect(flattenTrackedChanges(insertedTable, 'rejectAll')).not.toMatch(/<w:tbl[\s>]/);
    expect(flattenTrackedChanges(insertedTable, 'acceptAll')).toMatch(/<w:tbl[\s>]/);
    const deletedTable = doc(`<w:tr><w:trPr><w:del ${REV}/></w:trPr>${delCell('Deposit')}</w:tr>`);
    expect(flattenTrackedChanges(deletedTable, 'acceptAll')).not.toMatch(/<w:tbl[\s>]/);
    expect(flattenTrackedChanges(deletedTable, 'rejectAll')).toContain('<w:t>Deposit</w:t>');
  });

  it('leaves an unmarked table exactly as it was', () => {
    const plain = doc(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell('Fee')}${cell('S$1')}</w:tr>`);
    expect(flattenTrackedChanges(plain, 'acceptAll')).toBe(flattenTrackedChanges(plain, 'rejectAll'));
    expect(flattenTrackedChanges(plain, 'acceptAll')).toContain('<w:cantSplit/>');
  });
});
