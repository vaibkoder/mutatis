/**
 * ADR-0046 — a Word table stays a table in a Word-vs-Word
 * redline, and is revised inside its cells, on its rows, or — where the
 * structure changed — as a row out and a row in. Driven end to end through
 * `runPipeline`, the path every add-in user takes.
 *
 * Every case is also held to WT10: accepting every revision spells the revised
 * document and rejecting every one the original — rows, cells, words and the
 * objects that carry no text (`signature`).
 */
import { describe, it, expect, afterEach, vi } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline, type PipelineResult } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import { flattenTrackedChanges } from '@features/compare/preprocess/flatten.js';
import { authorWordTableChange, signature } from '@features/compare/output/author-word-table.js';
import { authorBlocks, makeAuthorContext } from '@features/compare/output/author.js';
import { analyseWordTables, diffWordTable } from '@features/compare/diff/word-table.js';
import { extractParaRecords, groupTableUnits } from '@features/compare/preprocess/index.js';
import type { DiffOp } from '@features/compare/model.js';
import { h, render } from 'preact';
import { ChangeList, applyFilters, previewText } from '@features/compare/components/ChangeList.js';
import { compareDriverState } from '@features/compare/driver.js';
import { tableAnchorDiffs } from '@features/compare/diff/table-summary.js';
import { setActiveHost, _resetActiveHost, type HostPort } from '@shared/host/index.js';
import {
  FEE_ROWS, countTag, para, table, tableBundle, type Cell,
} from '../../fixtures/word-table-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);

async function compare(origBody: string, revBody: string): Promise<PipelineResult> {
  return runPipeline({
    originalParts: extractParts(tableBundle(origBody)),
    revisedParts: extractParts(tableBundle(revBody)),
    config: CONFIG,
    author: 'Mutatis',
    date: '2026-10-18T00:00:00Z',
  });
}

/** WT10 — accept-all is the revised body and reject-all the original, word for word, cell for cell. */
function expectReconstructs(result: PipelineResult, origBody: string, revBody: string): void {
  const authored = result.authoredBodyXml;
  expect(signature(flattenTrackedChanges(authored, 'acceptAll')), 'accept-all ≠ revised').toBe(signature(revBody));
  expect(signature(flattenTrackedChanges(authored, 'rejectAll')), 'reject-all ≠ original').toBe(signature(origBody));
}

/** Every `w:id` minted in the redline is distinct. */
function expectUniqueIds(xml: string): void {
  const ids = [...xml.matchAll(/<w:(?:ins|del|cellIns|cellDel)\s[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
  expect(new Set(ids).size).toBe(ids.length);
}

const changed = (ops: readonly DiffOp[]): DiffOp[] => ops.filter((o) => o.type !== 'equal');
const fees = (edit?: (rows: string[][]) => void): string => {
  const rows = FEE_ROWS.map((r) => [...r]);
  edit?.(rows);
  return table(rows);
};
/** The `<w:tc>` XML holding `text`, in document order. */
function cellHolding(xml: string, text: string): string {
  const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];
  const found = cells.find((c) => c.includes(text));
  expect(found, `no cell holds ${text}`).toBeDefined();
  return found!;
}

describe('ADR-0046 — a Word table stays a table', () => {
  it('WT1: an unchanged table is re-emitted byte for byte, with no revision in it', async () => {
    const origBody = para('The fees are set out below.') + fees() + para('Clause one applies.');
    const revBody = para('The fees are set out below.') + fees() + para('Clause two applies.');
    const result = await compare(origBody, revBody);
    const unit = result.revisedParas.find((p) => p.wordTable !== undefined);
    expect(unit).toBeDefined();
    expect(result.authoredBodyXml).toContain(unit!.rawXml);
    expect(countTag(unit!.rawXml, 'w:ins') + countTag(unit!.rawXml, 'w:del')).toBe(0);
    const tableOp = result.ops.find((o) => o.revRange && result.revisedParas[o.revRange.start] === unit);
    expect(tableOp?.type).toBe('equal');
    expect(result.summary.total).toBe(1);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT2: one changed cell is revised inside that cell, insertion first; the rest is untouched', async () => {
    const origBody = fees();
    const revBody = fees((rows) => {
      rows[2]![1] = 'S$4,500';
    });
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:tbl')).toBe(1);
    expect(countTag(xml, 'w:tr')).toBe(FEE_ROWS.length);
    const cell = cellHolding(xml, 'S$4,500');
    expect(cell).toContain('S$4,000');
    expect(cell.indexOf('<w:ins ')).toBeLessThan(cell.indexOf('<w:del ')); // ADR-0043
    // One paragraph: a cell's last paragraph mark cannot be revised, so a value
    // struck in one paragraph and inserted in another would leave a blank line.
    expect(countTag(cell, 'w:p')).toBe(1);
    for (const untouched of ['Arrangement fee', 'S$12,000', 'Commitment fee']) {
      expect(cellHolding(xml, untouched)).not.toMatch(/<w:(ins|del)\s/);
    }
    const ops = changed(result.ops);
    expect(ops).toHaveLength(1);
    expect(ops[0]!.type).toBe('replace');
    expect(ops[0]!.tableDiff).toBeDefined();
    expect(ops[0]!.classification).not.toBe('formatting');
    expect(ops[0]!.classification).not.toBe('whitespacePunct');
    expect(result.summary.total).toBe(1);
    expectUniqueIds(xml);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT3: a row inserted among edited rows is one inserted row; no other row is re-paired', async () => {
    const origBody = fees();
    const rows = FEE_ROWS.map((r) => [...r]);
    rows[1]![1] = 'S$12,500';
    rows[3]![1] = '0.40%';
    rows.splice(2, 0, ['Security trustee fee', 'S$2,000', 'annually']);
    const revBody = table(rows);
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    const rowMarks = (kind: string): number => (xml.match(new RegExp(`<w:trPr>(?:(?!</w:trPr>)[\\s\\S])*<w:${kind}\\s`, 'g')) ?? []).length;
    expect(rowMarks('ins')).toBe(1);
    expect(rowMarks('del')).toBe(0);
    expect(countTag(xml, 'w:tr')).toBe(rows.length);
    // The edited rows were revised in place, not struck and re-inserted.
    expect(cellHolding(xml, 'Arrangement fee')).not.toMatch(/<w:(ins|del)\s/);
    expect(cellHolding(xml, 'S$12,500')).toContain('S$12,000');
    expectUniqueIds(xml);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT4: a deleted row carries a row deletion and keeps its own row and cell properties', async () => {
    const rows: Cell[][] = FEE_ROWS.map((r) => [...r]);
    const origBody = table(rows, { trPr: '<w:cantSplit/>' });
    const revBody = table(rows.filter((_r, i) => i !== 2), { trPr: '<w:cantSplit/>' });
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    const deleted = (xml.match(/<w:tr>(?:(?!<\/w:tr>)[\s\S])*Agency fee[\s\S]*?<\/w:tr>/) ?? [''])[0];
    expect(deleted).toMatch(/<w:trPr><w:cantSplit\/><w:del\s/);
    expect(deleted).toContain('<w:delText');
    expect(deleted).not.toContain('<w:t>');
    expectReconstructs(result, origBody, revBody);
  });

  it('WT5: a multi-paragraph cell and a nested table are revised in place at both levels', async () => {
    const nested = (amount: string): string => table([['Tranche', 'Amount'], ['A', amount]]);
    const body = (second: string, amount: string): string =>
      table([
        ['Party', { paras: ['The Borrower', second] }],
        ['Schedule', { xml: nested(amount) + para('') }],
      ]);
    const origBody = body('acting through its branch.', 'S$1,000,000');
    const revBody = body('acting through its head office.', 'S$1,250,000');
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:tbl')).toBe(2);
    const borrower = (xml.match(/<w:p>(?:(?!<\/w:p>)[\s\S])*The Borrower[\s\S]*?<\/w:p>/) ?? [''])[0];
    expect(borrower).not.toMatch(/<w:(ins|del)\s/);
    expect(cellHolding(xml, 'head office.')).toContain('branch.');
    // The nested table's changed cell is revised inside the nested table.
    const inner = xml.slice(xml.indexOf('<w:tbl>', 10));
    expect(cellHolding(inner, 'S$1,250,000')).toContain('S$1,000,000');
    expectUniqueIds(xml);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT6: cells merged with gridSpan read as the old row out and the new row in, every word kept', async () => {
    const origBody = table([['Item', 'Amount'], ['Arrangement fee', 'S$12,000']]);
    const revBody = table([['Item', 'Amount'], [{ paras: ['Arrangement fee S$12,000'], gridSpan: 2 }]], { gridCols: 2 });
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:tbl')).toBe(1);
    expect(xml).toMatch(/<w:trPr><w:del\s/);
    expect(xml).toMatch(/<w:trPr><w:ins\s/);
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(changed(result.ops)[0]!.classes).toContain('structural');
    expectReconstructs(result, origBody, revBody);
  });

  it('WT6: merges moved within a row of the same cell count read as the old row out and the new row in', async () => {
    // Same number of cells, so only the spans say the layout changed — and the
    // self-check compares merges too, or a revision could carry the revised
    // merges into the rejected table unseen.
    const origBody = table([[{ paras: ['Fees'], gridSpan: 2 }, 'Total'], ['A', 'B', 'C']]);
    const revBody = table([['Fees', { paras: ['Total'], gridSpan: 2 }], ['A', 'B', 'C']]);
    const result = await compare(origBody, revBody);
    expect(tableOp(result).tableDiff!.rows.map((r) => r.kind)).toEqual(['delete', 'insert', 'match']);
    expect(countTag(result.authoredBodyXml, 'w:tbl'), 'fell back to whole table out and in').toBe(1);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT7: a column added or removed in every row is marked cell by cell, other columns unmarked', async () => {
    const withCurrency = FEE_ROWS.map((r, i) => [r[0]!, i === 0 ? 'Currency' : 'SGD', r[1]!, r[2]!]);
    for (const [origRows, revRows, mark] of [
      [FEE_ROWS, withCurrency, 'w:cellIns'],
      [withCurrency, FEE_ROWS, 'w:cellDel'],
    ] as const) {
      const origBody = table(origRows.map((r) => [...r]));
      const revBody = table(revRows.map((r) => [...r]));
      const result = await compare(origBody, revBody);
      const xml = result.authoredBodyXml;
      expect(countTag(xml, mark)).toBe(FEE_ROWS.length);
      expect(xml).not.toMatch(/<w:trPr><w:(ins|del)\s/);
      for (const kept of ['Arrangement fee', 'S$12,000', 'quarterly']) {
        expect(cellHolding(xml, kept)).not.toMatch(/<w:(ins|del)\s/);
      }
      // The grid has one column per authored cell: four.
      expect(countTag(xml.slice(xml.indexOf('<w:tblGrid'), xml.indexOf('</w:tblGrid>')), 'w:gridCol')).toBe(4);
      expect(changed(result.ops)[0]!.classes).toContain('structural');
      expectUniqueIds(xml);
      expectReconstructs(result, origBody, revBody);
    }
  });

  it('WT7: a column added while a value in another column changed still revises that value in place', async () => {
    const origBody = fees();
    const revBody = table(FEE_ROWS.map((r, i) => [r[0]!, i === 0 ? 'Currency' : 'SGD', i === 2 ? 'S$4,500' : r[1]!, r[2]!]));
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:cellIns')).toBe(FEE_ROWS.length);
    const cell = cellHolding(xml, 'S$4,500');
    expect(cell).toContain('S$4,000');
    expect(countTag(cell, 'w:cellIns')).toBe(0);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT7: a column added while one row is deleted and another inserted still pairs every other row with itself', async () => {
    // A new column changes every row's hash, so without seeing past it the rows
    // pair by POSITION and the deleted row drags its neighbours into each other.
    const origRows = FEE_ROWS.map((r) => [...r]);
    const revRows = FEE_ROWS.map((r, i) => [`ref ${i}`, ...r]);
    revRows.splice(2, 1); // Agency fee out
    revRows.push(['ref 9', 'Prepayment fee', '1%', 'on demand']);
    const origBody = table(origRows);
    const revBody = table(revRows);
    const result = await compare(origBody, revBody);
    const diff = tableOp(result).tableDiff!;
    expect(diff.rows.map((r) => r.kind)).toEqual(['match', 'match', 'delete', 'match', 'insert']);
    expect(diff.columns?.map((c) => c.kind)).toEqual(['insert', 'match', 'match', 'match']);
    // Every surviving row's own cells are untouched: only the new column is marked.
    const xml = result.authoredBodyXml;
    expect(cellHolding(xml, 'Commitment fee')).not.toMatch(/<w:(ins|del|cellIns|cellDel)\s/);
    expectUniqueIds(xml);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT7: a column added to a table with a merged total row is still marked cell by cell in every other row', async () => {
    // A row merged across the table (a heading, a total) cannot take a column
    // map — but it used to stop the map for EVERY row, striking and re-inserting
    // the whole table.
    const total = (span: number): Cell[] => [{ paras: ['Total payable on demand'], gridSpan: span }];
    const origRows: Cell[][] = [...FEE_ROWS.map((r) => [...r]), total(3)];
    const revRows: Cell[][] = [...FEE_ROWS.map((r, i) => [...r, `note ${i}`]), total(4)];
    const origBody = table(origRows);
    const revBody = table(revRows);
    const result = await compare(origBody, revBody);
    const diff = tableOp(result).tableDiff!;
    expect(diff.columns?.map((c) => c.kind)).toEqual(['match', 'match', 'match', 'insert']);
    // Every full-width row takes the new column; the merged row, whose merge
    // widened, is a layout change and reads as out and in.
    expect(diff.rows.map((r) => r.kind)).toEqual(['match', 'match', 'match', 'match', 'delete', 'insert']);
    expect(countTag(result.authoredBodyXml, 'w:cellIns')).toBe(FEE_ROWS.length);
    expect(cellHolding(result.authoredBodyXml, 'Agency fee')).not.toMatch(/<w:(ins|del)\s/);
    expectUniqueIds(result.authoredBodyXml);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT7: the authored grid follows the column map even when the first row is a merged heading', async () => {
    // The grid builder used to check the FIRST row's cell count against the
    // grid, so a heading merged across row 0 sent it to the revised grid: one
    // column short of the rows that still carry the struck column's cells.
    const heading = (span: number): Cell[] => [{ paras: ['Fees payable'], gridSpan: span }];
    const origRows: Cell[][] = [heading(4), ...FEE_ROWS.map((r) => [...r, 'note'])];
    const revRows: Cell[][] = [heading(3), ...FEE_ROWS.map((r) => [...r])];
    const origBody = table(origRows);
    const revBody = table(revRows);
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(tableOp(result).tableDiff!.columns?.map((c) => c.kind)).toEqual(['match', 'match', 'match', 'delete']);
    const grid = countTag((xml.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) ?? [''])[0], 'w:gridCol');
    expect(grid).toBe(4);
    // Every row the map revised spans the whole grid, struck column included.
    const spans = (xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) ?? [])
      .filter((row) => row.includes('<w:cellDel'))
      .map((row) => (row.match(/<w:tc>/g) ?? []).length);
    expect(spans).toEqual(FEE_ROWS.map(() => 4));
    expectReconstructs(result, origBody, revBody);
  });

  it('a row whose cells match only in their short values still pairs with itself in an uneven gap', async () => {
    // Dice over the row's words is low when the long cell was rewritten, but
    // three of four cells are identical — the same row. Without the cell share
    // the gap's rows all read as out and in.
    const row = (k: string, long: string): string[] => [k, 'S$1,000', 'annually', long];
    const origRows = [['Item', 'Amount', 'Due', 'Basis'], row('Agency fee', 'payable to the Agent in advance on each anniversary of the Closing Date'), row('Commitment fee', 'calculated daily on the undrawn Commitments of each Lender')];
    const revRows = [['Item', 'Amount', 'Due', 'Basis'], row('Agency fee', 'settled by the Borrower against an invoice issued by the Facility Office'), ['Prepayment fee', '1%', 'on demand', 'see Schedule 4'], row('Commitment fee', 'accruing from signing until the end of the Availability Period')];
    const origBody = table(origRows);
    const revBody = table(revRows);
    const result = await compare(origBody, revBody);
    expect(tableOp(result).tableDiff!.rows.map((r) => r.kind)).toEqual(['match', 'match', 'insert', 'match']);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT7: a cell whose vertical merge changed takes its row out and in, even when a column map exists', async () => {
    const origRows: Cell[][] = FEE_ROWS.map((r) => [...r]);
    const revRows: Cell[][] = FEE_ROWS.map((r, i) => [...r, `note ${i}`]);
    // The Amount column merged down across the last two rows in the revision.
    revRows[2]![1] = { paras: ['S$4,000'], vMerge: 'restart' };
    revRows[3]![1] = { paras: [''], vMerge: 'continue' };
    origRows[3]![1] = '';
    const origBody = table(origRows);
    const revBody = table(revRows);
    const result = await compare(origBody, revBody);
    expect(tableOp(result).tableDiff!.rows.map((r) => r.kind)).toEqual(['match', 'match', 'delete', 'insert', 'delete', 'insert']);
    expect(countTag(result.authoredBodyXml, 'w:tbl'), 'fell back to whole table out and in').toBe(1);
    expectReconstructs(result, origBody, revBody);
  });

  it('an inserted row\'s mark goes before a property change the row already carries (CT_TrPr order)', async () => {
    const tracked = '<w:cantSplit/><w:trPrChange w:id="900" w:author="Counterparty" w:date="2026-01-01T00:00:00Z"><w:trPr/></w:trPrChange>';
    const rows: Cell[][] = FEE_ROWS.map((r) => [...r]);
    const origBody = table(rows.filter((_r, i) => i !== 2), { trPr: tracked });
    const revBody = table(rows, { trPr: tracked });
    const result = await compare(origBody, revBody);
    const inserted = (result.authoredBodyXml.match(/<w:tr>(?:(?!<\/w:tr>)[\s\S])*Agency fee[\s\S]*?<\/w:tr>/) ?? [''])[0];
    expect(inserted).toMatch(/<w:trPr><w:cantSplit\/><w:ins\s[^>]*\/><w:trPrChange\s/);
  });

  it('a row inserted from a cell that ends in a nested table is given the closing paragraph a cell needs', async () => {
    const nested = table([['Tranche', 'Amount'], ['A', 'S$1']]);
    const rows: Cell[][] = FEE_ROWS.map((r) => [...r]);
    const origBody = table(rows);
    const revBody = table([...rows, ['Schedule', { xml: para('See below.') + nested }, 'n/a']]);
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(xml).toContain('See below.');
    expect(xml, 'the nested table closes the cell').not.toMatch(/See below\.[\s\S]*?<\/w:tbl>\s*<\/w:tc>/);
    expect(xml).toMatch(/See below\.[\s\S]*?<\/w:tbl>\s*<w:p[\s/>]/);
  });

  it('a cell paragraph emptied in Word\'s short form is revised inside it, its footnote kept on reject', async () => {
    // Word writes a blank paragraph as `<w:p/>` — and a cell's LAST paragraph,
    // the one it must end with, is very often blank. The revision used to be
    // written outside any paragraph, losing the footnote and failing the
    // self-check (found by the sweep, seed 439 of 2000).
    const fn = '<w:p><w:r><w:t xml:space="preserve">see note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const origBody = table([['Item', { xml: para('Fee') + fn }]]);
    const revBody = table([['Item', { xml: para('Fee') + '<w:p/>' }]]);
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:tbl'), 'fell back to whole table out and in').toBe(1);
    expect(xml).not.toMatch(/<\/w:p><w:del\s/);
    expectReconstructs(result, origBody, revBody);
  });

  it('WT8: a style change on a cell paragraph alone is still reported, and nothing is marked', async () => {
    const body = (style?: string): string =>
      table([['Item', { xml: para('Agency fee', style) }]]);
    const origBody = body();
    const revBody = body('Heading2');
    const result = await compare(origBody, revBody);
    const ops = changed(result.ops);
    expect(ops).toHaveLength(1);
    expect(ops[0]!.type).toBe('property');
    expect(ops[0]!.propertyChange?.to.styleId).toBe('Heading2');
    expect(result.authoredBodyXml).not.toMatch(/<w:(ins|del)\s/);
    expect(result.authoredBodyXml).toContain('<w:pStyle w:val="Heading2"/>');
    expectReconstructs(result, origBody, revBody);
  });

  it('WT9: a whole table inserted or deleted keeps its structure, every row marked', async () => {
    const intro = para('The fees are set out below.');
    const tail = para('Clause one applies.');
    for (const [origBody, revBody, kind] of [
      [intro + tail, intro + fees() + tail, 'ins'],
      [intro + fees() + tail, intro + tail, 'del'],
    ] as const) {
      const result = await compare(origBody, revBody);
      const xml = result.authoredBodyXml;
      expect(countTag(xml, 'w:tbl')).toBe(1);
      const marks = xml.match(new RegExp(`<w:trPr><w:${kind}\\s`, 'g')) ?? [];
      expect(marks).toHaveLength(FEE_ROWS.length);
      expectUniqueIds(xml);
      expectReconstructs(result, origBody, revBody);
    }
  });

  it('WT9: a table paired with a paragraph is the table out and the paragraph in', async () => {
    const origBody = para('Intro.') + fees() + para('End.');
    const revBody = para('Intro.') + para('The fees are as agreed in the Fee Letter.') + para('End.');
    const result = await compare(origBody, revBody);
    expect(countTag(result.authoredBodyXml, 'w:tbl')).toBe(1);
    // The struck table is the original's own XML, properties and all — never
    // the Markdown path's synthesised table (ADR-0046 § Rejected alternatives).
    expect(result.ops.map((o) => o.type)).toEqual(['equal', 'replace', 'equal']);
    const struck = (result.authoredBodyXml.match(/<w:tbl>[\s\S]*<\/w:tbl>/) ?? [''])[0];
    const tblPr = (origBody.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/) ?? ['?'])[0];
    expect(struck).toContain(tblPr);
    expect(struck.split('<w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>').length - 1).toBe(FEE_ROWS.length * 3);
    expectReconstructs(result, origBody, revBody);
  });

  it('a table the model cannot re-assemble falls back to the whole table out and in, never a lost cell', async () => {
    // A cell's paragraph wrapped in a block-level content control.
    const wrapped = (amount: string): string =>
      table([['Item', { xml: `<w:sdt><w:sdtPr/><w:sdtContent>${para(amount)}</w:sdtContent></w:sdt>` }]]);
    const origBody = wrapped('S$4,000');
    const revBody = wrapped('S$4,500');
    const result = await compare(origBody, revBody);
    const xml = result.authoredBodyXml;
    expect(countTag(xml, 'w:sdt')).toBe(2);
    expect(xml).toMatch(/<w:trPr><w:del\s/);
    expect(xml).toMatch(/<w:trPr><w:ins\s/);
    expectReconstructs(result, origBody, revBody);
  });

  it('a row or cell inside a wrapper the model does not represent is read through, compared, and kept', async () => {
    // A repeating-section content control round <w:tr>, custom XML round a row,
    // a content control round <w:tc>: a unit's hash is built from the cells the
    // model reads, and the rows inside such a wrapper used to be skipped, so an
    // edit there reported NOTHING (summary.total 0 — ADR-0008, found by review).
    // The wrapper is now transparent: its rows and cells are in the model, and a
    // changed table is authored whole from its XML, wrapper and all.
    const sdt = (inner: string): string => `<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
    const custom = (inner: string): string => `<w:customXml w:element="fee">${inner}</w:customXml>`;
    const rowWrapped = (wrap: (x: string) => string, amount: string): string =>
      table(FEE_ROWS.map((r) => [...r])).replace(/<w:tr>(?:(?!<\/w:tr>)[\s\S])*Agency fee[\s\S]*?<\/w:tr>/, (row) => wrap(row.replace('S$4,000', amount)));
    const cellWrapped = (amount: string): string =>
      table(FEE_ROWS.map((r) => [...r])).replace(/<w:tc>(?:(?!<\/w:tc>)[\s\S])*S\$4,000[\s\S]*?<\/w:tc>/, (cell) => sdt(cell.replace('S$4,000', amount)));
    const nestedWrapped = (amount: string): string =>
      table([['Schedule', { xml: rowWrapped(sdt, amount) + para('') }]]);
    const shapes: Record<string, (amount: string) => string> = {
      'content control round a row': (a) => rowWrapped(sdt, a),
      'custom XML round a row': (a) => rowWrapped(custom, a),
      'content control round a cell': cellWrapped,
      'a nested table with a wrapped row': nestedWrapped,
    };
    for (const [name, build] of Object.entries(shapes)) {
      const origBody = build('S$4,000');
      const revBody = build('S$9,000');
      expect(revBody, name).toContain('S$9,000');
      const result = await compare(origBody, revBody);
      expect(result.summary.total, name).toBeGreaterThan(0);
      expect(changed(result.ops).length, name).toBeGreaterThan(0);
      expect(signature(flattenTrackedChanges(result.authoredBodyXml, 'acceptAll')), `${name}: accept-all`).toBe(signature(revBody));
      expect(signature(flattenTrackedChanges(result.authoredBodyXml, 'rejectAll')), `${name}: reject-all`).toBe(signature(origBody));
      // Unchanged, it is still re-emitted as written.
      const same = await compare(origBody, origBody);
      expect(same.summary.total, `${name}: unchanged`).toBe(0);
      expect(signature(same.authoredBodyXml), `${name}: unchanged`).toBe(signature(origBody));
    }
  });

  it('a paragraph outside any cell leaves its table flat, so every change to it is reported', async () => {
    // Schema-invalid (Word does not write it), but a generator can: a paragraph
    // in a table-level wrapper, outside every row. A unit has no place for its
    // style, its footnotes or its position, and folding only its text into the
    // hash reported NOTHING for those (found in review). Left flat, as
    // before ADR-0046, each is reported.
    const rows = FEE_ROWS.map((r) => [...r]);
    const inner = (table(rows).match(/<w:tr>[\s\S]*<\/w:tr>/) ?? [''])[0];
    const head = (table(rows).match(/^<w:tbl>(<w:tblPr>[\s\S]*?<\/w:tblGrid>)/) ?? ['', ''])[1];
    const fn = '<w:r><w:footnoteReference w:id="1"/></w:r>';
    const stray = (p: string, after = false): string =>
      `<w:tbl>${head}<w:sdt><w:sdtContent>${after ? inner + p : p + inner}</w:sdtContent></w:sdt></w:tbl>`;
    const note = 'Fees exclude GST.';
    const cases: Record<string, [string, string]> = {
      'its words edited': [stray(para(note)), stray(para('Fees include GST.'))],
      'its style changed': [stray(para(note)), stray(para(note, 'Heading2'))],
      'a footnote added to it': [stray(para(note)), stray(para(note).replace('</w:p>', `${fn}</w:p>`))],
      'moved below the rows': [stray(para(note)), stray(para(note), true)],
    };
    // A CELL outside every row, and a paragraph inside the table's properties:
    // the cells carry the edit, and the walk never consumes what it cannot place.
    const cellsOnly = (amount: string): string =>
      `<w:tbl>${head}<w:sdt><w:sdtContent>${(inner.replace('S$4,000', amount).match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).join('')}</w:sdtContent></w:sdt></w:tbl>`;
    cases['cells outside any row, in a wrapper'] = [cellsOnly('S$4,000'), cellsOnly('S$9,000')];
    cases['a cell directly under the table'] = [
      `<w:tbl>${head}${(inner.match(/<w:tc>[\s\S]*?<\/w:tc>/) ?? [''])[0]}${inner}</w:tbl>`,
      `<w:tbl>${head}${(inner.match(/<w:tc>[\s\S]*?<\/w:tc>/) ?? [''])[0]}${inner.replace('S$4,000', 'S$9,000')}</w:tbl>`,
    ];
    cases['a paragraph inside the table properties'] = [
      stray('').replace('</w:tblPr>', `${para(note)}</w:tblPr>`).replace('S$4,000', 'S$4,000'),
      stray('').replace('</w:tblPr>', `${para(note)}</w:tblPr>`).replace('S$4,000', 'S$9,000'),
    ];
    // A well-formed table beside it stays a table: only the stray one is flat
    // (the grouping's consistency fallback would flatten EVERY table instead).
    const beside = para('Schedule 2.') + fees();
    for (const [name, [origBody, revBody]] of Object.entries(cases)) {
      const result = await compare(origBody + beside, revBody + beside);
      expect(result.revisedParas.filter((p) => p.wordTable !== undefined), `${name}: grouped`).toHaveLength(1);
      expect(result.revisedParas.find((p) => p.wordTable)!.normText, name).toContain('Arrangement fee');
      expect(result.summary.total, name).toBeGreaterThan(0);
      // Nothing dropped from the redline, and the change is a real one.
      for (const word of ['Arrangement fee', 'Schedule 2.']) expect(result.authoredBodyXml, `${name}: ${word}`).toContain(word);
      if (name !== 'its style changed') {
        expect(changed(result.ops).some((o) => o.classification !== 'formatting'), `${name}: filed as formatting only`).toBe(true);
      }
    }
  });

  it('a footnote added or removed in a table is still marked, wherever it sits', async () => {
    // A table property op with no object change is authored as written; this
    // pins the other half — one WITH an object change is marked (ADR-0045 §
    // Amendment), including behind an earlier style change in a nested table,
    // whose op carries only its first change's flags.
    const fnPara = (text: string): string => para(text).replace('</w:p>', '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>');
    const nested = (a: string, b: string): string => table([[{ xml: a }, { xml: b }]]);
    const cases: Record<string, [string, string]> = {
      'in a cell': [table([['Fee', { xml: para('S$1') }]]), table([['Fee', { xml: fnPara('S$1') }]])],
      'in a nested table': [
        table([['Schedule', { xml: nested(para('A'), para('S$1')) + para('') }]]),
        table([['Schedule', { xml: nested(para('A'), fnPara('S$1')) + para('') }]]),
      ],
      'beside a style change in another cell': [
        table([[{ xml: para('Fee') }, { xml: para('S$1') }]]),
        table([[{ xml: para('Fee', 'Heading2') }, { xml: fnPara('S$1') }]]),
      ],
      'behind a style change in a nested table': [
        table([['Schedule', { xml: nested(para('A'), para('S$1')) + para('') }]]),
        table([['Schedule', { xml: nested(para('A', 'Heading2'), fnPara('S$1')) + para('') }]]),
      ],
    };
    for (const [name, [origBody, revBody]] of Object.entries(cases)) {
      for (const [o, r] of [[origBody, revBody], [revBody, origBody]] as const) {
        const result = await compare(o, r);
        expect(result.authoredBodyXml, `${name}: unmarked`).toMatch(/<w:(ins|del)\s/);
        expectReconstructs(result, o, r);
      }
    }
  });

  it('a style-only change in a table authored whole is not marked, as a paragraph\'s is not', async () => {
    // A wrapped cell makes the table `supported: false`, authored whole when it
    // changes — but a style alone is a property op the list files as
    // formatting, and the redline must not strike and re-insert the table for it.
    const sdt = (inner: string): string => `<w:sdt><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
    const body = (style?: string): string =>
      table([['Item', { xml: sdt(para('S$4,000', style)) }]]);
    const result = await compare(body(), body('Heading2'));
    const op = tableOp(result);
    expect(op.type).toBe('property');
    expect(result.authoredBodyXml).not.toMatch(/<w:(ins|del)\s/);
    expect(countTag(result.authoredBodyXml, 'w:tbl')).toBe(1);
  });

  it('numbered and signature paragraphs inside cells keep their record flags', async () => {
    const body = table([['Signed by', 'Director']]);
    const result = await compare(body, body);
    const unit = result.revisedParas.find((p) => p.wordTable !== undefined)!;
    expect(unit.isSignatureBlock).toBe(true);
    expect(unit.wordTable!.rows[0]!.cells[1]!.content[0]!.isTableCell).toBe(true);
  });
});

/** The one table op in a result. */
function tableOp(result: PipelineResult): DiffOp {
  const op = result.ops.find((o) => o.tableDiff !== undefined);
  expect(op, 'no table op').toBeDefined();
  return op!;
}

describe('WT11 — a table change is previewed and navigated from inside a cell', () => {
  afterEach(() => {
    _resetActiveHost();
    vi.restoreAllMocks();
  });

  it('the change-list row shows the changed cell\'s words, never a blank or the whole table', async () => {
    const result = await compare(fees(), fees((rows) => { rows[2]![1] = 'S$4,500'; }));
    const text = previewText(tableOp(result));
    expect(text).toContain('S$4,500');
    expect(text).toContain('S$4,000');
    expect(text).not.toContain('Arrangement fee'); // the table's own text spans every cell
    expect(text).not.toContain('more in this table');
  });

  it('a table with several changes says how many more there are behind the first', async () => {
    const result = await compare(fees(), fees((rows) => {
      rows[1]![1] = 'S$12,500';
      rows[3]![2] = 'monthly';
      rows.push(['Prepayment fee', '1%', 'on demand']);
    }));
    const text = previewText(tableOp(result));
    expect(text).toContain('S$12,500');
    expect(text).toContain('(+2 more in this table)');
  });

  it('a changed table\'s row carries the always-substantive badge, as a Markdown table cell does (§7.3)', async () => {
    const result = await compare(fees(), fees((rows) => { rows[2]![1] = 'S$4,500'; }));
    const host = document.createElement('div');
    document.body.appendChild(host);
    try {
      render(h(ChangeList, {
        ops: [tableOp(result)],
        filters: { hideFormattingOnly: true, collapseRenumbering: true },
        cascadeLabels: new Map(),
        onGoToChange: () => {},
      }), host);
      expect(host.querySelector('.dp-compare__change-badge')?.textContent).toBe('always substantive');
    } finally {
      render(null, host);
      host.remove();
    }
  });

  it('a changed table row is never hidden by Hide formatting-only, as its badge says', async () => {
    const result = await compare(fees(), fees((rows) => { rows[2]![1] = 'S$4,500'; }));
    // Forced to read as formatting-only: the badge promises it stays visible.
    const op: DiffOp = { ...tableOp(result), classification: 'formatting', classes: ['formatting'] };
    expect(applyFilters([op], { hideFormattingOnly: true, collapseRenumbering: true })).toHaveLength(1);
  });

  it('the "+N more" count reaches into a nested table', async () => {
    const nested = (a: string, b: string): string => table([['Tranche', 'Amount'], [a, b]]);
    const body = (x: string, a: string, b: string): string => table([['Schedule', { xml: para(x) + nested(a, b) + para('') }]]);
    const result = await compare(body('Fees', 'A', 'S$1'), body('Fees payable', 'B', 'S$2'));
    expect(previewText(tableOp(result))).toContain('(+2 more in this table)');
  });

  it('a table that changed only by rows names its shape', async () => {
    const result = await compare(fees(), fees((rows) => { rows.splice(2, 1); }));
    expect(previewText(tableOp(result))).toBe('Table: 1 row deleted');
  });

  it('a deleted row aims at a row still in the document, never the gone words or the table text', async () => {
    const result = await compare(fees(), fees((rows) => { rows.splice(2, 1); }));
    const anchor = tableAnchorDiffs(tableOp(result), result.originalParas, result.revisedParas);
    expect(anchor).toEqual([{ op: 'eq', text: 'Item' }]);
  });

  it('Go to change searches for the changed cell\'s words, which Word can find', async () => {
    const findAndReveal = vi.fn().mockResolvedValue(undefined);
    setActiveHost({ writeBack: vi.fn().mockResolvedValue(undefined), findAndReveal } as unknown as HostPort);
    const file = (name: string, body: string): File =>
      new File([new Uint8Array(tableBundle(body).toUint8Array())], name);
    const d = compareDriverState;
    d.uploadDocx('original', file('orig.docx', fees()));
    d.uploadDocx('revised', file('rev.docx', fees((rows) => { rows[2]![1] = 'S$4,500'; })));
    await d.runCompare();
    const op = d.ops.value!.find((o) => o.tableDiff !== undefined);
    expect(op, 'no table op').toBeDefined();
    d.goToChange(op!.id);
    expect(findAndReveal).toHaveBeenCalledTimes(1);
    const anchor = findAndReveal.mock.calls[0]![0] as string;
    expect(anchor).toContain('4,500');
    expect(anchor).not.toContain('|');
  });
});

describe('ADR-0046 — the guards behind a table redline', () => {
  it('the self-check catches a cell author that loses a word, and falls back to the whole table out and in', async () => {
    const result = await compare(fees(), fees((rows) => { rows[2]![1] = 'S$4,500'; }));
    const op = tableOp(result);
    const orig = result.originalParas[op.origRange!.start]!;
    const rev = result.revisedParas[op.revRange!.start]!;
    const ctx = makeAuthorContext('Mutatis', '2026-10-18T00:00:00Z');
    const mark = (xml: string, kind: 'ins' | 'del'): string =>
      xml.replace(/<w:r>([\s\S]*?)<\/w:r>/g, (_m, inner: string) => {
        const body = kind === 'del' ? inner.replace(/<w:t(\s|>)/g, '<w:delText$1').replace(/<\/w:t>/g, '</w:delText>') : inner;
        return `<w:${kind} w:id="${ctx.nextId()}" w:author="x" w:date="d"><w:r>${body}</w:r></w:${kind}>`;
      });
    const xml = authorWordTableChange(op, orig, rev, ctx, {
      // A broken cell author: the new amount goes missing.
      authorBlocks: (ops, o, r) => authorBlocks(ops, o, r, ctx).replace('4,500', ''),
      markParagraph: mark,
    });
    expect(countTag(xml, 'w:tbl'), 'the broken output was kept').toBe(2);
    expect(signature(flattenTrackedChanges(xml, 'acceptAll'))).toBe(signature(rev.rawXml));
    expect(signature(flattenTrackedChanges(xml, 'rejectAll'))).toBe(signature(orig.rawXml));
  });

  it('a table whose every cell was rewritten is still revised cell by cell, never demoted to the whole table out and in', async () => {
    // ADR-0026's demotion strikes a dissimilar pair whole; for a table that
    // would be every row struck and re-inserted for a set of in-place edits.
    const origBody = table([['Lender', 'Agent'], ['Borrower', 'Guarantor']]);
    const revBody = table([['Arranger', 'Trustee'], ['Obligor', 'Surety']]);
    const result = await compare(origBody, revBody);
    expect(tableOp(result).tableDiff!.rows.map((r) => r.kind)).toEqual(['match', 'match']);
    expect(countTag(result.authoredBodyXml, 'w:tbl')).toBe(1);
    expectReconstructs(result, origBody, revBody);
  });

  it('the self-check sees a vertical merge that a wrongly paired row would carry into the rejected table', async () => {
    // No aligner path pairs rows whose merges differ (sameLayout, fitsColumns);
    // this simulates one that did. The words are identical both ways, so only
    // the merge in the signature can tell.
    const origBody = table([['Fee', 'S$1'], ['Fee', 'S$1']]);
    const revBody = table([[{ paras: ['Fee'], vMerge: 'restart' }, 'S$1'], [{ paras: ['Fee'], vMerge: 'continue' }, 'S$1']]);
    const result = await compare(origBody, revBody);
    const op = tableOp(result);
    const forced: DiffOp = {
      ...op,
      tableDiff: { rows: [0, 1].map((k) => ({ kind: 'match' as const, o: k, r: k, cells: [0, 1].map((c) => ({ kind: 'match' as const, o: c, r: c, ops: [] })) })) },
    };
    const ctx = makeAuthorContext('Mutatis', '2026-10-18T00:00:00Z');
    const orig = result.originalParas[op.origRange!.start]!;
    const rev = result.revisedParas[op.revRange!.start]!;
    const xml = authorWordTableChange(forced, orig, rev, ctx, {
      authorBlocks: (ops, o, r) => authorBlocks(ops, o, r, ctx),
      markParagraph: (x) => x,
    });
    expect(countTag(xml, 'w:tbl'), 'the merge-carrying output was kept').toBe(2);
  });

  it('a row or cell mark minted after a cell\'s content is filed under the TABLE op\'s classes', async () => {
    // A cell's content is authored op by op, each op setting the class its ids
    // are filed under; a row mark minted after it must be filed under the
    // table's own classes (structural), not the last cell op's.
    const result = await compare(fees(), fees((rows) => {
      rows[1]![1] = 'S$12,500';
      rows.push(['Prepayment fee', '1%', 'on demand']);
    }));
    const xml = result.authoredBodyXml;
    const rowId = /<w:trPr><w:ins w:id="(\d+)"/.exec(xml)?.[1];
    expect(rowId).toBeDefined();
    expect(result.classByRevId.get(rowId!)).toEqual(tableOp(result).classes);
    expect(result.classByRevId.get(rowId!)).toContain('structural');
  });

  it('a table property op never carries a cascade marker, so §6.1 cannot collapse a table', async () => {
    const body = (style: string): string => table([[{ xml: para('Fee', style) }, 'S$1']]);
    const result = await compare(body('Normal'), body('Heading2'));
    const original = result.originalParas;
    const revised = result.revisedParas;
    const op = result.ops.find((o) => o.revRange && revised[o.revRange.start]?.wordTable)!;
    const cascading: DiffOp = {
      ...op, id: 'x', type: 'property', classification: 'crossRefRenumber', classes: ['crossRefRenumber'],
      propertyChange: {
        styleId: false, numberingIdentity: false, numberLabel: true,
        from: { styleId: null, numberLabel: '(a)', ilvl: 0 }, to: { styleId: null, numberLabel: '(b)', ilvl: 0 },
        cascade: { from: '(a)', to: '(b)', shiftedIlvl: 0, shift: 1 },
      },
    };
    const [out] = await analyseWordTables([{ ...op, type: 'equal' }], original, revised, CONFIG, async () => [cascading]);
    expect(out!.type).toBe('property');
    expect(out!.propertyChange).toBeDefined();
    expect(out!.propertyChange!.cascade).toBeUndefined();
  });

  it('grouping returns the flat records untouched when the walk and the records disagree', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const xml = `<w:document ${W}><w:body>${para('Before.')}${fees()}${para('After.')}</w:body></w:document>`;
    const flat = extractParaRecords(xml, 'document', CONFIG);
    expect(groupTableUnits(flat, xml).some((r) => r.wordTable !== undefined)).toBe(true);
    // Fresh records: the call above renumbered its own on success.
    const short = extractParaRecords(xml, 'document', CONFIG).slice(0, -1);
    const indices = short.map((r) => r.index);
    expect(groupTableUnits(short, xml)).toBe(short);
    // Untouched, not just the same array: a walk that gave up used to have
    // renumbered the cell records already, and the flat path then read every
    // op's range from the wrong paragraph (found in review).
    expect(short.map((r) => r.index)).toEqual(indices);
  });

  it('only a cell whose content changed is analysed', async () => {
    const result = await compare(fees(), fees((rows) => { rows[2]![1] = 'S$4,500'; }));
    const op = tableOp(result);
    let calls = 0;
    await diffWordTable(
      result.originalParas[op.origRange!.start]!.wordTable!,
      result.revisedParas[op.revRange!.start]!.wordTable!,
      CONFIG,
      async () => { calls++; return []; },
    );
    expect(calls).toBe(1);
  });
});
