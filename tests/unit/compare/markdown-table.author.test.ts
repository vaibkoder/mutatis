/**
 * ADR-0027 — structured table redline author unit tests.
 * Covers MT1–MT5 at the author level and the ADR-0008 never-suppress guarantee.
 */
import { describe, it, expect } from 'vitest';

import { makeAuthorContext } from '@features/compare/output/author.js';
import {
  authorTableReplace,
  authorTableInsert,
  authorTableDelete,
} from '@features/compare/output/author-table.js';
import { toRowModels } from '@features/compare/io/text/extract-text-records.js';

const AUTHOR = 'Mutatis';
const DATE = '2026-06-19T00:00:00Z';
const ctx = () => makeAuthorContext(AUTHOR, DATE);

const count = (xml: string, re: RegExp): number => (xml.match(re) ?? []).length;

// A real revision open-tag — `\b` excludes the `<w:insideH>/<w:insideV>` table
// BORDER elements, whose names contain the substring "ins".
const ANY_MARK = /<w:ins\b|<w:del\b|<w:cellIns\b|<w:cellDel\b/;

describe('authorTableReplace', () => {
  it('MT1 — identical tables produce no revision markup', () => {
    const g = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
    const out = authorTableReplace(g, g, ctx());
    expect(out).toContain('<w:tbl>');
    expect(out).not.toMatch(ANY_MARK);
  });

  it('MT2 — a changed cell is struck (old) and inserted (new); siblings stay clean', () => {
    const orig = toRowModels([
      ['Item', 'Amount'],
      ['Consideration', 'S$1,250,000'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount'],
      ['Consideration', 'S$1,500,000'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    expect(out).toMatch(/<w:del\b[^>]*>.*?<w:delText[^>]*>S\$1,250,000<\/w:delText>/s);
    expect(out).toMatch(/<w:ins\b[^>]*>.*?<w:t[^>]*>S\$1,500,000<\/w:t>/s);
    // The unchanged label is present exactly once, with no run-level markup.
    expect(out).toContain('Consideration');
    expect(count(out, /S\$1,250,000/g)).toBe(1);
    expect(count(out, /S\$1,500,000/g)).toBe(1);
    // The header cell is emitted clean (a plain run, no revision wrapper).
    expect(out).toMatch(/<w:r><w:t[^>]*>Amount<\/w:t><\/w:r>/);
  });

  it('MT3 — an added column becomes per-cell <w:cellIns>, existing columns clean', () => {
    const orig = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount', 'Note'],
      ['Fee', '100', 'urgent'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    expect(count(out, /<w:cellIns\b/g)).toBe(2); // one per row (header + body)
    expect(out).toMatch(/<w:ins\b[^>]*>.*?<w:t[^>]*>Note<\/w:t>/s);
    expect(out).toMatch(/<w:ins\b[^>]*>.*?<w:t[^>]*>urgent<\/w:t>/s);
    // No column was deleted, and existing values are not struck.
    expect(out).not.toMatch(/<w:cellDel\b/);
    expect(out).not.toMatch(/<w:delText/);
  });

  it('MT4 — a removed column becomes per-cell <w:cellDel>', () => {
    const orig = toRowModels([
      ['Item', 'Amount', 'Note'],
      ['Fee', '100', 'urgent'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    expect(count(out, /<w:cellDel\b/g)).toBe(2);
    expect(out).toMatch(/<w:delText[^>]*>Note<\/w:delText>/);
    expect(out).toMatch(/<w:delText[^>]*>urgent<\/w:delText>/);
    expect(out).not.toMatch(/<w:cellIns\b/);
  });

  it('MT5 — an added row carries a <w:trPr> insertion mark + inserted cells', () => {
    const orig = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
      ['Tax', '20'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    expect(out).toMatch(/<w:trPr><w:ins\b[^>]*\/><\/w:trPr>/);
    expect(out).toMatch(/<w:ins\b[^>]*>.*?<w:t[^>]*>Tax<\/w:t>/s);
    expect(out).toMatch(/<w:ins\b[^>]*>.*?<w:t[^>]*>20<\/w:t>/s);
  });

  it('MT5 — a removed row carries a <w:trPr> deletion mark + deleted cells', () => {
    const orig = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
      ['Tax', '20'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    expect(out).toMatch(/<w:trPr><w:del\b[^>]*\/><\/w:trPr>/);
    expect(out).toMatch(/<w:delText[^>]*>Tax<\/w:delText>/);
  });

  it('ADR-0008 — every original and revised cell value survives into the output', () => {
    const orig = toRowModels([
      ['Item', 'Amount'],
      ['Fee', '100'],
      ['Old', 'x'],
    ]);
    const rev = toRowModels([
      ['Item', 'Amount', 'Note'],
      ['Fee', '150', 'see s.4'],
    ]);
    const out = authorTableReplace(orig, rev, ctx());
    for (const v of ['Item', 'Amount', 'Fee', '100', '150', 'Old', 'Note', 'see s.4']) {
      expect(out).toContain(v);
    }
  });
});

describe('authorTableInsert / authorTableDelete', () => {
  it('a wholly inserted table marks every row + wraps every cell in <w:ins>', () => {
    const g = toRowModels([
      ['A', 'B'],
      ['1', '2'],
    ]);
    const out = authorTableInsert(g, ctx());
    expect(count(out, /<w:trPr><w:ins\b[^>]*\/><\/w:trPr>/g)).toBe(2);
    expect(out).not.toMatch(/<w:del\b/);
    expect(out).toMatch(/<w:ins\b/);
  });

  it('a wholly deleted table marks every row + wraps every cell in <w:del>', () => {
    const g = toRowModels([
      ['A', 'B'],
      ['1', '2'],
    ]);
    const out = authorTableDelete(g, ctx());
    expect(count(out, /<w:trPr><w:del\b[^>]*\/><\/w:trPr>/g)).toBe(2);
    expect(out).not.toMatch(/<w:ins\b/);
    expect(out).toMatch(/<w:delText[^>]*>A<\/w:delText>/);
  });
});
