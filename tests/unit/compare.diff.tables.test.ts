import { describe, it, expect } from 'vitest';

import { diffTables } from '@features/compare/diff/index.js';

const AUTHOR = 'Test';
const DATE = '2026-06-08T00:00:00.000Z';

const cell = (t: string): string => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const row = (...cs: string[]): string => `<w:tr>${cs.map(cell).join('')}</w:tr>`;
const tbl = (...rs: string[]): string => `<w:tbl>${rs.join('')}</w:tbl>`;
const doc = (t: string): string => `<w:document><w:body>${t}</w:body></w:document>`;

describe('diffTables — T10 (added column is structural, not a rewrite)', () => {
  it('surfaces a new column as structural insert ops for the new cells', () => {
    const orig = doc(tbl(row('A1', 'B1'), row('A2', 'B2')));
    const rev = doc(tbl(row('A1', 'B1', 'C1'), row('A2', 'B2', 'C2')));

    const ops = diffTables(orig, rev, AUTHOR, DATE);
    const inserts = ops.filter((o) => o.type === 'insert');
    const equals = ops.filter((o) => o.type === 'equal');

    // The two new cells are structural inserts; the four old cells stay equal.
    expect(inserts).toHaveLength(2);
    expect(equals).toHaveLength(4);
    for (const ins of inserts) {
      expect(ins.classification).toBe('structural');
      expect(ins.partId).toMatch(/^tableCell:r\d+c\d+$/);
    }
    // Not a wholesale rewrite: no replace/delete of the unchanged cells.
    expect(ops.some((o) => o.type === 'replace')).toBe(false);
    expect(ops.some((o) => o.type === 'delete')).toBe(false);
  });

  it('reports a changed cell value as replace (word diff deferred)', () => {
    const orig = doc(tbl(row('A1', 'S$1,250,000')));
    const rev = doc(tbl(row('A1', 'S$1,500,000')));

    const ops = diffTables(orig, rev, AUTHOR, DATE);
    const replaces = ops.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(replaces[0]!.classification).toBe('substantiveText');
    expect(replaces[0]!.wordDiffs).toBeUndefined();
  });

  it('reports an added row as structural', () => {
    const orig = doc(tbl(row('A1', 'B1')));
    const rev = doc(tbl(row('A1', 'B1'), row('A2', 'B2')));

    const ops = diffTables(orig, rev, AUTHOR, DATE);
    const inserts = ops.filter((o) => o.type === 'insert');
    expect(inserts.length).toBeGreaterThanOrEqual(1);
    for (const ins of inserts) expect(ins.classification).toBe('structural');
  });
});
