import { describe, it, expect } from 'vitest';
import { runDefinedTerms } from '@features/defined-terms/index.js';
import { buildDefinitionsMap } from '@features/defined-terms/definitions.js';
import { mergeTableDefinitions } from '@features/defined-terms/tables.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

let seq = 0;
function body(text: string): ParaNode {
  const i = seq++;
  return { id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}
function cell(text: string, tableId: string, row: number, col: number): ParaNode {
  const i = seq++;
  return { id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'tableCell',
    tableInfo: { tableId, row, col }, getRange: () => { throw new Error('not used'); } };
}

function byCategory(fs: ReturnType<typeof runDefinedTerms>, suffix: string) {
  return fs.filter(f => f.category === `defined-terms/${suffix}`);
}

describe('table-format definitions (T4, T8)', () => {
  it('merges a quoted single-row definition table into the map', () => {
    const map = buildDefinitionsMap([], [], SG_PROFILE);
    mergeTableDefinitions(map, [
      cell('"Group"', 'table-0', 0, 0),
      cell('the Company and its subsidiaries', 'table-0', 0, 1),
    ], SG_PROFILE);
    const e = map.get('group');
    expect(e?.definitions[0]?.source).toBe('table');
  });

  it('T4 — table definition + prose definition → defined-multiple-times', () => {
    seq = 0;
    const fs = runDefinedTerms([
      cell('"Group"', 'table-0', 0, 0),
      cell('the Company and its subsidiaries', 'table-0', 0, 1),
      body('"Group" means the Company and each subsidiary undertaking.'),
      body('The Group shall consolidate its accounts.'),
    ], SG_PROFILE);
    const c = byCategory(fs, 'defined-multiple');
    expect(c).toHaveLength(1);
    expect(c[0]?.title).toContain('Group');
  });

  it('T8 — Schedule table with a leading numbering column (firstContentCol=1)', () => {
    seq = 0;
    const fs = runDefinedTerms([
      // numbering | term | definition
      cell('1.', 'table-0', 0, 0),
      cell('"Long-Stop Date"', 'table-0', 0, 1),
      cell('31 December 2026', 'table-0', 0, 2),
      cell('2.', 'table-0', 1, 0),
      cell('"Completion Date"', 'table-0', 1, 1),
      cell('15 January 2027', 'table-0', 1, 2),
      body('The Long-Stop Date and the Completion Date are key milestones.'),
    ], SG_PROFILE);
    expect(byCategory(fs, 'used-not-defined')).toHaveLength(0);
    expect(byCategory(fs, 'defined-not-used')).toHaveLength(0);
    expect(byCategory(fs, 'defined-multiple')).toHaveLength(0);
  });

  it('splits a multi-term cell ("Buyer" and "Purchaser")', () => {
    const map = buildDefinitionsMap([], [], SG_PROFILE);
    mergeTableDefinitions(map, [
      cell('"Buyer" and "Purchaser"', 'table-0', 0, 0),
      cell('Acme Pte. Ltd.', 'table-0', 0, 1),
      cell('"Seller"', 'table-0', 1, 0),
      cell('Beta Pte. Ltd.', 'table-0', 1, 1),
    ], SG_PROFILE);
    expect(map.has('buyer')).toBe(true);
    expect(map.has('purchaser')).toBe(true);
  });
});
