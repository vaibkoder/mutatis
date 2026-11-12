/**
 * Feature A — reference scanner tests (spec §5.1.2, §6, T4 acceptance).
 * Tests A6/A7/A8/A9 and regex unit tests.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { scanForReferences, resetMatchCounter } from '@features/cross-reference/scanner.js';
import { buildSectionIndex } from '@shared/section-index/builder.js';
import { UK_PROFILE, SG_PROFILE } from '@shared/config/index.js';
import { para } from '../../fixtures/cross-reference/paras.js';

beforeEach(() => resetMatchCounter());

function makeIndex(paras: ReturnType<typeof para>[]) {
  return buildSectionIndex(paras, UK_PROFILE);
}

describe('Feature A scanner — clause references', () => {
  it('detects "Clause 4.2" in prose', () => {
    const paras = [
      para(0, 'Payment', { listString: '4', listLevel: 0 }),
      para(1, '4.2', { listString: '4.2', listLevel: 1 }),
      para(2, 'As set out in Clause 4.2, the amount shall be paid…'),
    ];
    const index = makeIndex(paras);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'clause:4.2')).toBe(true);
  });

  it('detects "Clause 4.2(b)(i)" compound reference', () => {
    const paras = [
      para(0, 'See Clause 4.2(b)(i) for details.'),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'clause:4.2(b)(i)')).toBe(true);
  });

  it('A9: self-reference phrases not matched', () => {
    const paras = [
      para(0, 'This Agreement shall be governed by Singapore law.'),
      para(1, 'This Deed is made between the parties.'),
      para(2, 'This Clause does not apply after termination.'),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches).toHaveLength(0);
  });
});

describe('Feature A scanner — schedule references', () => {
  it('A6: Recital (B) reference recognised', () => {
    const recitals = [
      para(0, 'Recitals'),
      para(1, '(A) The Seller…'),
      para(2, '(B) The Buyer…'),
    ];
    const body = [
      para(3, 'As stated in Recital (B), the Buyer wishes to…'),
    ];
    const allParas = [...recitals, ...body];
    const index = buildSectionIndex(allParas, UK_PROFILE);
    const matches = scanForReferences(allParas, UK_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'recital:B')).toBe(true);
  });

  it('A7: "Schedules 2 to 5 (inclusive)" expands to 4 keys', () => {
    const paras = [
      para(0, 'Schedule 2', { styleName: 'Heading 1' }),
      para(1, 'Schedule 3', { styleName: 'Heading 1' }),
      para(2, 'Schedule 4', { styleName: 'Heading 1' }),
      para(3, 'Schedule 5', { styleName: 'Heading 1' }),
      para(4, 'See Schedules 2 to 5 (inclusive) for details.'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    const rangeMatch = matches.find(m => m.expandedKeys.length > 1);
    expect(rangeMatch?.expandedKeys).toContain('schedule:2');
    expect(rangeMatch?.expandedKeys).toContain('schedule:5');
  });

  it('A8: "Schedules 2 and 3" — list expansion', () => {
    const paras = [
      para(0, 'See Schedules 2 and 3 for the forms.'),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    const listMatch = matches.find(m => m.expandedKeys.length > 1);
    expect(listMatch?.expandedKeys).toContain('schedule:2');
    expect(listMatch?.expandedKeys).toContain('schedule:3');
  });

  it('A10: "Section 4.2" recognised and mapped to clause space (SG default)', () => {
    const paras = [
      para(0, 'Payment', { listString: '4.2', listLevel: 1 }),
      para(1, 'As set out in Section 4.2…'),
    ];
    const index = buildSectionIndex(paras, SG_PROFILE);
    const matches = scanForReferences(paras, SG_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'clause:4.2')).toBe(true);
  });

  it('A11: "Exhibit B" recognised via alias, mapped to schedule space', () => {
    const paras = [
      para(0, 'Schedule B', { styleName: 'Heading 1' }),
      para(1, 'See Exhibit B for the form of letter.'),
    ];
    const index = buildSectionIndex(paras, SG_PROFILE);
    const matches = scanForReferences(paras, SG_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'schedule:B')).toBe(true);
  });
});

describe('Feature A scanner — sentence-boundary false positives', () => {
  it('does not read "…in part. assign…" as a reference "part. a"', () => {
    const paras = [
      para(0, 'Neither Party may, in whole or in part. assign this MOU without consent.'),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches).toHaveLength(0);
  });

  it('does not treat a unit noun followed by an ordinary word as a reference', () => {
    const paras = [
      para(0, 'The schedule applies and the appendix illustrates the recital basis.'),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches).toHaveLength(0);
  });

  it('still matches a genuine "Part A" reference followed by a space', () => {
    const paras = [
      para(0, 'Part A', { styleName: 'Heading 1' }),
      para(1, 'The obligations in Part A survive termination.'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches.some(m => m.normalizedKey === 'part:A')).toBe(true);
  });
});

describe('Feature A scanner — TOC exclusion', () => {
  it('TOC paragraphs are not scanned for references', () => {
    const paras = [
      para(0, 'Clause 4 …… 12', { styleName: 'TOC 1' }),
    ];
    const index = buildSectionIndex([], UK_PROFILE);
    const matches = scanForReferences(paras, UK_PROFILE, index);
    expect(matches).toHaveLength(0);
  });
});
