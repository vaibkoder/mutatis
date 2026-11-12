/**
 * Feature B — static-numbering detector tests (spec §5.2, T6 acceptance).
 * Tests B1–B5 + cross-population.
 */
import { describe, it, expect } from 'vitest';
import { detectStaticNumbering, crossPopulateStaticNumbered } from '@features/cross-reference/static-numbering.js';
import { buildSectionIndex } from '@shared/section-index/builder.js';
import { UK_PROFILE } from '@shared/config/index.js';
import { para } from '../../fixtures/cross-reference/paras.js';

describe('Feature B — static numbering detector (layers a+b)', () => {
  it('B1: fully auto-numbered clauses → no static findings', () => {
    const paras = [
      para(0, 'Definitions', { listString: '1', listLevel: 0 }),
      para(1, 'In this Agreement…', { listString: '1.1', listLevel: 1 }),
      para(2, 'Payment', { listString: '2', listLevel: 0 }),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    expect(results.filter(r => r.verdict === 'static')).toHaveLength(0);
  });

  it('B2: hand-typed "4.2" with no list item → static, high confidence, low severity (→info)', () => {
    const paras = [
      para(0, '4.2 The Buyer shall pay the Purchase Price…'),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    const statics = results.filter(r => r.verdict === 'static');
    expect(statics).toHaveLength(1);
    expect(statics[0]?.confidence).toBe('high');
    // low → mapped to Finding Severity 'info'
    expect(statics[0]?.severity).toBe('info');
    expect(statics[0]?.leadingNumberText).toBe('4.2');
  });

  it('B3: static number on a Heading-styled paragraph → medium severity (→warning)', () => {
    const paras = [
      para(0, '4.2 Representations', { styleName: 'Heading 2' }),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    const statics = results.filter(r => r.verdict === 'static');
    expect(statics).toHaveLength(1);
    // medium → mapped to Finding Severity 'warning'
    expect(statics[0]?.severity).toBe('warning');
  });

  it('B4: duplicate-on-list — list item present AND duplicate typed number (medium→warning)', () => {
    // Same text as list string → duplicate
    const paras = [
      para(0, '4.2 The Buyer shall pay…', { listString: '4.2', listLevel: 1 }),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    const statics = results.filter(r => r.verdict === 'static');
    expect(statics).toHaveLength(1);
    expect(statics[0]?.severity).toBe('warning');
    expect(statics[0]?.hasListItem).toBe(true);
  });

  it('B5: no leading number → no finding (noNumber)', () => {
    const paras = [
      para(0, 'The Buyer agrees to purchase the property.'),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    expect(results).toHaveLength(0);
  });

  it('B5b: numeric-leading prose (bare integer + space) is NOT flagged as static', () => {
    // Regression for the false-positive surface flagged in correctness review.
    const paras = [
      para(0, '5 days after the date of this Agreement…'),
      para(1, 'the Companies Act 1976 provides that…'),
      para(2, '30 June 2026 is the Completion Date.'),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    expect(results.filter(r => r.verdict === 'static')).toHaveLength(0);
  });

  it('B7: TOC paragraph excluded from detection', () => {
    const paras = [
      para(0, 'Clause 4 …… 12', { styleName: 'TOC 1' }),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    expect(results).toHaveLength(0);
  });

  it('non-body containers are skipped', () => {
    const paras = [
      para(0, '4.2 Something', { container: 'tableCell' }),
      para(1, '4.3 Something', { container: 'footnote' }),
    ];
    const results = detectStaticNumbering(paras, UK_PROFILE);
    expect(results).toHaveLength(0);
  });
});

describe('Feature B — cross-population into Section Index (spec §5.2.4)', () => {
  it('static verdict paragraphs set isStaticNumbered=true on matching index node', () => {
    const paras = [
      para(0, '4.2 The Buyer shall pay…'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const results = detectStaticNumbering(paras, UK_PROFILE);
    crossPopulateStaticNumbered(results, index.byParaId);
    // The para was indexed by static fallback; cross-population confirms it.
    const node = index.resolve('clause:4.2')[0];
    expect(node?.isStaticNumbered).toBe(true);
  });
});
