import { describe, it, expect } from 'vitest';
import { scanAmounts } from '@features/factual-consistency/amounts.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function para(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}

function scan(text: string, profile = SG_PROFILE) {
  return scanAmounts([para('p0', text)], profile);
}

// ---------------------------------------------------------------------------
// §9.1 — Number/word value
// ---------------------------------------------------------------------------

describe('Amounts — §9.1 value checks', () => {
  it('Ten Thousand Dollars (S$10,000) — no finding', () => {
    expect(scan('Ten Thousand Dollars (S$10,000)')).toHaveLength(0);
  });

  it('Ten Thousand Dollars (S$1,000) — Critical value-mismatch', () => {
    const f = scan('Ten Thousand Dollars (S$1,000)');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
    expect(f[0]?.category).toBe('factual-consistency/amount-mismatch');
  });

  it('One Hundred and Fifty Thousand Dollars (S$150,000) — no finding (British "and")', () => {
    expect(scan('One Hundred and Fifty Thousand Dollars (S$150,000)')).toHaveLength(0);
  });

  it('Five Thousand Pounds (£5,000.00) — no finding (cent tolerance)', () => {
    expect(scan('Five Thousand Pounds (£5,000.00)', UK_PROFILE)).toHaveLength(0);
  });

  it('One Thousand Dollars and Fifty Cents (S$1,000.50) — no finding (sub-unit)', () => {
    expect(scan('One Thousand Dollars and Fifty Cents (S$1,000.50)')).toHaveLength(0);
  });

  it('Ten Thousand Dollars and No/100 (S$10,000) only — no finding (suffix stripped)', () => {
    expect(scan('Ten Thousand Dollars and No/100 (S$10,000) only')).toHaveLength(0);
  });

  it('Pounds Sterling £10,000 — no finding (GBP word + symbol)', () => {
    expect(scan('Pounds Sterling £10,000', UK_PROFILE)).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// §9.2 — Currency checks
// ---------------------------------------------------------------------------

describe('Amounts — §9.2 currency checks', () => {
  it('Ten Thousand Dollars (€10,000) — Critical currency-mismatch', () => {
    const f = scan('Ten Thousand Dollars (€10,000)');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
    expect(f[0]?.category).toBe('factual-consistency/currency-mismatch');
  });

  it('Ten Thousand Dollars ($10,000) — Informational ambiguous-currency (not Critical)', () => {
    const f = scan('Ten Thousand Dollars ($10,000)');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('info');
    expect(f[0]?.category).toBe('factual-consistency/ambiguous-currency');
  });

  it('Ten Thousand Pounds (£10,000) — no finding', () => {
    expect(scan('Ten Thousand Pounds (£10,000)', UK_PROFILE)).toHaveLength(0);
  });

  it('bare $ never produces a Critical (AC-2)', () => {
    const f = scan('Five Thousand Dollars ($5,000)');
    const criticals = f.filter(x => x.severity === 'error');
    expect(criticals).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// AC-4 — Structural identifiers excluded
// ---------------------------------------------------------------------------

describe('Amounts — AC-4 structural identifiers', () => {
  it('Clause Three (3) — no finding', () => {
    expect(scan('Clause Three (3)')).toHaveLength(0);
  });

  it('thirty (30) days — no finding', () => {
    expect(scan('thirty (30) days')).toHaveLength(0);
  });

  it('Schedule Two (2) — no finding', () => {
    expect(scan('Schedule Two (2)')).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Additional coverage
// ---------------------------------------------------------------------------

describe('Amounts — additional cases', () => {
  it('findings route to proof tab (category prefix)', () => {
    const f = scan('Ten Thousand Dollars (S$1,000)');
    expect(f[0]?.category).toMatch(/^factual-consistency\//);
  });

  it('S$10,000 vs SGD10,000 — both map to SGD, no currency mismatch', () => {
    expect(scan('Ten Thousand Dollars (SGD10,000)')).toHaveLength(0);
  });

  it('prose word before number words — matches from the number word start', () => {
    // "Pay Ten Thousand Dollars (S$1,000)" — "Pay" is not a number word;
    // the regex must start from "Ten", giving wordsValue=10000 vs figure=1000 → mismatch.
    const f = scan('Pay Ten Thousand Dollars (S$1,000)');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
  });

  it('two distinct amount pairs in one paragraph — both detected', () => {
    const f = scan('Ten Thousand Dollars (S$1,000) and Five Thousand Dollars (S$4,000)');
    // Both pairs are mismatches (10000≠1000 and 5000≠4000)
    expect(f.length).toBeGreaterThanOrEqual(2);
    expect(f.every(x => x.severity === 'error')).toBe(true);
  });
});
