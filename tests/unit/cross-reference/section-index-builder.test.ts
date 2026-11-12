/**
 * Section Index builder — extended tests (spec §4, T2 acceptance).
 *
 * Tests compound sub-clause keys, static-numbered fallback, §4.4 normalisation,
 * roman/letter disambiguation, bookmarkNames default, has/resolve/ordered.
 */
import { describe, it, expect } from 'vitest';
import { buildSectionIndex } from '@shared/section-index/builder.js';
import { UK_PROFILE, SG_PROFILE } from '@shared/config/index.js';
import { para } from '../../fixtures/cross-reference/paras.js';

describe('section index builder — compound keys (spec §4.3)', () => {
  it('clause:4.2.1 — decimal already compound', () => {
    const paras = [
      para(0, 'Definitions', { listString: '1', listLevel: 0 }),
      para(1, 'Interpretation', { listString: '4.2.1', listLevel: 2 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:4.2.1')).toBe(true);
    expect(index.resolve('clause:4.2.1')).toHaveLength(1);
  });

  it('clause:5(b) — lettered sub-clause uses parent compound', () => {
    const paras = [
      para(0, 'Warranties', { listString: '5', listLevel: 0 }),
      para(1, 'The Seller warrants that…', { listString: '(a)', listLevel: 1 }),
      para(2, 'The Seller further warrants…', { listString: '(b)', listLevel: 1 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:5')).toBe(true);
    expect(index.has('clause:5(a)')).toBe(true);
    expect(index.has('clause:5(b)')).toBe(true);
    expect(index.resolve('clause:5(b)')[0]?.parentKey).toBe('clause:5');
  });

  it('clause:5(b)(i) — two levels of sub-clauses', () => {
    const paras = [
      para(0, 'Warranties', { listString: '5', listLevel: 0 }),
      para(1, 'The Seller warrants…', { listString: '(b)', listLevel: 1 }),
      para(2, 'In particular…', { listString: '(i)', listLevel: 2 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:5(b)(i)')).toBe(true);
    expect(index.resolve('clause:5(b)(i)')[0]?.parentKey).toBe('clause:5(b)');
  });

  it('schedule:2/para:3 — paragraph inside Schedule', () => {
    const paras = [
      para(0, 'Schedule 2', { styleName: 'Heading 1' }),
      para(1, 'The property…', { listString: '1', listLevel: 0 }),
      para(2, 'Land and buildings…', { listString: '3', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('schedule:2')).toBe(true);
    expect(index.has('schedule:2/para:3')).toBe(true);
    const node = index.resolve('schedule:2/para:3')[0];
    expect(node?.displayLabel).toBe('paragraph 3 of Schedule 2');
  });

  it('recital:B — lettered recital', () => {
    const paras = [
      para(0, 'Recitals'),
      para(1, '(A) The Seller is the legal owner…'),
      para(2, '(B) The Buyer wishes to purchase…'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('recital:A')).toBe(true);
    expect(index.has('recital:B')).toBe(true);
    expect(index.resolve('recital:B')[0]?.displayLabel).toBe('Recital (B)');
  });

  it('part:II — roman-numbered Part', () => {
    const paras = [
      para(0, 'Part II — Additional Terms', { styleName: 'Heading 1' }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('part:II')).toBe(true);
  });
});

describe('section index builder — static-numbered fallback (spec §4.3 step 2)', () => {
  it('indexes a static clause (no listString) with isStaticNumbered=true', () => {
    const paras = [
      para(0, '4.2 The Buyer shall pay the Purchase Price…'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:4.2')).toBe(true);
    expect(index.resolve('clause:4.2')[0]?.isStaticNumbered).toBe(true);
  });

  it('does not drop static units — no false Not Found', () => {
    // Static unit "3" is present; another paragraph references it.
    // The index should contain it so a scanner can resolve it.
    const paras = [
      para(0, '3. Payment Terms'),
      para(1, 'The amount is due as per clause 3 above.'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:3')).toBe(true);
    expect(index.resolve('clause:3')[0]?.isStaticNumbered).toBe(true);
  });

  it('does NOT create phantom clauses from numeric-leading prose (bare integer + space)', () => {
    // Regression: "5 days after…", "1976 Act", "30 June 2026" must NOT be indexed.
    const paras = [
      para(0, '5 days after the date of this Agreement the Buyer shall…'),
      para(1, 'the Companies Act 1976 applies to the following…'),
      para(2, '30 June 2026 is the Completion Date.'),
      para(3, '2 parties have agreed the following terms.'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:5')).toBe(false);
    expect(index.has('clause:1976')).toBe(false);
    expect(index.has('clause:30')).toBe(false);
    expect(index.has('clause:2')).toBe(false);
    expect(index.nodes).toHaveLength(0);
  });

  it('still indexes real static clauses with non-space delimiters ("5.", "5)")', () => {
    const paras = [
      para(0, '5. Payment Terms'),
      para(1, '6) Warranties'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:5')).toBe(true);
    expect(index.has('clause:6')).toBe(true);
  });

  it('phantom-prevention preserves A3: prose "5 days" does not defeat implied-parent ambiguity', () => {
    // 5.1/5.2 exist (list items); a prose paragraph starts "5 days…".
    // Without the fix, "5 days" would create a phantom clause:5 and break A3.
    const paras = [
      para(0, 'The Seller warrants…', { listString: '5.1', listLevel: 1 }),
      para(1, 'The Seller further warrants…', { listString: '5.2', listLevel: 1 }),
      para(2, '5 days after Completion the Buyer shall pay…'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.has('clause:5')).toBe(false);
    expect(index.has('clause:5.1')).toBe(true);
  });
});

describe('section index builder — roman/letter disambiguation (spec §4.4)', () => {
  it('marks (i) as ambiguousRomanLetter when first in sequence', () => {
    const paras = [
      para(0, 'X', { listString: '5', listLevel: 0 }),
      para(1, 'Y', { listString: '(i)', listLevel: 1 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const node = index.resolve('clause:5(i)')[0];
    expect(node?.ambiguousRomanLetter).toBe(true);
  });

  it('does NOT mark (i) ambiguous when preceded by (h)', () => {
    const paras = [
      para(0, 'X', { listString: '5', listLevel: 0 }),
      para(1, 'a', { listString: '(g)', listLevel: 1 }),
      para(2, 'b', { listString: '(h)', listLevel: 1 }),
      para(3, 'c', { listString: '(i)', listLevel: 1 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const node = index.resolve('clause:5(i)')[0];
    // (h) is a non-roman letter, so (i) is clearly a letter → not ambiguous
    expect(node?.ambiguousRomanLetter).toBeFalsy();
  });
});

describe('section index builder — config-driven displayLabel (spec §4.3 step 5)', () => {
  it('uses "Clause" from config for body clauses', () => {
    const paras = [para(0, 'Payment', { listString: '4', listLevel: 0 })];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.resolve('clause:4')[0]?.displayLabel).toBe('Clause 4');
  });

  it('uses "Clause" (not "Section") for SG profile too', () => {
    const paras = [para(0, 'Payment', { listString: '4', listLevel: 0 })];
    const index = buildSectionIndex(paras, SG_PROFILE);
    expect(index.resolve('clause:4')[0]?.displayLabel).toBe('Clause 4');
  });
});

describe('section index builder — interface (has/resolve/ordered/bookmarkNames)', () => {
  it('has() returns false for absent key', () => {
    const index = buildSectionIndex([], UK_PROFILE);
    expect(index.has('clause:99')).toBe(false);
  });

  it('ordered equals nodes array', () => {
    const paras = [
      para(0, 'A', { listString: '1', listLevel: 0 }),
      para(1, 'B', { listString: '2', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.ordered).toBe(index.nodes);
    expect(index.ordered).toHaveLength(2);
  });

  it('bookmarkNames defaults to empty array on every node', () => {
    const paras = [para(0, 'X', { listString: '1', listLevel: 0 })];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.resolve('clause:1')[0]?.bookmarkNames).toEqual([]);
  });

  it('Ambiguous: same key twice → resolve returns 2', () => {
    const paras = [
      para(0, 'Term A', { listString: '5', listLevel: 0 }),
      para(1, 'Term B', { listString: '5', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(index.resolve('clause:5')).toHaveLength(2);
  });
});
