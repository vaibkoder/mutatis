import { describe, it, expect } from 'vitest';
import { buildSectionIndex, classifyKey, resolveKey } from '@shared/section-index/index.js';
import { UK_PROFILE } from '@shared/config/index.js';
import { para } from '../fixtures/paragraphs.js';

describe('section index builder', () => {
  it('builds clause keys from listString auto-numbers', () => {
    const paras = [
      para(0, 'Definitions and Interpretation', { listString: '1', listLevel: 0 }),
      para(1, 'In this Agreement…', { listString: '1.1', listLevel: 1 }),
      para(2, 'Defined terms have the meanings…', { listString: '1.2', listLevel: 1 }),
      para(3, 'Payment', { listString: '2', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);

    expect(resolveKey(index, 'clause:1')).toHaveLength(1);
    expect(resolveKey(index, 'clause:1.1')).toHaveLength(1);
    expect(resolveKey(index, 'clause:2')).toHaveLength(1);
    expect(classifyKey(index, 'clause:1')).toBe('ok');
    expect(classifyKey(index, 'clause:99')).toBe('not-found');
  });

  it('sets parent keys for nested clauses', () => {
    const paras = [
      para(0, 'Payment', { listString: '2', listLevel: 0 }),
      para(1, 'The Buyer shall pay…', { listString: '2.1', listLevel: 1 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    const child = resolveKey(index, 'clause:2.1')[0];
    expect(child?.parentKey).toBe('clause:2');
  });

  it('opens a Schedule region and keys its paragraphs as paragraph-of-Schedule', () => {
    const paras = [
      para(0, 'Schedule 1', { styleName: 'Heading 1' }),
      para(1, 'The Property', { listString: '1', listLevel: 0 }),
      para(2, 'The land known as…', { listString: '2', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);

    expect(resolveKey(index, 'schedule:1')).toHaveLength(1);
    const p = resolveKey(index, 'schedule:1/para:1')[0];
    expect(p).toBeDefined();
    expect(p?.displayLabel).toBe('paragraph 1 of Schedule 1');
  });

  it('keys recitals as uppercase letters', () => {
    const paras = [
      para(0, 'Recitals'),
      para(1, '(A) The Seller is the legal owner…'),
      para(2, '(B) The Buyer wishes to purchase…'),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(resolveKey(index, 'recital:A')).toHaveLength(1);
    expect(resolveKey(index, 'recital:B')).toHaveLength(1);
    expect(resolveKey(index, 'recital:A')[0]?.displayLabel).toBe('Recital (A)');
  });

  it('drives Ambiguous classification when a key appears twice', () => {
    const paras = [
      para(0, 'Term', { listString: '5', listLevel: 0 }),
      para(1, 'Duplicate term', { listString: '5', listLevel: 0 }),
    ];
    const index = buildSectionIndex(paras, UK_PROFILE);
    expect(resolveKey(index, 'clause:5')).toHaveLength(2);
    expect(classifyKey(index, 'clause:5')).toBe('ambiguous');
  });
});
