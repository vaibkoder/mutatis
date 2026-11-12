import { describe, it, expect } from 'vitest';
import { makeCiter } from '@features/ai-export/cite.js';
import { makeNode, makeIndex } from './_helpers.js';

describe('makeCiter', () => {
  const nodes = [
    makeNode({ key: 'clause:1', displayLabel: 'Clause 1', paraId: 'p1', level: 0 }),
    makeNode({ key: 'clause:1.1', displayLabel: 'Clause 1.1', paraId: 'p3', level: 1 }),
    makeNode({ key: 'clause:2', displayLabel: 'Clause 2', paraId: 'p6', level: 0 }),
  ];
  const index = makeIndex(nodes);
  const order = ['p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7'];

  it('returns the displayLabel for an exact heading paragraph', () => {
    const cite = makeCiter(index, order);
    expect(cite('p1')).toBe('Clause 1');
    expect(cite('p6')).toBe('Clause 2');
  });

  it('resolves a mid-clause paragraph to its enclosing heading', () => {
    const cite = makeCiter(index, order);
    expect(cite('p2')).toBe('Clause 1');   // between Clause 1 (p1) and Clause 1.1 (p3)
    expect(cite('p4')).toBe('Clause 1.1'); // after Clause 1.1 (p3), before Clause 2 (p6)
    expect(cite('p7')).toBe('Clause 2');   // after the last heading
  });

  it('returns "" with no index (cross-reference not run)', () => {
    const cite = makeCiter(null, order);
    expect(cite('p2')).toBe('');
  });

  it('returns "" for a paragraph before any heading or unknown to the order', () => {
    const cite = makeCiter(index, order);
    expect(cite('unknown')).toBe('');
  });
});
