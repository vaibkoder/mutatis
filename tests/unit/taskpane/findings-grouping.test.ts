/**
 * Host-neutral findings-grouping helpers — the single source of
 * the category labels/ordering/split shared by the add-in's FindingsList and the
 * standalone web view.
 */
import { describe, it, expect } from 'vitest';
import type { Finding } from '@shared/findings/index.js';
import {
  categoryLabel,
  groupByCategory,
  splitFindingsByFeature,
  XREF_CATEGORY_ORDER,
  DEFINED_TERMS_CATEGORY_ORDER,
} from '../../../src/taskpane/components/findings-grouping.js';

function f(id: string, category: string): Finding {
  return { id, category, featureId: 'x', severity: 'info', title: id, location: { label: 'p' } } as Finding;
}

describe('findings-grouping', () => {
  it('categoryLabel returns the friendly label, or the raw category when unknown', () => {
    expect(categoryLabel('cross-reference/not-found')).toBe('Broken (not found)');
    expect(categoryLabel('made-up/category')).toBe('made-up/category');
  });

  it('groupByCategory buckets a flat list preserving order', () => {
    const grouped = groupByCategory([f('a', 'cat-1'), f('b', 'cat-2'), f('c', 'cat-1')]);
    expect([...grouped.keys()]).toEqual(['cat-1', 'cat-2']);
    expect(grouped.get('cat-1')!.map(x => x.id)).toEqual(['a', 'c']);
  });

  it('splitFindingsByFeature separates the four features and orders sub-groups', () => {
    const grouped = groupByCategory([
      f('s1', 'static-numbering/manually-typed'),
      f('x-static', 'cross-reference/static'),
      f('x-nf', 'cross-reference/not-found'),
      f('dt-cap', 'defined-terms/capitalisation'),
      f('dt-und', 'defined-terms/used-not-defined'),
      f('fc', 'factual-consistency/placeholder'),
    ]);
    const { xrefOrdered, staticNum, definedTermsOrdered, other } = splitFindingsByFeature(grouped);

    // Cross-reference sub-groups appear in XREF_CATEGORY_ORDER (not-found before static).
    expect(xrefOrdered.map(g => g.cat)).toEqual(
      XREF_CATEGORY_ORDER.filter(c => c === 'cross-reference/not-found' || c === 'cross-reference/static'),
    );
    // Defined-terms sub-groups in DEFINED_TERMS_CATEGORY_ORDER (used-not-defined before capitalisation).
    expect(definedTermsOrdered.map(g => g.cat)).toEqual(
      DEFINED_TERMS_CATEGORY_ORDER.filter(
        c => c === 'defined-terms/used-not-defined' || c === 'defined-terms/capitalisation',
      ),
    );
    expect(staticNum.map(x => x.id)).toEqual(['s1']);
    expect(other.map(g => g.cat)).toEqual(['factual-consistency/placeholder']);
  });

  it('drops empty sub-groups from the ordered features', () => {
    const grouped = groupByCategory([f('x', 'cross-reference/ambiguous')]);
    const { xrefOrdered } = splitFindingsByFeature(grouped);
    expect(xrefOrdered.map(g => g.cat)).toEqual(['cross-reference/ambiguous']);
  });
});
