import { describe, it, expect, beforeEach } from 'vitest';
import {
  addFinding,
  addFindings,
  removeFinding,
  removeByPrefix,
  clearFindings,
  findingsForTab,
  findingsByCategory,
  tabCounts,
  findingToTab,
} from '@shared/findings/index.js';
import type { Finding } from '@shared/findings/index.js';

function makeFinding(id: string, category: string, severity: Finding['severity'] = 'warning'): Finding {
  return {
    id,
    category,
    severity,
    title: `Finding ${id}`,
    location: {
      label: `Clause ${id}`,
    },
  };
}

describe('findings store', () => {
  beforeEach(() => clearFindings());

  it('adds and retrieves findings by tab', () => {
    addFinding(makeFinding('1', 'defined-terms/used-not-defined'));
    addFinding(makeFinding('2', 'cross-reference/not-found'));
    const proof = findingsForTab('proof');
    expect(proof).toHaveLength(2);
  });

  it('maps categories to the correct tab', () => {
    expect(findingToTab(makeFinding('a', 'repair/metadata'))).toBe('repair');
    expect(findingToTab(makeFinding('b', 'numbering/scheme'))).toBe('numbering');
    expect(findingToTab(makeFinding('c', 'compare/insertion'))).toBe('compare');
    expect(findingToTab(makeFinding('d', 'interactive-lookup/preview'))).toBe('review');
  });

  it('replaces a finding with the same id', () => {
    addFinding(makeFinding('1', 'defined-terms/x'));
    addFinding(makeFinding('1', 'defined-terms/y'));
    expect(findingsForTab('proof')).toHaveLength(1);
  });

  it('sorts findings error-first within a tab', () => {
    addFindings([
      makeFinding('a', 'defined-terms/x', 'info'),
      makeFinding('b', 'defined-terms/x', 'error'),
      makeFinding('c', 'defined-terms/x', 'warning'),
    ]);
    const sorted = findingsForTab('proof');
    expect(sorted.map(f => f.severity)).toEqual(['error', 'warning', 'info']);
  });

  it('groups findings by category', () => {
    addFindings([
      makeFinding('a', 'defined-terms/used-not-defined'),
      makeFinding('b', 'defined-terms/used-not-defined'),
      makeFinding('c', 'cross-reference/not-found'),
    ]);
    const grouped = findingsByCategory('proof');
    expect(grouped.get('defined-terms/used-not-defined')).toHaveLength(2);
    expect(grouped.get('cross-reference/not-found')).toHaveLength(1);
  });

  it('updates tab counts reactively', () => {
    addFinding(makeFinding('1', 'repair/metadata'));
    addFinding(makeFinding('2', 'repair/comments'));
    expect(tabCounts.value.repair).toBe(2);
  });

  it('removes findings by id and by prefix', () => {
    addFindings([
      makeFinding('xref-1', 'cross-reference/not-found'),
      makeFinding('xref-2', 'cross-reference/not-found'),
      makeFinding('dt-1', 'defined-terms/x'),
    ]);
    removeFinding('xref-1');
    expect(findingsForTab('proof')).toHaveLength(2);
    removeByPrefix('xref-');
    expect(findingsForTab('proof')).toHaveLength(1);
  });

  it('clears all findings', () => {
    addFinding(makeFinding('1', 'repair/x'));
    clearFindings();
    expect(findingsForTab('repair')).toHaveLength(0);
  });
});
