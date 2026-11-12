import { describe, it, expect, beforeEach } from 'vitest';
import { runFactualConsistency, clearThenEmit } from '@features/factual-consistency/index.js';
import { allFindings, addFinding, clearFindings, findingsForTab } from '@shared/findings/index.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function makePara(id: string, text: string): ParaNode {
  return {
    id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); },
  };
}

describe('Factual Consistency driver (M0)', () => {
  beforeEach(() => clearFindings());

  it('returns an empty array on a plain paragraph (stubs)', () => {
    const paras = [makePara('p0', 'This is a normal clause.')];
    const findings = runFactualConsistency(paras, SG_PROFILE);
    expect(findings).toHaveLength(0);
  });

  it('returns an empty array on an empty paragraph list', () => {
    expect(runFactualConsistency([], SG_PROFILE)).toHaveLength(0);
  });

  it('clearThenEmit clears previous factual-consistency findings before emitting', () => {
    addFinding({
      id: 'factual-consistency/old-1',
      category: 'factual-consistency/placeholder',
      severity: 'info',
      title: 'Old finding',
      location: { label: 'p0' },
    });
    expect(allFindings.value).toHaveLength(1);

    clearThenEmit([makePara('p0', 'clean text')], SG_PROFILE);

    // stubs return nothing → store should be empty after clear
    expect(allFindings.value).toHaveLength(0);
  });

  it('clearThenEmit does not remove non-factual-consistency findings', () => {
    addFinding({
      id: 'defined-terms/unused-1',
      category: 'defined-terms/unused',
      severity: 'warning',
      title: 'Unused term',
      location: { label: 'p0' },
    });
    clearThenEmit([makePara('p0', 'clean text')], SG_PROFILE);
    expect(allFindings.value).toHaveLength(1);
    expect(allFindings.value[0]?.id).toBe('defined-terms/unused-1');
  });

  it('factual-consistency categories route to the proof tab', () => {
    addFinding({
      id: 'factual-consistency/placeholder/p0-0-3',
      category: 'factual-consistency/placeholder',
      severity: 'info',
      title: 'Placeholder',
      location: { label: 'p0' },
    });
    const proofFindings = findingsForTab('proof');
    expect(proofFindings).toHaveLength(1);
  });
});
