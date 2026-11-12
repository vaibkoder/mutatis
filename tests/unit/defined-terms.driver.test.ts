import { describe, it, expect } from 'vitest';
import {
  analyseDefinedTerms, runDefinedTerms,
} from '@features/defined-terms/index.js';
import { findingToTab } from '@shared/findings/index.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function para(id: string, text: string, container: ParaNode['container'] = 'body'): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container,
    getRange: () => { throw new Error('not used'); } };
}

describe('defined-terms driver — skeleton orchestration', () => {
  it('emits zero findings on an empty document', () => {
    expect(runDefinedTerms([], SG_PROFILE)).toEqual([]);
  });

  it('reports "No definitions detected" when no definitions are present', () => {
    const a = analyseDefinedTerms([para('p0', 'The Buyer shall pay the Seller.')], SG_PROFILE);
    expect(a.noDefinitions).toBe(true);
    expect(a.definitionCount).toBe(0);
    expect(a.findings).toEqual([]);
  });

  it('honours a capMode override', () => {
    const a = analyseDefinedTerms([para('p0', 'text')], SG_PROFILE, { capModeOverride: 'strict' });
    expect(a.capMode).toBe('strict');
  });

  it('every emitted defined-terms finding routes to the proof tab', () => {
    // Skeleton emits none yet; assert the routing contract on the category prefix.
    const sample = { category: 'defined-terms/used-not-defined' } as { category: string };
    expect(findingToTab(sample as never)).toBe('proof');
  });
});
