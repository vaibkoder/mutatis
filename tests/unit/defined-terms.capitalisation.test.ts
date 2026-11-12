import { describe, it, expect } from 'vitest';
import { runDefinedTerms } from '@features/defined-terms/index.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function paras(...texts: string[]): ParaNode[] {
  return texts.map((text, i) => ({
    id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'body' as const,
    getRange: () => { throw new Error('not used'); },
  }));
}

function cap(texts: string[], mode?: 'smart' | 'strict') {
  const opts = mode ? { capModeOverride: mode } : undefined;
  return runDefinedTerms(paras(...texts), SG_PROFILE, opts)
    .filter(f => f.category === 'defined-terms/capitalisation');
}

describe('sub-check 4 — capitalisation (T6, T11)', () => {
  const propertyDoc = [
    '"Property" means the land and buildings.',
    'When the property is sold all property and assets transfer.',
  ];

  it('T6 — smart (default) flags only the determiner-led lowercase use', () => {
    expect(cap(propertyDoc)).toHaveLength(1);
    expect(cap(propertyDoc, 'smart')).toHaveLength(1);
  });

  it('T6 — strict flags every lowercase use', () => {
    expect(cap(propertyDoc, 'strict')).toHaveLength(2);
  });

  it('T11 — en-GB lowercase use yields a capitalisation finding, never a spelling finding', () => {
    const all = runDefinedTerms(paras(
      '"Authorised Signatory" means a person authorised to sign.',
      'Each authorised signatory shall countersign.',
    ), SG_PROFILE);
    const capHits = all.filter(f => f.category === 'defined-terms/capitalisation');
    expect(capHits).toHaveLength(1);
    // No spelling category is ever emitted by this feature.
    expect(all.some(f => f.category.includes('spell'))).toBe(false);
  });

  it('capitalisation findings are Informational', () => {
    expect(cap(propertyDoc)[0]?.severity).toBe('info');
  });

  it('capitalisation finding carries a sourceSnippet pointing to the definition paragraph', () => {
    const findings = cap(propertyDoc);
    expect(findings.length).toBeGreaterThan(0);
    const f = findings[0]!;
    expect(f.sourceSnippet).toBeDefined();
    expect(f.sourceSnippet?.label).toBe('Property');
    // The source text should be the definition paragraph, not the usage paragraph.
    expect(f.sourceSnippet?.text).toContain('"Property" means');
  });
});
