import { describe, it, expect } from 'vitest';
import { runDefinedTerms } from '@features/defined-terms/index.js';
import { buildDefinitionsMap } from '@features/defined-terms/definitions.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function body(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}
function footnote(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'FootnoteText', container: 'footnote',
    getRange: () => { throw new Error('not used'); } };
}

describe('footnote definition pass (T7)', () => {
  const doc = [
    body('body-para-0', 'A Material Adverse Effect entitles the Buyer to terminate.'),
    footnote('footnote-para-0', '"Material Adverse Effect" means any event materially adverse to the business.'),
  ];

  it('records a footnote definition with source = "footnote"', () => {
    const map = buildDefinitionsMap(
      doc.filter(p => p.container === 'body'),
      doc.filter(p => p.container === 'footnote'),
      SG_PROFILE,
    );
    const e = map.get('material adverse effect');
    expect(e).toBeDefined();
    expect(e?.definitions[0]?.source).toBe('footnote');
  });

  it('a footnote-defined term used in the body is neither used-but-not-defined nor defined-but-not-used', () => {
    const fs = runDefinedTerms(doc, SG_PROFILE);
    expect(fs.filter(f => f.category === 'defined-terms/used-not-defined')).toHaveLength(0);
    expect(fs.filter(f => f.category === 'defined-terms/defined-not-used')).toHaveLength(0);
  });
});
