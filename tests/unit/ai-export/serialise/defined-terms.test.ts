import { describe, it, expect } from 'vitest';
import { serialiseDefinedTerms } from '@features/ai-export/serialise/defined-terms.js';
import type {
  DefinitionsMap, DefinitionEntry, DefinitionRecord, UsageIndex, UsageRecord,
} from '@features/defined-terms/index.js';
import { makeCiter } from '@features/ai-export/cite.js';
import { makeNode, makeIndex, paraTextFrom } from '../_helpers.js';

function rec(p: Partial<DefinitionRecord> & { paragraphId: string }): DefinitionRecord {
  return {
    source: 'prose',
    verb: 'means',
    shape: 'A',
    isCrossRef: false,
    isCrossDoc: false,
    crossRefTarget: null,
    anchor: { paragraphId: p.paragraphId, startOffset: 0, endOffset: 4 },
    styling: { bold: false, italic: false },
    rawQuote: 'double',
    contextScope: null,
    ...p,
  };
}

function entry(key: string, canonical: string, definitions: DefinitionRecord[]): DefinitionEntry {
  return { key, canonical, definitions };
}

function usageOf(pairs: Array<[string, string[]]>): UsageIndex {
  const m: UsageIndex = new Map();
  for (const [key, surfaces] of pairs) {
    m.set(key, surfaces.map<UsageRecord>(s => ({
      surface: s, isLowercase: /^[a-z]/.test(s), precededByDeterminer: false,
      isSentenceInitial: false, anchor: { paragraphId: 'x', startOffset: 0, endOffset: 0 },
      inFootnote: false, isDefinitionOccurrence: false,
    })));
  }
  return m;
}

describe('serialiseDefinedTerms', () => {
  const index = makeIndex([
    makeNode({ key: 'clause:14.2', displayLabel: 'Clause 14.2', paraId: 'pTarget' }),
    makeNode({ key: 'clause:1.1', displayLabel: 'Clause 1.1', paraId: 'pDef' }),
  ]);
  const paraText = paraTextFrom({
    pDef: 'Completion Date means the date of completion.',
    pXref: 'Long Stop Date has the meaning given in Clause 14.2.',
    pTarget: 'means 31 December 2026',
    pMulti: 'Buyer means ACME.',
  });
  const cite = makeCiter(index, ['pDef', 'pXref', 'pTarget', 'pMulti']);

  it('emits the canonical term, definition body, citation and source', () => {
    const map: DefinitionsMap = new Map([
      ['completion date', entry('completion date', 'Completion Date', [rec({ paragraphId: 'pDef' })])],
    ]);
    const out = serialiseDefinedTerms(map, { usage: new Map(), index, cite, paraText });
    expect(out.glossary).toHaveLength(1);
    const t = out.glossary[0]!;
    expect(t.term).toBe('Completion Date');
    expect(t.definition).toContain('means the date of completion');
    expect(t.defined_at).toBe('Clause 1.1'); // pDef is the Clause 1.1 heading
    expect(t.source).toBe('prose');
    expect(t.defined_multiple_times).toBeUndefined();
  });

  it('resolves a cross-reference definition inline (F2)', () => {
    const map: DefinitionsMap = new Map([
      ['long stop date', entry('long stop date', 'Long Stop Date', [
        rec({ paragraphId: 'pXref', isCrossRef: true, crossRefTarget: 'Clause 14.2', verb: 'cross-reference' }),
      ])],
    ]);
    const out = serialiseDefinedTerms(map, { usage: new Map(), index, cite, paraText });
    expect(out.glossary[0]!.definition).toContain('[→ "means 31 December 2026"]');
  });

  it('marks "[→ target not found...]" when the referent is missing', () => {
    const map: DefinitionsMap = new Map([
      ['x', entry('x', 'X', [rec({ paragraphId: 'pXref', isCrossRef: true, crossRefTarget: 'Clause 99' })])],
    ]);
    const out = serialiseDefinedTerms(map, { usage: new Map(), index, cite, paraText });
    expect(out.glossary[0]!.definition).toContain('[→ target not found in this document]');
  });

  it('flags defined_multiple_times when >1 substantive record', () => {
    const map: DefinitionsMap = new Map([
      ['buyer', entry('buyer', 'Buyer', [rec({ paragraphId: 'pMulti' }), rec({ paragraphId: 'pDef' })])],
    ]);
    const out = serialiseDefinedTerms(map, { usage: new Map(), index, cite, paraText });
    expect(out.glossary[0]!.defined_multiple_times).toBe(true);
  });

  it('collects surface variants from the usage index (F5)', () => {
    const map: DefinitionsMap = new Map([
      ['completion date', entry('completion date', 'Completion Date', [rec({ paragraphId: 'pDef' })])],
    ]);
    const usage = usageOf([['completion date', ['Completion Date', 'completion date', 'Completion Date']]]);
    const out = serialiseDefinedTerms(map, { usage, index, cite, paraText });
    expect(out.glossary[0]!.variants).toEqual(['completion date']);
  });
});
