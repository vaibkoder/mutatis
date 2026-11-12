import { describe, it, expect } from 'vitest';
import { buildBriefingPack, type ExportContext, type BriefingData } from '@features/ai-export/index.js';
import type { DefinitionsMap } from '@features/defined-terms/index.js';
import type { DateRecord, PlaceholderHit } from '@features/factual-consistency/index.js';
import type { RoleRecord } from '@features/party-role/index.js';
import { makeNode, makeIndex, paraTextFrom } from '../_helpers.js';

const anchor = (paragraphId: string) => ({ paragraphId, startOffset: 0, endOffset: 1 });

const index = makeIndex([
  makeNode({ key: 'clause:1', displayLabel: 'Clause 1', paraId: 'p1', level: 0 }),
  makeNode({ key: 'clause:1.1', displayLabel: 'Clause 1.1', paraId: 'p2', level: 1, parentKey: 'clause:1' }),
]);

const ctx: ExportContext = {
  generated: '2026-06-09T14:32:00Z',
  document: 'sale-agreement.docx',
  jurisdiction: 'SG',
  index,
  orderedParaIds: ['p1', 'p2', 'pDate', 'pPh'],
  paraText: paraTextFrom({ p1: 'Definitions', p2: 'Interpretation', pDef: 'Buyer means ACME.' }),
};

const map: DefinitionsMap = new Map([
  ['buyer', { key: 'buyer', canonical: 'Buyer', definitions: [{
    source: 'prose', verb: 'means', shape: 'A', isCrossRef: false, isCrossDoc: false,
    crossRefTarget: null, anchor: anchor('pDef'), styling: { bold: false, italic: false },
    rawQuote: 'double', contextScope: null,
  }] }],
]);

const dates: DateRecord[] = [{
  anchor: anchor('pDate'), rawText: '6 June 2026', parsed: new Date(2026, 5, 6),
  format: 'dd-month-yyyy', centuryAmbiguous: false, isTemplate: false,
  roleHint: 'execution', dayOfWeek: 6, flaggedWeekend: true,
}];
const placeholders: PlaceholderHit[] = [{ anchor: anchor('pPh'), rawText: '[TBC]', category: 'keyword', suppressed: false, confidence: 'high' }];
const roles: RoleRecord[] = [{ key: 'purchaser', label: 'Purchaser', introduced: true, introSites: [anchor('p1')], uses: [anchor('a')] }];

const data: BriefingData = {
  definedTerms: { map, usage: new Map() },
  sectionOutline: index,
  dealFacts: { dates, amounts: [], placeholders },
  castList: { roles, anchors: [] },
  undefinedTerms: [{ raw: 'Completion', clause: 'Clause 5' }],
};

describe('buildBriefingPack', () => {
  const env = buildBriefingPack(data, ctx);

  it('checks_run matches the contributed sections (and excludes Compare)', () => {
    expect(env.checks_run).toEqual(['defined-terms', 'cross-reference', 'factual-consistency', 'party-role']);
    expect(env.blind_spots).toContain('Compare was not run; change set absent.');
  });

  it('assembles a mechanical preamble summary from the available data', () => {
    const s = env.preamble!.summary;
    expect(s).toContain('Singapore agreement');
    expect(s).toContain('between Purchaser');
    expect(s).toContain('dated 6 June 2026');
    expect(s).toContain('1 defined term');
    expect(s).toContain('2 clauses');
  });

  it('open_items includes both a placeholder and an undefined-term', () => {
    const types = env.preamble!.open_items.map(o => o.type);
    expect(types).toContain('placeholder');
    expect(types).toContain('undefined-term');
    expect(env.preamble!.summary).toContain('2 open items');
  });

  it('embeds the contributed Tier-1 sections (no change_set)', () => {
    expect(env.defined_terms).toBeDefined();
    expect(env.section_outline).toBeDefined();
    expect(env.deal_facts).toBeDefined();
    expect(env.cast_list).toBeDefined();
    expect(env.change_set).toBeUndefined();
  });
});
