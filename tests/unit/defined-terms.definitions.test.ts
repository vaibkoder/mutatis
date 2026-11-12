import { describe, it, expect } from 'vitest';
import { buildDefinitionsMap } from '@features/defined-terms/definitions.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';
import type { DefinitionsMap } from '@features/defined-terms/index.js';

function para(text: string, id = 'body-para-0'): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}

function defs(text: string, profile = SG_PROFILE): DefinitionsMap {
  return buildDefinitionsMap([para(text)], [], profile);
}

describe('prose definition pass — Shape A (quoted + verb)', () => {
  it('T1 — recognises a bold-double-quote definition (no double-quote FP)', () => {
    const m = defs('1.1 "Completion Date" means the date on which Completion occurs.');
    const e = m.get('completion date');
    expect(e).toBeDefined();
    expect(e?.canonical).toBe('Completion Date');
    expect(e?.definitions[0]?.shape).toBe('A');
    expect(e?.definitions[0]?.verb).toBe('means');
    expect(e?.definitions[0]?.rawQuote).toBe('double');
  });

  it('recognises single and curly quote delimiters', () => {
    expect(defs("'Escrow Agent' means Bank X.").has('escrow agent')).toBe(true);
    expect(defs('“Material Adverse Effect” means any event.').has('material adverse effect')).toBe(true);
  });

  it('"means and includes" is matched whole over "means"', () => {
    const e = defs('"Assets" means and includes all property.').get('assets');
    expect(e?.definitions[0]?.verb).toBe('means and includes');
  });
});

describe('prose definition pass — cross-reference vs cross-document', () => {
  it('T5 — "has the meaning given in Clause X" → cross-reference (same doc)', () => {
    const e = defs('"Group" has the meaning given in Clause 14.2.').get('group');
    const d = e?.definitions[0];
    expect(d?.isCrossRef).toBe(true);
    expect(d?.isCrossDoc).toBe(false);
    expect(d?.crossRefTarget).toBe('Clause 14.2');
  });

  it('T10 — "has the meaning given in the Facility Agreement" → cross-document', () => {
    const e = defs('"Permitted Encumbrances" has the meaning given in the Facility Agreement.')
      .get('permitted encumbrances');
    const d = e?.definitions[0];
    expect(d?.isCrossDoc).toBe(true);
    expect(d?.isCrossRef).toBe(false);
    expect(d?.crossRefTarget).toBe('Facility Agreement');
  });
});

describe('prose definition pass — Shape A′ (colon) and Shape B (parenthetical)', () => {
  it('ADR-0013 — colon-as-verb definition', () => {
    const e = defs('"Business Day": any day on which banks in Singapore are open.').get('business day');
    expect(e?.definitions[0]?.shape).toBe('A-prime');
    expect(e?.definitions[0]?.verb).toBeNull();
  });

  it('inline parenthetical — (the "Disclosing Party") / (each a "Shareholder")', () => {
    expect(defs('the disclosing party (the "Disclosing Party") shall.').get('disclosing party')
      ?.definitions[0]?.shape).toBe('B');
    expect(defs('the holders (each a "Shareholder") agree.').has('shareholder')).toBe(true);
  });
});

describe('prose definition pass — Shape C (unquoted, verb-anchored, ADR-0013)', () => {
  it('recognises an unquoted Title-Case headword + "means"', () => {
    const e = defs('Affiliate means any holding company of the Buyer.').get('affiliate');
    expect(e?.definitions[0]?.shape).toBe('C');
    expect(e?.definitions[0]?.verb).toBe('means');
  });

  it('recognises an ALL-CAPS headword', () => {
    expect(defs('AFFILIATE means any holding company.').has('affiliate')).toBe(true);
  });

  it('recognises an expanded "shall have the meaning given to it in" verb (either position)', () => {
    const e = defs('In this Agreement, Indemnity shall have the meaning given to it in Clause 3.')
      .get('indemnity');
    expect(e).toBeDefined();
    expect(e?.definitions[0]?.isCrossRef).toBe(true);
    expect(e?.definitions[0]?.crossRefTarget).toBe('Clause 3');
  });

  // --- The ADR-0013 false-positive traps ----------------------------------
  it('FP trap — "The Company means business." produces NO definition', () => {
    expect(defs('The Company means business.').size).toBe(0);
  });

  it('FP trap — "This Clause means that …" produces NO definition', () => {
    expect(defs('This Clause means that the parties agree.').size).toBe(0);
  });

  it('FP trap — an ordinary capitalised noun + "means" mid-sentence is not a definition', () => {
    expect(defs('Under the agreement the Property means a great deal to them.').size).toBe(0);
  });
});

describe('prose definition pass — works for both profiles', () => {
  it('UK profile recognises the same shapes', () => {
    const m = buildDefinitionsMap([para('"Long-Stop Date" means 31 December 2026.')], [], UK_PROFILE);
    expect(m.has('long-stop date')).toBe(true);
  });
});
