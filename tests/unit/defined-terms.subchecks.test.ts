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

function run(...texts: string[]) {
  return runDefinedTerms(paras(...texts), SG_PROFILE);
}

function byCategory(fs: ReturnType<typeof run>, suffix: string) {
  return fs.filter(f => f.category === `defined-terms/${suffix}`);
}

describe('sub-check 1 — used-but-not-defined (T2, T9, T14)', () => {
  it('T2 — flags an undefined capitalised term', () => {
    const fs = run(
      '"Completion Date" means a date.', // gives a definition so map is non-empty
      'The Seller shall deliver the Disclosure Letter to the Buyer.',
    );
    const c = byCategory(fs, 'used-not-defined');
    expect(c).toHaveLength(1);
    expect(c[0]?.title).toContain('Disclosure Letter');
    expect(c[0]?.severity).toBe('error');
  });

  it('T9 — statutory citation and party-roles are suppressed', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Each Director must comply with the Companies Act 2006 and the Act.',
    );
    expect(byCategory(fs, 'used-not-defined')).toHaveLength(0);
  });

  it('T14 — a term capitalised only sentence-initially is not flagged', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Property of the Company. The property shall vest.',
    );
    const undefHits = byCategory(fs, 'used-not-defined').map(f => f.title);
    expect(undefHits.join(' ')).not.toContain('Property');
  });

  it('does not flag a term that IS defined (incl. plural use)', () => {
    const fs = run(
      '"Shareholder" means a holder of shares.',
      'The Shareholders shall meet annually and each Shareholder may vote.',
    );
    expect(byCategory(fs, 'used-not-defined')).toHaveLength(0);
  });
});

describe('sub-check 2 — defined-but-not-used (T3)', () => {
  it('T3 — a definition never referenced again is flagged Informational', () => {
    const fs = run('"Escrow Agent" means Bank X.');
    const c = byCategory(fs, 'defined-not-used');
    expect(c).toHaveLength(1);
    expect(c[0]?.severity).toBe('info');
    expect(c[0]?.title).toContain('Escrow Agent');
  });

  it('a defined term that is used is not flagged', () => {
    const fs = run(
      '"Escrow Agent" means Bank X.',
      'The Escrow Agent shall hold the funds.',
    );
    expect(byCategory(fs, 'defined-not-used')).toHaveLength(0);
  });
});

describe('sub-check 3 — defined-multiple-times (T4, T5)', () => {
  it('T4 — two substantive definitions of one term → Critical', () => {
    const fs = run(
      '"Group" means the Company and its subsidiaries.',
      '"Group" means the Company and each subsidiary undertaking.',
    );
    const c = byCategory(fs, 'defined-multiple');
    expect(c).toHaveLength(1);
    expect(c[0]?.severity).toBe('error');
    expect(c[0]?.title).toContain('Group');
  });

  it('T5 — cross-reference + one substantive definition is NOT a defect', () => {
    const fs = run(
      '"Group" has the meaning given in Clause 14.2.',
      'In this Agreement, "Group" means the Company and its subsidiaries.',
      'The Group shall consolidate.',
    );
    expect(byCategory(fs, 'defined-multiple')).toHaveLength(0);
  });
});

describe('sub-check 1 — proper noun + boilerplate exclusions', () => {
  // Helper: all-in-one para with one real definition + a proper noun mid-sentence
  function runWith(noun: string) {
    return run(
      '"Completion Date" means a date.',
      `The parties submit to the jurisdiction of ${noun} in all matters.`,
    );
  }

  it('does not flag country names (Singapore)', () => {
    expect(byCategory(runWith('Singapore'), 'used-not-defined')).toHaveLength(0);
  });

  it('does not flag multi-word regions (Southeast Asia)', () => {
    expect(byCategory(runWith('Southeast Asia'), 'used-not-defined')).toHaveLength(0);
  });

  it('does not flag month names (January)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Completion shall occur on 1 January in each calendar year.',
    );
    expect(byCategory(fs, 'used-not-defined').map(f => f.title).join(' ')).not.toContain('January');
  });

  it('does not flag day names (Monday)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Payments are due on the Monday immediately following the Completion Date.',
    );
    expect(byCategory(fs, 'used-not-defined').map(f => f.title).join(' ')).not.toContain('Monday');
  });

  it('does not flag company suffixes (LLC) when appearing standalone', () => {
    // "LLC" alone mid-sentence (e.g. after a comma in a list context) must not be flagged.
    const fs = run(
      '"Completion Date" means a date.',
      'The entity type is LLC or LLP, as elected by the counterparty.',
    );
    expect(byCategory(fs, 'used-not-defined').map(f => f.title).join(' ')).not.toContain('LLC');
  });

  it('does not flag legal boilerplate all-caps (THEREFORE)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'NOW, THEREFORE, in consideration of the foregoing, the parties agree as follows.',
    );
    expect(byCategory(fs, 'used-not-defined').map(f => f.title).join(' ')).not.toContain('THEREFORE');
  });

  it('does not flag nationality adjectives (British)', () => {
    expect(byCategory(runWith('British'), 'used-not-defined')).toHaveLength(0);
  });

  it('does not flag United Kingdom', () => {
    expect(byCategory(runWith('United Kingdom'), 'used-not-defined')).toHaveLength(0);
  });
});

describe('sub-check 1 — all-caps acronym suppression', () => {
  it('does not flag an undefined standalone all-caps acronym (LIBOR)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Interest is calculated by reference to LIBOR.',
    );
    expect(byCategory(fs, 'used-not-defined').map(f => f.title).join(' ')).not.toContain('LIBOR');
  });

  it('does not flag a compound containing an undefined all-caps word (Commencement of FEED)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'The parties agree that the Commencement of FEED shall occur by the Completion Date.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).not.toContain('FEED');
    expect(titles).not.toContain('Commencement of FEED');
  });

  it('does NOT suppress a compound when its all-caps component IS defined', () => {
    // "FEED" is defined → resolveLookup returns non-null → suppression does not apply.
    // "Commencement of FEED" is used mid-sentence but not itself defined → flagged.
    const fs = run(
      '"FEED" means Front-End Engineering and Design.',
      '"Completion Date" means a date.',
      'The parties agree that the Commencement of FEED shall occur by the Completion Date.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).toContain('Commencement of FEED');
  });

  it('a defined all-caps term used elsewhere is not flagged', () => {
    const fs = run(
      '"EBITDA" means earnings before interest, taxes, depreciation and amortisation.',
      'The EBITDA of the Group shall be calculated quarterly.',
    );
    expect(byCategory(fs, 'used-not-defined')).toHaveLength(0);
  });
});

describe('sub-check 1 — occurrence counting + all-caps recurrence', () => {
  it('records every occurrence of an undefined term (count + anchors)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'The Disclosure Letter is delivered. The Buyer reviews the Disclosure Letter and signs the Disclosure Letter.',
    );
    const c = byCategory(fs, 'used-not-defined').filter(f => f.title.includes('Disclosure Letter'));
    expect(c).toHaveLength(1);
    expect(c[0]?.occurrences?.length).toBe(3);
  });

  it('does not attach occurrences[] for a single-occurrence term', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'The Seller delivers the Disclosure Letter to the Buyer.',
    );
    const c = byCategory(fs, 'used-not-defined').filter(f => f.title.includes('Disclosure Letter'));
    expect(c).toHaveLength(1);
    expect(c[0]?.occurrences).toBeUndefined();
  });

  it('SURFACES an undefined all-caps compound that recurs (≥2×): Commencement of FEED', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'The Commencement of FEED is a milestone. Payment follows the Commencement of FEED.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).toContain('Commencement of FEED');
  });

  it('SURFACES a recurring standalone unknown acronym (LIBOR ×3)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Interest references LIBOR. The LIBOR screen rate applies. A LIBOR fallback exists.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).toContain('LIBOR');
  });

  it('still SUPPRESSES allowlisted boilerplate all-caps even when recurring (THEREFORE)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'Now THEREFORE the parties agree. And THEREFORE they covenant. THEREFORE it is settled.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).not.toContain('THEREFORE');
  });

  it('still SUPPRESSES a single undefined all-caps occurrence (FEED once)', () => {
    const fs = run(
      '"Completion Date" means a date.',
      'The Commencement of FEED shall occur by the Completion Date.',
    );
    const titles = byCategory(fs, 'used-not-defined').map(f => f.title).join(' ');
    expect(titles).not.toContain('FEED');
  });
});
