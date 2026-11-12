import { describe, it, expect } from 'vitest';
import { runDefinedTerms } from '@features/defined-terms/index.js';
import { buildDefinitionsMap } from '@features/defined-terms/definitions.js';
import { SG_PROFILE } from '@shared/config/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

let seq = 0;
function body(text: string): ParaNode {
  const i = seq++;
  return { id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}
function defs(text: string, profile: JurisdictionProfile = SG_PROFILE) {
  return buildDefinitionsMap([body(text)], [], profile);
}
function count(fs: ReturnType<typeof runDefinedTerms>, suffix: string): number {
  return fs.filter(f => f.category === `defined-terms/${suffix}`).length;
}

// ---------------------------------------------------------------------------
// ADR-0013 false-positive trap suite
// ---------------------------------------------------------------------------

describe('ADR-0013 FP traps — must NOT create a definition', () => {
  it('"The Company means business." → no definition', () => {
    expect(defs('The Company means business.').size).toBe(0);
  });

  it('"This Clause means that …" → no definition', () => {
    expect(defs('This Clause means that the parties agree.').size).toBe(0);
  });

  it('ordinary capitalised noun + "means" mid-sentence → no definition', () => {
    expect(defs('Under the deal the Property means a great deal to them.').size).toBe(0);
  });

  it('the broadened forms ARE still recognised', () => {
    expect(defs('"Business Day": any day banks are open.').has('business day')).toBe(true);
    expect(defs('Affiliate means a holding company.').has('affiliate')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Config gating — recogniseUnquoted / recogniseColon (ADR-0013)
// ---------------------------------------------------------------------------

describe('config gating disables the broadened shapes', () => {
  const noUnquoted: JurisdictionProfile = {
    ...SG_PROFILE,
    definedTerms: { ...SG_PROFILE.definedTerms, recogniseUnquoted: false },
  };
  const noColon: JurisdictionProfile = {
    ...SG_PROFILE,
    definedTerms: { ...SG_PROFILE.definedTerms, recogniseColon: false },
  };

  it('recogniseUnquoted=false disables Shape C', () => {
    expect(defs('Affiliate means a holding company.', noUnquoted).has('affiliate')).toBe(false);
    // quoted shapes still work
    expect(defs('"Affiliate" means a holding company.', noUnquoted).has('affiliate')).toBe(true);
  });

  it('recogniseColon=false disables Shape A′', () => {
    expect(defs('"Business Day": any day banks are open.', noColon).has('business day')).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// spec §8 edge cases
// ---------------------------------------------------------------------------

describe('spec §8 edge cases', () => {
  it('plurals and possessives resolve to the singular key', () => {
    const fs = runDefinedTerms([
      body('"Shareholder" means a holder of shares.'),
      body("The Shareholders meet and each Shareholder's vote counts."),
    ], SG_PROFILE);
    expect(count(fs, 'used-not-defined')).toBe(0);
    expect(count(fs, 'defined-not-used')).toBe(0);
  });

  it('a term used inside another definition counts as usage (not defined-but-not-used)', () => {
    const fs = runDefinedTerms([
      body('"Group" means the holding structure.'),
      body('"Affiliate" means a member of the Group.'),
    ], SG_PROFILE);
    const dnu = fs.filter(f => f.category === 'defined-terms/defined-not-used').map(f => f.title);
    expect(dnu.join(' ')).not.toContain('Group');
  });

  it('plural forms of allowlisted party-roles are suppressed (allowlist is plural-aware)', () => {
    const fs = runDefinedTerms([
      body('"Completion Date" means a date; the Completion Date applies.'),
      body('The Buyers and the Sellers and the Parties shall each perform.'),
    ], SG_PROFILE);
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('all-caps acronym introduced parenthetically is defined; allowlisted acronyms are quiet', () => {
    const fs = runDefinedTerms([
      body('The sale and purchase agreement (the "SPA") governs the deal.'),
      body('The SPA and any VAT or GST shall apply.'),
    ], SG_PROFILE);
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('smart-quote folding — curly-quoted definition recognised and used', () => {
    const fs = runDefinedTerms([
      body('“Closing Date” means the date of completion.'),
      body('On the Closing Date the parties act.'),
    ], SG_PROFILE);
    expect(fs).toHaveLength(0);
  });

  it('a self-referential heading is not flagged used-but-not-defined', () => {
    const fs = runDefinedTerms([
      body('"Completion Date" means a date; the Completion Date applies.'),
      body('Definitions and Interpretation'),
      body('The obligations under this Agreement and this Deed bind the Parties.'),
    ], SG_PROFILE);
    expect(count(fs, 'used-not-defined')).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Performance smoke test — large synthetic document
// ---------------------------------------------------------------------------

describe('performance — >100-paragraph document', () => {
  it('runs over 120 paragraphs and returns deterministic findings', () => {
    seq = 0;
    const ps: ParaNode[] = [];
    ps.push(body('"Master Term" means a defined concept; the Master Term recurs throughout.'));
    for (let i = 0; i < 120; i++) {
      ps.push(body(`In Clause ${i} the Buyer shall perform its obligations under this Agreement.`));
    }
    const a = runDefinedTerms(ps, SG_PROFILE);
    const b = runDefinedTerms(ps, SG_PROFILE);
    expect(a.length).toBe(b.length);
    // No spurious used-but-not-defined from the boilerplate (Clause/Buyer/Agreement allowlisted/self-ref).
    expect(count(a, 'used-not-defined')).toBe(0);
  });
});
