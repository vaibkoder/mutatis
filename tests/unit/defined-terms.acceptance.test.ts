import { describe, it, expect } from 'vitest';
import { runDefinedTerms, analyseDefinedTerms } from '@features/defined-terms/index.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

let seq = 0;
function body(text: string): ParaNode {
  const i = seq++;
  return { id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}
function footnote(text: string): ParaNode {
  const i = seq++;
  return { id: `footnote-para-${i}`, index: i, text, styleName: 'FootnoteText', container: 'footnote',
    getRange: () => { throw new Error('not used'); } };
}
function cell(text: string, tableId: string, row: number, col: number): ParaNode {
  const i = seq++;
  return { id: `body-para-${i}`, index: i, text, styleName: 'Normal', container: 'tableCell',
    tableInfo: { tableId, row, col }, getRange: () => { throw new Error('not used'); } };
}
function count(fs: ReturnType<typeof runDefinedTerms>, suffix: string): number {
  return fs.filter(f => f.category === `defined-terms/${suffix}`).length;
}

const PROFILES: ReadonlyArray<[string, JurisdictionProfile]> = [['SG', SG_PROFILE], ['UK', UK_PROFILE]];

describe.each(PROFILES)('Defined-Terms acceptance T1–T15 — %s profile', (_name, profile) => {
  const run = (...p: ParaNode[]) => runDefinedTerms(p, profile);

  it('T1 — used correctly → no findings', () => {
    const fs = run(
      body('1.1 "Completion Date" means the date on which Completion occurs.'),
      body('On the Completion Date the Buyer shall pay.'),
    );
    expect(fs).toHaveLength(0);
  });

  it('T2 — used-but-not-defined → 1 Critical', () => {
    const fs = run(
      body('"Completion Date" means a date; the Completion Date is fixed.'),
      body('The Seller shall deliver the Disclosure Letter to the Buyer.'),
    );
    expect(count(fs, 'used-not-defined')).toBe(1);
    expect(count(fs, 'defined-not-used')).toBe(0);
    expect(fs.find(f => f.category === 'defined-terms/used-not-defined')?.title).toContain('Disclosure Letter');
  });

  it('T3 — defined-but-not-used → 1 Informational (definition-body proper noun not flagged)', () => {
    const fs = run(body('"Escrow Agent" means Bank X.'));
    expect(count(fs, 'defined-not-used')).toBe(1);
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('T4 — defined-multiple-times (table + prose) → 1 Critical', () => {
    seq = 0;
    const fs = run(
      cell('"Group"', 't0', 0, 0),
      cell('the Company and its subsidiaries', 't0', 0, 1),
      body('"Group" means the Company and each subsidiary undertaking.'),
      body('The Group shall consolidate its accounts.'),
    );
    expect(count(fs, 'defined-multiple')).toBe(1);
  });

  it('T5 — cross-ref + substance → no multiple-definition; banner counts 1 cross-ref', () => {
    const ps = [
      body('"Group" has the meaning given in Clause 14.2.'),
      body('In this Agreement, "Group" means the Company and its subsidiaries.'),
      body('The Group shall consolidate.'),
    ];
    const a = analyseDefinedTerms(ps, profile);
    expect(count(a.findings, 'defined-multiple')).toBe(0);
    expect(a.crossRefCount).toBe(1);
  });

  it('T6 — capitalisation smart=1 / strict=2', () => {
    const ps = [
      body('"Property" means the land and buildings.'),
      body('When the property is sold all property and assets transfer.'),
    ];
    expect(count(runDefinedTerms(ps, profile), 'capitalisation')).toBe(1);
    expect(count(runDefinedTerms(ps, profile, { capModeOverride: 'strict' }), 'capitalisation')).toBe(2);
  });

  it('T7 — footnote definition recognised; no used/defined-not-used noise', () => {
    seq = 0;
    const fs = run(
      body('A Material Adverse Effect entitles the Buyer to terminate.'),
      footnote('"Material Adverse Effect" means any materially adverse event.'),
    );
    expect(count(fs, 'used-not-defined')).toBe(0);
    expect(count(fs, 'defined-not-used')).toBe(0);
  });

  it('T8 — Schedule table with leading numbering column → no findings', () => {
    seq = 0;
    const fs = run(
      cell('1.', 't0', 0, 0), cell('"Long-Stop Date"', 't0', 0, 1), cell('31 December 2026', 't0', 0, 2),
      cell('2.', 't0', 1, 0), cell('"Completion Date"', 't0', 1, 1), cell('15 January 2027', 't0', 1, 2),
      body('The Long-Stop Date and the Completion Date are milestones.'),
    );
    expect(fs).toHaveLength(0);
  });

  it('T9 — statutory citation and party-roles suppressed', () => {
    const fs = run(
      body('"Completion Date" means a date; the Completion Date applies.'),
      body('Each Director must comply with the Companies Act 2006 and the Act.'),
    );
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('T10 — cross-document detected, suppressed, counted, never resolved', () => {
    const ps = [
      body('"Permitted Encumbrances" has the meaning given in the Facility Agreement.'),
      body('The Permitted Encumbrances are listed in the disclosure bundle.'),
    ];
    const a = analyseDefinedTerms(ps, profile);
    expect(count(a.findings, 'used-not-defined')).toBe(0);
    expect(a.crossDocCount).toBe(1);
  });

  it('T11 — en-GB lowercase → capitalisation finding, never a spelling finding', () => {
    const fs = run(
      body('"Authorised Signatory" means a person authorised to sign.'),
      body('Each authorised signatory shall countersign.'),
    );
    expect(count(fs, 'capitalisation')).toBe(1);
    expect(fs.some(f => f.category.includes('spell'))).toBe(false);
  });

  it('T12 — longest-match: "Closing" and "Closing Date" both clean', () => {
    const fs = run(
      body('"Closing" means completion; "Closing Date" means the date of completion.'),
      body('On the Closing Date and at Closing the parties act.'),
    );
    expect(fs).toHaveLength(0);
  });

  it('T13 — S$ currency term; amount is not a separate undefined term', () => {
    const fs = run(
      body('"S$" means Singapore dollars.'),
      body('The price is S$1,000,000 and amounts are stated in S$ throughout.'),
    );
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('T14 — sentence-initial-only capitalisation is not used-but-not-defined', () => {
    const fs = run(
      body('"Completion Date" means a date; the Completion Date applies.'),
      body('Property of the Company. The property shall vest.'),
    );
    expect(count(fs, 'used-not-defined')).toBe(0);
  });

  it('T15 — empty definitions → "No definitions detected", everything quiet', () => {
    const a = analyseDefinedTerms(
      [body('The Buyer and the Seller agree on Material Terms and Conditions.')],
      profile,
    );
    expect(a.noDefinitions).toBe(true);
    expect(a.findings).toHaveLength(0);
  });
});
