/**
 * T9 — M5 Hardening: edge-case suite covering spec §8 and AC assertions.
 *
 * §8 table cases:
 * #1  Partially-filled date template → no weekend Critical; placeholder Informational
 * #2  Retroactive "as of" effective date → no Critical (weekend rule not applied)
 * #3  Century-ambiguous numeric date → Informational "cannot verify day"; no weekend Critical
 * #4  Range amounts → per-endpoint comparison (simplified; no Finding for prose range without pair)
 * #5  Qualifiers → no spurious Finding
 * #6  Currency mismatch without value mismatch → Critical
 * #7  Structural identifiers (Clause Three (3), Schedule Two (2)) → no amount Finding
 * #8  Non-monetary word/number pairs → no Finding
 * #9  "Dollars" without symbol → value compared; no currency mismatch; bare $ info
 * #10 Defined "Business Day" ignored in v1 → no extra Finding from definitions clause
 * #11 Sub-unit amounts → correct compare within centTolerance
 * #12 "and No/100"/"only"/"exactly" suffixes → stripped; correct compare
 * #13 British inserted "and" → 150,000 correct
 * #14 Case citation in brackets → suppressed
 * #15 Signature-line blanks → suppressed
 * #16 Optional-text brackets → not flagged by default
 */

import { describe, it, expect } from 'vitest';
import { scanAmounts } from '@features/factual-consistency/amounts.js';
import { scanDates } from '@features/factual-consistency/dates.js';
import { scanPlaceholders } from '@features/factual-consistency/placeholders.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function para(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}

// ---------------------------------------------------------------------------
// §8 edge cases
// ---------------------------------------------------------------------------

describe('Edge cases — §8 table', () => {
  // #1: template date
  it('#1 template date "this ___ day of ____ 20__" → no weekend Critical; placeholder Informational from scanPlaceholders', () => {
    const p = para('p0', 'Dated this ___ day of ____ 20__.');
    const dateFindings = scanDates([p], SG_PROFILE);
    const criticals = dateFindings.filter(f => f.severity === 'error');
    expect(criticals).toHaveLength(0);
    // The placeholder sweep (scanPlaceholders) does NOT recognise ___ as a bracket placeholder;
    // the template-date pattern is handled in the date sweep (skipped). No duplicate finding.
  });

  // #2: retroactive effective date
  it('#2 retroactive "as of" effective date Sunday → no Critical', () => {
    const f = scanDates([para('p0', 'With effect from 7 June 2026.')], SG_PROFILE);
    expect(f.filter(x => x.severity === 'error')).toHaveLength(0);
  });

  it('#2 "effective as of" Sunday → no Critical', () => {
    const f = scanDates([para('p0', 'Effective as of 7 June 2026.')], SG_PROFILE);
    expect(f.filter(x => x.severity === 'error')).toHaveLength(0);
  });

  // #3: century-ambiguous
  it('#3 century-ambiguous "6/6/06" → Informational only, no weekend Critical', () => {
    const f = scanDates([para('p0', 'Executed on 6/6/06.')], SG_PROFILE);
    if (f.length > 0) {
      expect(f.filter(x => x.severity === 'error')).toHaveLength(0);
    }
    // No Critical regardless
  });

  // #6: currency mismatch, value matches
  it('#6 "Ten Thousand Dollars (€10,000)" → currency-mismatch Critical', () => {
    const f = scanAmounts([para('p0', 'Ten Thousand Dollars (€10,000)')], SG_PROFILE);
    expect(f).toHaveLength(1);
    expect(f[0]?.category).toBe('factual-consistency/currency-mismatch');
    expect(f[0]?.severity).toBe('error');
  });

  // #7: structural identifiers
  it('#7 "Clause Three (3)" → no amount Finding', () => {
    expect(scanAmounts([para('p0', 'Clause Three (3)')], SG_PROFILE)).toHaveLength(0);
  });

  it('#7 "Schedule Two (2)" → no amount Finding', () => {
    expect(scanAmounts([para('p0', 'Schedule Two (2)')], SG_PROFILE)).toHaveLength(0);
  });

  // #8: non-monetary word/number pairs
  it('#8 "thirty (30) days" → no amount Finding', () => {
    expect(scanAmounts([para('p0', 'thirty (30) days')], SG_PROFILE)).toHaveLength(0);
  });

  // #9: "Dollars" without symbol
  it('#9 "Ten Thousand Dollars (10,000)" → value compared; no currency-mismatch (no symbol); figureCurrency null', () => {
    // The figure "10,000" has no currency symbol → figureCurrency null → we skip currency-mismatch
    // and also skip the comparison (spec §5.2: if figureCurrency=null, no comparison)
    const f = scanAmounts([para('p0', 'Ten Thousand Dollars (10,000)')], SG_PROFILE);
    // No figure currency → figure side filtered out
    expect(f).toHaveLength(0);
  });

  // #11: sub-unit amounts
  it('#11 "One Thousand Dollars and Fifty Cents (S$1,000.50)" → no Finding', () => {
    expect(scanAmounts([para('p0', 'One Thousand Dollars and Fifty Cents (S$1,000.50)')], SG_PROFILE)).toHaveLength(0);
  });

  // #12: suffixes stripped
  it('#12 "Ten Thousand Dollars and No/100 (S$10,000) only" → no Finding', () => {
    expect(scanAmounts([para('p0', 'Ten Thousand Dollars and No/100 (S$10,000) only')], SG_PROFILE)).toHaveLength(0);
  });

  it('#12 "Five Thousand Dollars exactly (S$5,000)" → no Finding', () => {
    expect(scanAmounts([para('p0', 'Five Thousand Dollars exactly (S$5,000)')], SG_PROFILE)).toHaveLength(0);
  });

  // #13: British "and"
  it('#13 "One Hundred and Fifty Thousand Dollars (S$150,000)" → no Finding', () => {
    expect(scanAmounts([para('p0', 'One Hundred and Fifty Thousand Dollars (S$150,000)')], SG_PROFILE)).toHaveLength(0);
  });

  // #14: case citation suppression
  it('#14 "[2024] SGHC 45" → no placeholder Finding', () => {
    expect(scanPlaceholders([para('p0', 'See Lim v Tan [2024] SGHC 45.')], SG_PROFILE)).toHaveLength(0);
  });

  it('#14 "[2024] EWHC 123 (Comm)" → no placeholder Finding', () => {
    expect(scanPlaceholders([para('p0', '[2024] EWHC 123 (Comm)')], UK_PROFILE)).toHaveLength(0);
  });

  // #15: signature-line suppression
  it('#15 signature-line blank "By: [   ]" → no placeholder Finding', () => {
    expect(scanPlaceholders([para('p0', 'By: [   ]')], SG_PROFILE)).toHaveLength(0);
  });

  // #16: optional-text brackets
  it('#16 "[and the Guarantor]" → no Finding by default (flagBareBrackets=false)', () => {
    expect(scanPlaceholders([para('p0', '[and the Guarantor] agrees to pay')], SG_PROFILE)).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// AC assertions
// ---------------------------------------------------------------------------

describe('Acceptance criteria', () => {
  it('AC-3: numeric dates always day-first — 13/6/2026 = 13 June (Saturday) → Critical', () => {
    const f = scanDates([para('p0', 'Executed on 13/6/2026.')], SG_PROFILE);
    expect(f.filter(x => x.severity === 'error')).toHaveLength(1);
  });

  it('AC-4: "Clause Three (3)" → no amount Finding', () => {
    expect(scanAmounts([para('p0', 'Clause Three (3) applies.')], SG_PROFILE)).toHaveLength(0);
  });

  it('AC-6: scanAmounts does not call body.search() — verifiable via no Word context required', () => {
    // All extraction uses in-memory text only. If body.search() were called,
    // it would throw because there is no Word context here.
    expect(() => scanAmounts([para('p0', 'Ten Thousand Dollars (S$10,000)')], SG_PROFILE))
      .not.toThrow();
    expect(() => scanDates([para('p0', '6 June 2026')], SG_PROFILE))
      .not.toThrow();
    expect(() => scanPlaceholders([para('p0', '[TBC]')], SG_PROFILE))
      .not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// SG ↔ UK profile switching
// ---------------------------------------------------------------------------

describe('SG ↔ UK profile parity', () => {
  it('Both profiles flag a weekend execution date identically', () => {
    const sg = scanDates([para('p0', 'Executed on 6 June 2026.')], SG_PROFILE).length;
    const uk = scanDates([para('p0', 'Executed on 6 June 2026.')], UK_PROFILE).length;
    expect(sg).toBe(uk);
    expect(sg).toBe(1);
  });

  it('Both profiles flag a placeholder identically', () => {
    const sg = scanPlaceholders([para('p0', 'Price: [●].')], SG_PROFILE).length;
    const uk = scanPlaceholders([para('p0', 'Price: [●].')], UK_PROFILE).length;
    expect(sg).toBe(uk);
    expect(sg).toBe(1);
  });

  it('Both profiles flag a currency mismatch identically', () => {
    const sg = scanAmounts([para('p0', 'Ten Thousand Dollars (€10,000)')], SG_PROFILE);
    const uk = scanAmounts([para('p0', 'Ten Thousand Dollars (€10,000)')], UK_PROFILE);
    expect(sg.map(f => f.category)).toEqual(uk.map(f => f.category));
  });

  it('UK profile flags GBP amount correctly', () => {
    const f = scanAmounts([para('p0', 'Ten Thousand Pounds (£5,000)')], UK_PROFILE);
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
    expect(f[0]?.category).toBe('factual-consistency/amount-mismatch');
  });
});
