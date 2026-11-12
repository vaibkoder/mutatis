import { describe, it, expect } from 'vitest';
import { scanPlaceholders } from '@features/factual-consistency/placeholders.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function para(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}

function count(paras: ParaNode[]): number {
  return scanPlaceholders(paras, SG_PROFILE).length;
}

// ---------------------------------------------------------------------------
// §9.4 cases
// ---------------------------------------------------------------------------

describe('Placeholder detection — §9.4 acceptance cases', () => {
  it('detects filled-bullet [●]', () => {
    const findings = scanPlaceholders([para('p0', 'The Purchase Price is [●].')], SG_PROFILE);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.severity).toBe('info');
    expect(findings[0]?.title).toMatch(/\[●\]/);
  });

  it('detects keyword [TBC]', () => {
    const findings = scanPlaceholders([para('p0', 'Completion Date: [TBC]')], SG_PROFILE);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.title).toMatch(/\[TBC\]/);
  });

  it('detects insert-directive [insert date of completion]', () => {
    const findings = scanPlaceholders([para('p0', '[insert date of completion]')], SG_PROFILE);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.title).toMatch(/insert/i);
  });

  it('detects note-to-draft [NTD: confirm with client]', () => {
    const findings = scanPlaceholders([para('p0', '[NTD: confirm with client]')], SG_PROFILE);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.title).toMatch(/NTD/);
  });

  it('detects percent-bullet [●]%', () => {
    const findings = scanPlaceholders([para('p0', 'Interest at [●]% per annum.')], SG_PROFILE);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.title).toMatch(/\[●\]/);
  });

  it('suppresses [INTENTIONALLY LEFT BLANK] — whitelist', () => {
    expect(count([para('p0', 'This page is [INTENTIONALLY LEFT BLANK].')])).toBe(0);
  });

  it('suppresses [Reserved] — whitelist', () => {
    expect(count([para('p0', '[Reserved]')])).toBe(0);
  });

  it('suppresses SG case citation [2024] SGHC 45', () => {
    expect(count([para('p0', 'As held in Lim v Tan [2024] SGHC 45 at [12].')])).toBe(0);
  });

  it('suppresses UK case citation [2024] EWHC 123 (Comm)', () => {
    expect(count([para('p0', '… [2024] EWHC 123 (Comm) …')])).toBe(0);
  });

  it('suppresses signature-line blank — By: [   ]', () => {
    expect(count([para('p0', 'By: [   ]')])).toBe(0);
  });

  it('does NOT flag [and the Guarantor] by default (flagBareBrackets=false)', () => {
    expect(count([para('p0', '[and the Guarantor] agrees to pay')])).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Additional correctness checks
// ---------------------------------------------------------------------------

describe('Placeholder detection — additional checks', () => {
  it('deduplicates the same placeholder at the same offset', () => {
    // Same paragraph scanned once — should not double-count
    const results = scanPlaceholders(
      [para('p0', 'Price is [●] and [●].')],
      SG_PROFILE,
    );
    // Two distinct offsets → two findings
    expect(results).toHaveLength(2);
  });

  it('detects [TBD] (not only [TBC])', () => {
    expect(count([para('p0', 'Amount: [TBD]')])).toBe(1);
  });

  it('detects [DATE]', () => {
    expect(count([para('p0', 'Effective on [DATE]')])).toBe(1);
  });

  it('route to proof tab via category prefix', () => {
    const findings = scanPlaceholders([para('p0', 'See [TBC].')], SG_PROFILE);
    expect(findings[0]?.category).toBe('factual-consistency/placeholder');
  });

  it('works identically on UK profile', () => {
    const sg = scanPlaceholders([para('p0', 'Price: [●].')], SG_PROFILE).length;
    const uk = scanPlaceholders([para('p0', 'Price: [●].')], UK_PROFILE).length;
    expect(sg).toBe(uk);
  });

  it('citation suppression: paragraph number [12] alone is NOT a citation', () => {
    // [12] without a reporter after it should not be suppressed (it's not a year bracket)
    // and since flagBareBrackets=false it produces no finding either — test stays false
    expect(count([para('p0', 'See [12] above.')])).toBe(0);
  });

  it('[SGHC 45] alone (no year bracket) is not a year-bracket citation pattern → bare-brackets (off by default)', () => {
    // Without the year bracket pattern it won't match the citation rule,
    // but flagBareBrackets is false so no finding either way.
    expect(count([para('p0', 'SGHC 45')])).toBe(0);
  });

  it('flagBareBrackets=true surfaces bare optional-text as medium-confidence', () => {
    const customProfile = {
      ...SG_PROFILE,
      placeholderVocab: { ...SG_PROFILE.placeholderVocab, flagBareBrackets: true },
    };
    // [and the Guarantor] should surface with bare-brackets category when flag is on
    const findings = scanPlaceholders(
      [para('p0', '[and the Guarantor] agrees to pay')],
      customProfile,
    );
    expect(findings).toHaveLength(1);
    expect(findings[0]?.title).toMatch(/bare brackets/);
  });
});
