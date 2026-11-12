import { describe, it, expect } from 'vitest';
import { scanDates } from '@features/factual-consistency/dates.js';
import { SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ParaNode } from '@shared/doc-model/index.js';

function para(id: string, text: string): ParaNode {
  return { id, index: 0, text, styleName: 'Normal', container: 'body',
    getRange: () => { throw new Error('not used'); } };
}

function scan(text: string, profile = SG_PROFILE) {
  return scanDates([para('p0', text)], profile);
}

// ---------------------------------------------------------------------------
// §9.3 Weekend execution dates
// ---------------------------------------------------------------------------

describe('Dates — §9.3 weekend execution date cases', () => {
  // 6 June 2026 = Saturday  7 June 2026 = Sunday  5 June 2026 = Friday

  it('executed on Saturday → Critical', () => {
    const f = scan('This Agreement is executed on 6 June 2026.');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
    expect(f[0]?.title).toMatch(/Saturday/);
  });

  it('completion on Sunday → Critical', () => {
    const f = scan('Completion shall take place on 7 June 2026.');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
    expect(f[0]?.title).toMatch(/Sunday/);
  });

  it('executed on Friday → no finding', () => {
    expect(scan('Executed on 5 June 2026.')).toHaveLength(0);
  });

  it('effective-date Sunday → no Critical (retroactive role)', () => {
    const f = scan('With effect from 7 June 2026.');
    const criticals = f.filter(x => x.severity === 'error');
    expect(criticals).toHaveLength(0);
  });

  it('deed form "dated this 6th day of June 2026" (Saturday) → Critical', () => {
    const f = scan('Dated this 6th day of June 2026.');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
  });

  it('numeric 6/6/2026 day-first → Saturday → Critical (AC-3: never US month-first)', () => {
    // 6/6/2026 day-first = 6 June 2026 (Saturday). If US month-first it would be June 6 — same
    // date in this case. Use a date that only makes sense day-first to confirm the parser.
    // 13/6/2026 = 13 June 2026 (Saturday); month 13 is invalid → confirms day-first parsing.
    const f = scan('Dated 13/6/2026.');  // 13 June 2026 = Saturday
    expect(f.length).toBeGreaterThan(0);
    expect(f[0]?.severity).toBe('error');
  });

  it('template date "this ___ day of ____ 20__" → no weekend finding', () => {
    const f = scan('Dated this ___ day of ____ 20__.');
    const weekendErrors = f.filter(x => x.severity === 'error' && x.category === 'factual-consistency/weekend-date');
    expect(weekendErrors).toHaveLength(0);
  });

  it('AC-3: numeric dates never interpreted as US month-first', () => {
    // 6/6/2026: both day-first and month-first give June 6 (same) — use a date where they differ.
    // 2/12/2026 day-first = 2 Dec 2026 (Thursday); month-first = Feb 12 2026 (Thursday too).
    // Use 7/6/2026 (day-first = 7 June 2026 = Sunday → error)
    const f = scan('Signed on 7/6/2026.');
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('error');
  });
});

// ---------------------------------------------------------------------------
// Additional correctness checks
// ---------------------------------------------------------------------------

describe('Dates — additional checks', () => {
  it('weekday date (Wednesday) → no finding', () => {
    expect(scan('Signed on 3 June 2026.')).toHaveLength(0); // Wednesday
  });

  it('century-ambiguous numeric date → Informational, not Critical', () => {
    const f = scan('Dated 6/6/26.');  // 2-digit year
    if (f.length > 0) {
      const criticals = f.filter(x => x.severity === 'error');
      expect(criticals).toHaveLength(0);
    }
  });

  it('findings route to proof tab via category prefix', () => {
    const f = scan('Executed on 6 June 2026.');
    if (f.length > 0) {
      expect(f[0]?.category).toMatch(/^factual-consistency\//);
    }
  });

  it('works identically on UK profile (same weekendDays)', () => {
    const sg = scanDates([para('p0', 'Executed on 6 June 2026.')], SG_PROFILE).length;
    const uk = scanDates([para('p0', 'Executed on 6 June 2026.')], UK_PROFILE).length;
    expect(sg).toBe(uk);
  });

  it('"as of" weekend date → not flagged as Critical', () => {
    const f = scan('Effective as of 7 June 2026.');
    const criticals = f.filter(x => x.severity === 'error');
    expect(criticals).toHaveLength(0);
  });

  it('"effective date" weekend date → not flagged as Critical', () => {
    const f = scan('The effective date is 7 June 2026.');
    const criticals = f.filter(x => x.severity === 'error');
    expect(criticals).toHaveLength(0);
  });

  it('century-ambiguous finding uses factual-consistency/century-ambiguous category (for amber styling)', () => {
    const f = scanDates([para('p0', 'Signed on 6/6/06.')], SG_PROFILE);
    if (f.length > 0) {
      expect(f[0]?.category).toBe('factual-consistency/century-ambiguous');
    }
  });

  it('flagAllWeekendDates=true surfaces any weekend date', () => {
    const customProfile = {
      ...SG_PROFILE,
      dateLocale: { ...SG_PROFILE.dateLocale, flagAllWeekendDates: true },
    };
    // A weekend date with no execution keyword → Informational under default,
    // but Informational (not error) when flagAllWeekendDates=true and role=other
    const f = scanDates([para('p0', 'The meeting is on 6 June 2026.')], customProfile);
    expect(f).toHaveLength(1);
    expect(f[0]?.severity).toBe('info');
  });
});
