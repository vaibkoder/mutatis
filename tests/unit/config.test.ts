import { describe, it, expect } from 'vitest';
import { getProfile, SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';

describe('config / jurisdiction profiles', () => {
  it('returns the SG profile with Clause as the top-level unit', () => {
    const sg = getProfile('SG');
    expect(sg.id).toBe('SG');
    expect(sg.terminology.topLevelUnit).toBe('Clause');
    expect(sg.terminology.scheduleWord).toBe('Schedule');
  });

  it('returns the UK profile with en-GB spelling and day-first dates', () => {
    const uk = getProfile('UK');
    expect(uk.id).toBe('UK');
    expect(uk.spellingLocale).toBe('en-GB');
    expect(uk.dateLocale.order).toBe('little-endian');
    expect(uk.dateLocale.weekendDays).toEqual([0, 6]);
    expect(uk.dateLocale.executionAnchorKeywords).toContain('executed');
    expect(uk.dateLocale.executionAnchorKeywords).toContain('completion date');
    expect(uk.dateLocale.effectiveDateKeywords).toContain('with effect from');
    expect(uk.dateLocale.effectiveDateKeywords).toContain('effective as of');
    expect(uk.dateLocale.flagAllWeekendDates).toBe(false);
    expect(uk.dateLocale.roleHintWindow).toBe(120);
  });

  it('SG profile has matching date locale config keys', () => {
    const sg = getProfile('SG');
    expect(sg.dateLocale.executionAnchorKeywords).toContain('execution date');
    expect(sg.dateLocale.effectiveDateKeywords).toContain('as of');
    expect(sg.dateLocale.flagAllWeekendDates).toBe(false);
    expect(sg.dateLocale.roleHintWindow).toBe(120);
  });

  it('never defaults a bare $ to USD (ambiguous policy)', () => {
    expect(SG_PROFILE.currency.table['$']?.code).toBe('AMBIGUOUS');
    expect(UK_PROFILE.currency.table['$']?.code).toBe('AMBIGUOUS');
    expect(SG_PROFILE.currency.ambiguousDollarPolicy).toBe('flag-soft');
  });

  it('uses jurisdiction-correct entity suffixes (no US forms)', () => {
    expect(UK_PROFILE.entityFields.suffixes).toContain('Ltd');
    expect(SG_PROFILE.entityFields.suffixes).toContain('Pte. Ltd.');
    expect(UK_PROFILE.entityFields.suffixes).not.toContain('Inc.');
    expect(SG_PROFILE.entityFields.suffixes).not.toContain('LLC');
  });

  it('suppresses US-litigation numbering preset for SG/UK', () => {
    const usPreset = UK_PROFILE.numberingPresets.find(p => p.jurisdiction === 'US');
    expect(usPreset?.suppressForSgUk).toBe(true);
  });

  it('commercial preset ships with 9 levels, default id, and SG/UK jurisdiction', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial');
    expect(commercial).toBeDefined();
    expect(commercial?.jurisdiction).toBe('SG-UK');
    expect(commercial?.suppressForSgUk).toBe(false);
    expect(commercial?.levels.length).toBe(9);
    expect(commercial?.builtIn).toBe(true);
    expect(commercial?.schemaVersion).toBe(1);
    expect(commercial?.unitLabel).toBe('clause');
  });

  it('commercial preset level 0 is arabic continuous (no restart)', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial');
    const l0 = commercial?.levels[0];
    expect(l0?.levelIndex).toBe(0);
    expect(l0?.listNumbering).toBe('arabic');
    expect(l0?.formatString).toBe('%1.');
    expect(l0?.restartAfterLevel).toBe(-1);
    expect(l0?.indent.textIndentPt).toBe(0);
  });

  it('commercial preset uses closed-bracket forms (a)/(i)/(A) at levels 3/4/5', () => {
    const commercial = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial');
    const levels = commercial?.levels ?? [];
    expect(levels[3]?.listNumbering).toBe('lowerLetter');
    expect(levels[3]?.formatString).toBe('(%4)');
    expect(levels[4]?.listNumbering).toBe('lowerRoman');
    expect(levels[4]?.formatString).toBe('(%5)');
    expect(levels[5]?.listNumbering).toBe('upperLetter');
    expect(levels[5]?.formatString).toBe('(%6)');
  });

  it('recitals preset uses upperLetter parenthesised form', () => {
    const recitals = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-recitals');
    expect(recitals?.levels[0]?.listNumbering).toBe('upperLetter');
    expect(recitals?.levels[0]?.formatString).toBe('(%1)');
    expect(recitals?.unitLabel).toBe('recital');
  });

  it('schedule preset has 4 levels with paragraph unit label', () => {
    const schedule = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-schedule');
    expect(schedule?.levels.length).toBe(4);
    expect(schedule?.unitLabel).toBe('paragraph');
  });

  it('plain-outline preset has all-arabic levels', () => {
    const outline = UK_PROFILE.numberingPresets.find(p => p.id === 'preset-plain-outline');
    expect(outline?.levels.every(l => l.listNumbering === 'arabic')).toBe(true);
  });

  it('SG profile carries the same enriched presets as UK (shared SG-UK hierarchy)', () => {
    const sgCommercial = SG_PROFILE.numberingPresets.find(p => p.id === 'preset-uk-sg-commercial');
    expect(sgCommercial).toBeDefined();
    expect(sgCommercial?.levels.length).toBe(9);
    expect(SG_PROFILE.numberingPresets.find(p => p.id === 'preset-recitals')).toBeDefined();
  });

  // --- Defined-Terms recognition keys (ADR-0013) ------------------------------

  for (const [name, profile] of [['SG', SG_PROFILE], ['UK', UK_PROFILE]] as const) {
    describe(`${name} defined-terms config`, () => {
      const dt = profile.definedTerms;

      it('carries the exclusion allowlist (party-role, statutory pattern, acronym)', () => {
        expect(dt.exclusions.partyRoles).toContain('Buyer');
        expect(dt.exclusions.acronyms).toContain('GST');
        expect(dt.exclusions.conventional).toContain('Business Day');
        expect(dt.exclusions.statutoryWords).toContain('the Act');
        // The statutory pattern matches "<X> Act <Year>" and "the Regulations" family.
        expect(dt.exclusions.statutoryCitation.test('Companies Act 2006')).toBe(true);
        expect(dt.exclusions.statutoryCitation.test('Companies Act 1967')).toBe(true);
        expect(dt.exclusions.statutoryCitation.test('Securities and Futures Act 2001')).toBe(true);
      });

      it('expanded verb set includes a "shall have the meaning" form and "is defined in"', () => {
        expect(dt.verbs).toContain('shall have the meaning given to it in');
        expect(dt.verbs).toContain('has the meaning ascribed to it in');
        expect(dt.verbs).toContain('is defined in');
        expect(dt.verbs).toContain('shall mean and include');
        // The original default verbs are preserved.
        expect(dt.verbs).toContain('means');
        expect(dt.verbs).toContain('has the meaning given to it in');
      });

      it('enables the broadened shapes by default; local redefinitions stay off', () => {
        expect(dt.recogniseUnquoted).toBe(true);
        expect(dt.recogniseColon).toBe(true);
        expect(dt.flagLocalRedefinitions).toBe(false);
      });
    });
  }
});
