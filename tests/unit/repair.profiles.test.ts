import { describe, it, expect } from 'vitest';
import {
  SG_COURT_FILING,
  UK_COURT_FILING,
  FIRM_COMMERCIAL,
  getProfile,
  listProfileIds,
  cloneProfile,
  convertMarginsToPoints,
} from '@features/repair/profiles.js';

describe('style profiles', () => {
  it('SG Court Filing has correct margins (99.21 pt)', () => {
    expect(SG_COURT_FILING.page.margins.top).toBeCloseTo(99.21, 1);
    expect(SG_COURT_FILING.page.margins.left).toBeCloseTo(99.21, 1);
  });

  it('SG Court Filing is locked and double-spaced', () => {
    expect(SG_COURT_FILING.locked).toBe(true);
    expect(SG_COURT_FILING.body.spacing.lineSpacingRule).toBe('multiple');
    expect(SG_COURT_FILING.body.spacing.lineSpacing).toBe(2.0);
    expect(SG_COURT_FILING.body.font.size).toBe(12);
  });

  it('UK Court Filing has correct margins (70.87 pt) and 1.5 spacing', () => {
    expect(UK_COURT_FILING.page.margins.top).toBeCloseTo(70.87, 1);
    expect(UK_COURT_FILING.body.spacing.lineSpacingRule).toBe('onePtFive');
    expect(UK_COURT_FILING.locked).toBe(true);
  });

  it('Firm Commercial has 72 pt margins and single spacing with spaceAfter', () => {
    expect(FIRM_COMMERCIAL.page.margins.top).toBe(72);
    expect(FIRM_COMMERCIAL.body.spacing.lineSpacingRule).toBe('single');
    expect(FIRM_COMMERCIAL.body.spacing.spaceAfter).toBe(6);
    expect(FIRM_COMMERCIAL.body.font.size).toBe(11);
    expect(FIRM_COMMERCIAL.locked).toBe(true);
  });

  it('getProfile retrieves by id', () => {
    expect(getProfile('sg-court-filing')).toBe(SG_COURT_FILING);
    expect(getProfile('uk-court-filing')).toBe(UK_COURT_FILING);
    expect(getProfile('firm-commercial')).toBe(FIRM_COMMERCIAL);
    expect(getProfile('nonexistent')).toBeUndefined();
  });

  it('listProfileIds returns all three', () => {
    const ids = listProfileIds();
    expect(ids).toContain('sg-court-filing');
    expect(ids).toContain('uk-court-filing');
    expect(ids).toContain('firm-commercial');
    expect(ids).toHaveLength(3);
  });

  it('cloneProfile produces an unlocked copy with new id/name', () => {
    const clone = cloneProfile(SG_COURT_FILING, 'my-profile', 'My Profile');
    expect(clone.id).toBe('my-profile');
    expect(clone.name).toBe('My Profile');
    expect(clone.locked).toBe(false);
    // Original unchanged
    expect(SG_COURT_FILING.locked).toBe(true);
    expect(SG_COURT_FILING.id).toBe('sg-court-filing');
  });

  it('convertMarginsToPoints converts 35mm → ~99.21 pt', () => {
    const mmProfile = cloneProfile(SG_COURT_FILING, 'test', 'Test');
    mmProfile.page.margins = { top: 35, bottom: 35, left: 35, right: 35 };
    const converted = convertMarginsToPoints(mmProfile, 'mm');
    expect(converted.page.margins.top).toBeCloseTo(99.21, 0);
    // source not mutated
    expect(mmProfile.page.margins.top).toBe(35);
  });

  it('convertMarginsToPoints with pt unit returns same object', () => {
    const result = convertMarginsToPoints(SG_COURT_FILING, 'pt');
    expect(result).toBe(SG_COURT_FILING);
  });
});
