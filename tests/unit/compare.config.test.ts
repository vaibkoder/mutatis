import { describe, it, expect } from 'vitest';
import {
  getCompareConfig,
  COMMON_SPELLING_VARIANTS,
  SG_PROFILE,
  UK_PROFILE,
} from '@shared/config/index.js';
import type { JurisdictionProfile, CompareConfig } from '@shared/config/index.js';

// ---------------------------------------------------------------------------
// getCompareConfig — default resolver
// ---------------------------------------------------------------------------

describe('getCompareConfig — defaults', () => {
  it('returns defaults when profile has no compare key (SG)', () => {
    const cfg = getCompareConfig(SG_PROFILE);
    expect(cfg.author).toBe('Mutatis Compare');
    expect(cfg.referenceTokens).toEqual(['Clause', 'Schedule', 'paragraph']);
    expect(cfg.renumberThreshold).toBe(3);
    expect(cfg.replaceSimilarityThreshold).toBe(0.5);
    // ADR-0047 — headers and footers are compared by default.
    expect(cfg.excludeHeaders).toBe(false);
    expect(cfg.excludeFooters).toBe(false);
    expect(cfg.footnotes).toBe('compare');
    expect(cfg.flattenOriginal).toBe('acceptAll');
    expect(cfg.flattenRevised).toBe('acceptAll');
  });

  it('returns defaults when profile has no compare key (UK)', () => {
    const cfg = getCompareConfig(UK_PROFILE);
    expect(cfg.author).toBe('Mutatis Compare');
    expect(cfg.renumberThreshold).toBe(3);
  });

  it('includes spelling variants in defaults', () => {
    const cfg = getCompareConfig(SG_PROFILE);
    expect(cfg.spellingVariants['organize']).toBe('organise');
    expect(cfg.spellingVariants['color']).toBe('colour');
    expect(cfg.spellingVariants['defense']).toBe('defence');
    expect(cfg.spellingVariants['program']).toBe('programme');
    expect(cfg.spellingVariants['fulfill']).toBe('fulfil');
    expect(cfg.spellingVariants['practice']).toBe('practise');
  });
});

describe('getCompareConfig — profile override merging', () => {
  function withCompare(overrides: Partial<CompareConfig>): JurisdictionProfile {
    return { ...SG_PROFILE, compare: { ...getCompareConfig(SG_PROFILE), ...overrides } };
  }

  it('overrides author', () => {
    const profile = withCompare({ author: 'Acme Firm Compare' });
    expect(getCompareConfig(profile).author).toBe('Acme Firm Compare');
  });

  it('overrides renumberThreshold', () => {
    const profile = withCompare({ renumberThreshold: 5 });
    expect(getCompareConfig(profile).renumberThreshold).toBe(5);
  });

  it('overrides footnotes mode', () => {
    const profile = withCompare({ footnotes: 'ignore' });
    expect(getCompareConfig(profile).footnotes).toBe('ignore');
  });

  it('overrides flattenRevised', () => {
    const profile = withCompare({ flattenRevised: 'rejectAll' });
    expect(getCompareConfig(profile).flattenRevised).toBe('rejectAll');
    // flattenOriginal still defaults
    expect(getCompareConfig(profile).flattenOriginal).toBe('acceptAll');
  });

  it('merges spellingVariants with defaults (additional entry wins)', () => {
    const profile = withCompare({ spellingVariants: { customus: 'customgb' } });
    const cfg = getCompareConfig(profile);
    // custom entry present
    expect(cfg.spellingVariants['customus']).toBe('customgb');
    // default entries still present
    expect(cfg.spellingVariants['organize']).toBe('organise');
  });

  it('allows profile to override an existing spelling entry', () => {
    const profile = withCompare({ spellingVariants: { organize: 'reorganise' } });
    const cfg = getCompareConfig(profile);
    expect(cfg.spellingVariants['organize']).toBe('reorganise');
  });

  it('uses provided referenceTokens without merging', () => {
    const profile = withCompare({ referenceTokens: ['Clause', 'Section', 'Annex'] });
    expect(getCompareConfig(profile).referenceTokens).toEqual(['Clause', 'Section', 'Annex']);
  });
});

// ---------------------------------------------------------------------------
// COMMON_SPELLING_VARIANTS — spot checks
// ---------------------------------------------------------------------------

describe('COMMON_SPELLING_VARIANTS', () => {
  it('contains no identity mappings (key === value)', () => {
    for (const [k, v] of Object.entries(COMMON_SPELLING_VARIANTS)) {
      expect(k).not.toBe(v);
    }
  });

  it('all keys are lowercase', () => {
    for (const k of Object.keys(COMMON_SPELLING_VARIANTS)) {
      expect(k).toBe(k.toLowerCase());
    }
  });

  it('all values are lowercase', () => {
    for (const v of Object.values(COMMON_SPELLING_VARIANTS)) {
      expect(v).toBe(v.toLowerCase());
    }
  });

  const pairs: [string, string][] = [
    ['organize', 'organise'],
    ['recognize', 'recognise'],
    ['analyze', 'analyse'],
    ['realise', 'realise'], // not in map — en-GB is already correct
    ['color', 'colour'],
    ['labor', 'labour'],
    ['behavior', 'behaviour'],
    ['defense', 'defence'],
    ['license', 'licence'],
    ['center', 'centre'],
    ['program', 'programme'],
    ['fulfill', 'fulfil'],
    ['practice', 'practise'],
  ];

  for (const [us, gb] of pairs) {
    if (us === gb) {
      it(`does NOT contain '${us}' (already en-GB)`, () => {
        expect(COMMON_SPELLING_VARIANTS[us]).toBeUndefined();
      });
    } else {
      it(`maps '${us}' → '${gb}'`, () => {
        expect(COMMON_SPELLING_VARIANTS[us]).toBe(gb);
      });
    }
  }
});
