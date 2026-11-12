/**
 * Pairing signals (ADR-0037; spec §6.3; acceptance M6).
 *
 * The signals are the part of pairing a user has to be able to trust from a
 * number in a grid, so each is tested for the property it is relied on for —
 * not merely for "returns a number between 0 and 1".
 *
 * M6 is the jurisdiction check and it is a real one: `03/04/2026` is 3 April in
 * SG and E&W. A month-first date pattern here would be both a `no-us-literals`
 * violation and a correctness bug, so the absence is asserted rather than
 * assumed.
 */
import { describe, it, expect } from 'vitest';

import {
  castOverlap,
  filenameSimilarity,
  jaccard,
  normaliseFileName,
} from '@features/compare/batch/pairing/signals.js';

const set = (...values: string[]): Set<string> => new Set(values);

describe('jaccard — the dominant signal', () => {
  it('is 1 for identical sets and 0 for disjoint ones', () => {
    expect(jaccard(set('a', 'b', 'c'), set('a', 'b', 'c'))).toBe(1);
    expect(jaccard(set('a', 'b'), set('c', 'd'))).toBe(0);
  });

  it('measures overlap against the union, not the smaller set', () => {
    // 2 shared of 4 distinct. A containment-style measure would say 1.0 here,
    // which would pair a short side letter with the SPA that quotes it.
    expect(jaccard(set('a', 'b'), set('a', 'b', 'c', 'd'))).toBeCloseTo(0.5, 10);
  });

  it('scores two empty sets 0, not 1', () => {
    // Two documents with no comparable content are not a confident match;
    // returning 1 would auto-pair them.
    expect(jaccard(set(), set())).toBe(0);
    expect(jaccard(set('a'), set())).toBe(0);
  });

  it('is symmetric', () => {
    const a = set('a', 'b', 'c');
    const b = set('b', 'c', 'd', 'e');
    expect(jaccard(a, b)).toBe(jaccard(b, a));
  });

  it('separates two versions of one agreement from two different agreements', () => {
    const spaV6 = set(...Array.from({ length: 100 }, (_, i) => `spa-${i}`));
    const spaV7 = set(
      ...Array.from({ length: 95 }, (_, i) => `spa-${i}`),
      'new-1', 'new-2', 'new-3',
    );
    const disclosureLetter = set(...Array.from({ length: 40 }, (_, i) => `dl-${i}`));

    expect(jaccard(spaV6, spaV7)).toBeGreaterThan(0.85);
    expect(jaccard(spaV6, disclosureLetter)).toBe(0);
  });
});

describe('normaliseFileName — M6', () => {
  it('strips every version word in the SG/UK vocabulary', () => {
    for (const word of [
      'engrossment', 'engrossed', 'conformed', 'initialled', 'execution',
      'blackline', 'redline', 'signed', 'final', 'clean', 'draft', 'copy', 'exec',
    ]) {
      expect(normaliseFileName(`Project Falcon SPA ${word}.docx`)).toBe('project falcon spa');
    }
  });

  it('strips version numbers and duplicate markers', () => {
    expect(normaliseFileName('SPA_v7.docx')).toBe('spa');
    expect(normaliseFileName('SPA v12.docx')).toBe('spa');
    expect(normaliseFileName('SPA (2).docx')).toBe('spa');
  });

  it('strips day-first dates in every form the spec names', () => {
    for (const date of [
      '2026-08-30', '30-08-2026', '30.08.26', '30 Aug 2026', '30 August 2026',
    ]) {
      expect(normaliseFileName(`SPA ${date}.docx`)).toBe('spa');
    }
  });

  it('does NOT strip a month-first date — the jurisdiction has no such format', () => {
    // 03/04/2026 is 3 April here. A month-first pattern would be a US default
    // (ADR-0002) and would mis-read the day-first form as a month.
    const source = normaliseFileName('SPA 12/25/2026.docx');
    expect(source).toContain('12/25/2026');
  });

  it('collapses the two real-world names for one document toward each other', () => {
    expect(normaliseFileName('Project_Falcon_SPA_v7_exec.docx')).toBe('project falcon spa');
    expect(normaliseFileName('SPA (clean) 2026-08-30.docx')).toBe('spa');
  });

  it('never returns an empty string for a name that was only version tokens', () => {
    // `final.docx` reduces to nothing; the caller scores 0 rather than dividing
    // by zero, which is checked in filenameSimilarity below.
    expect(normaliseFileName('final.docx')).toBe('');
  });
});

describe('filenameSimilarity', () => {
  it('is 1 for names that normalise identically', () => {
    expect(filenameSimilarity('SPA_v6.docx', 'SPA (final).docx')).toBe(1);
  });

  it('scores an unrelated pair far below a related one', () => {
    const related = filenameSimilarity('Project Falcon SPA v7.docx', 'Falcon SPA clean.docx');
    const unrelated = filenameSimilarity('Project Falcon SPA v7.docx', 'Tax Deed.docx');
    expect(related).toBeGreaterThan(unrelated);
    expect(unrelated).toBeLessThan(0.4);
  });

  it('is 0 when either name is nothing but version tokens', () => {
    expect(filenameSimilarity('final.docx', 'SPA.docx')).toBe(0);
  });

  it('is symmetric', () => {
    const a = 'Project Falcon SPA v7.docx';
    const b = 'falcon-spa-clean.docx';
    expect(filenameSimilarity(a, b)).toBeCloseTo(filenameSimilarity(b, a), 10);
  });

  it('tolerates punctuation and separator differences', () => {
    expect(filenameSimilarity('project_falcon_spa.docx', 'project-falcon-spa.docx')).toBe(1);
  });
});

describe('castOverlap — the tie-breaker', () => {
  const bundleCast = {
    definedTerms: set('completion date', 'purchase price', 'disclosure letter'),
    partyNames: set('seller', 'buyer'),
  };

  it('is near 1 across documents in one deal bundle — which is exactly why it cannot lead', () => {
    // The SPA, the disclosure letter and the side letter share the same cast
    // and most defined terms. This signal barely discriminates, so it is a
    // tie-breaker and deliberately absent from the weighted sum (score.ts).
    const sideLetter = {
      definedTerms: set('completion date', 'purchase price', 'disclosure letter'),
      partyNames: set('seller', 'buyer'),
    };
    expect(castOverlap(bundleCast, sideLetter)).toBe(1);
  });

  it('distinguishes documents from different deals', () => {
    const otherDeal = {
      definedTerms: set('facility', 'availability period'),
      partyNames: set('borrower', 'lender'),
    };
    expect(castOverlap(bundleCast, otherDeal)).toBe(0);
  });

  it('treats an empty cast as no information rather than a match', () => {
    expect(castOverlap(bundleCast, { definedTerms: set(), partyNames: set() })).toBe(0);
  });
});
