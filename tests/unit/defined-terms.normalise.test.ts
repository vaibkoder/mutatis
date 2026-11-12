import { describe, it, expect } from 'vitest';
import {
  normaliseTerm, cleanTerm, foldQuotes, resolveLookup, buildAlternation,
} from '@features/defined-terms/normalise.js';

describe('normaliseTerm — spec §4.1', () => {
  it('folds curly quotes and strips surrounding quotes from the key', () => {
    expect(foldQuotes('“Completion Date”')).toBe('"Completion Date"');
    expect(normaliseTerm('“Completion Date”')).toBe('completion date');
    expect(normaliseTerm('"Completion Date"')).toBe('completion date');
  });

  it('collapses internal whitespace (incl. non-breaking space)', () => {
    expect(normaliseTerm('Completion  Date')).toBe('completion date');
  });

  it('strips a trailing possessive', () => {
    expect(normaliseTerm("Buyer's")).toBe('buyer');
    expect(normaliseTerm('“Buyer’s”')).toBe('buyer');
    expect(cleanTerm("Buyer's")).toBe('Buyer');
  });

  it('does NOT fold en-GB ↔ en-US spelling — distinct keys', () => {
    expect(normaliseTerm('Authorised')).not.toBe(normaliseTerm('Authorized'));
  });

  it('preserves currency glyphs and case in canonical form', () => {
    expect(cleanTerm('"S$"')).toBe('S$');
    expect(normaliseTerm('"S$"')).toBe('s$');
  });
});

describe('resolveLookup — plural / possessive lookup', () => {
  const map = new Map<string, number>([
    ['shareholder', 1],
    ['buyer', 1],
    ['company', 1],
    ['closing date', 1],
  ]);

  it('resolves a plural to the singular key', () => {
    expect(resolveLookup(map, 'Shareholders')).toBe('shareholder');
    expect(resolveLookup(map, 'Companies')).toBe('company');
  });

  it('resolves a possessive to the bare key', () => {
    expect(resolveLookup(map, "Buyer's")).toBe('buyer');
  });

  it('returns null when nothing matches (never invents a key)', () => {
    expect(resolveLookup(map, 'Seller')).toBeNull();
  });
});

describe('buildAlternation — longest-match-first', () => {
  it('matches "Completion Date" greedily over "Completion" + "Date"', () => {
    const re = buildAlternation(['Completion', 'Date', 'Completion Date']);
    expect(re).not.toBeNull();
    const text = 'on the Completion Date the work completes';
    const hits = [...text.matchAll(re!)].map(m => m[0]);
    expect(hits).toContain('Completion Date');
    expect(hits).not.toContain('Completion'); // consumed as part of the long match
  });

  it('matches case-insensitively but not as a partial word', () => {
    const re = buildAlternation(['Closing']);
    const hits = [...'at closing, the Closings list'.matchAll(re!)].map(m => m[0]);
    expect(hits).toEqual(['closing']); // "Closings" is not a Closing hit (plural lookup handles it)
  });

  it('returns null for an empty form list', () => {
    expect(buildAlternation([])).toBeNull();
  });
});
