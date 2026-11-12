/**
 * `normaliseTextWithOffsets` is a second implementation of `normaliseText`
 * (ADR-0045). Duplicating a rule is a defect unless it is held to the
 * first — this is the hold. The replace author also re-checks it at run time
 * and falls back on any disagreement, but a disagreement nothing measures would
 * silently send every paragraph with an object down the old, lossy path.
 */
import { describe, it, expect } from 'vitest';
import { normaliseText, normaliseTextWithOffsets } from '@features/compare/preprocess/normalise.js';

/** Every character class `normaliseText` treats specially, plus ordinary text. */
const ALPHABET = [
  'a', 'B', '1', '.', ',', '(', ')', ' ', ' ', '\t', '\n', '\r',
  ' ', ' ', ' ', '“', '”', '‘', '’', '—', '–',
  ' ', '﻿', '"', "'", '-',
];

/** Deterministic PRNG (mulberry32) so a failure reproduces. */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe('normaliseTextWithOffsets', () => {
  it('produces exactly normaliseText, on 5,000 random strings', () => {
    const next = rng(20261017);
    for (let n = 0; n < 5000; n++) {
      const len = Math.floor(next() * 24);
      let s = '';
      for (let i = 0; i < len; i++) s += ALPHABET[Math.floor(next() * ALPHABET.length)]!;
      const { text, offsets } = normaliseTextWithOffsets(s);
      expect(text, JSON.stringify(s)).toBe(normaliseText(s));
      expect(offsets).toHaveLength(text.length);
      // Offsets point into the input, never backwards.
      for (let i = 0; i < offsets.length; i++) {
        expect(offsets[i]).toBeGreaterThanOrEqual(0);
        expect(offsets[i]).toBeLessThan(s.length);
        if (i > 0) expect(offsets[i]).toBeGreaterThanOrEqual(offsets[i - 1]!);
      }
    }
  });

  it('maps an em-dash expansion and a collapsed run back to their sources', () => {
    const { text, offsets } = normaliseTextWithOffsets('a—b   c');
    expect(text).toBe('a -- b c');
    expect(offsets).toEqual([0, 1, 1, 1, 1, 2, 3, 6]);
  });
});
