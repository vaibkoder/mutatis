/**
 * Feature 06 upgrade (ADR-0023) — token-span detector (T1, M9/M12).
 * Constraints #2 (batch guard), #3 (whitespace), #5 (leading-occurrence-only).
 */
import { describe, it, expect } from 'vitest';
import { computeRemovalSpan } from '@features/numbering/upgrade/token-span.js';

/** Helper: the text that survives after the computed deletion. */
function residue(text: string, override?: Parameters<typeof computeRemovalSpan>[1]): string {
  const span = computeRemovalSpan(text, override);
  if (!span) return text;
  return text.slice(0, span.start) + text.slice(span.end);
}

describe('computeRemovalSpan — detection & token', () => {
  it('detects dotted decimal at start', () => {
    const span = computeRemovalSpan('1.1 The parties agree.');
    expect(span).not.toBeNull();
    expect(span!.token).toBe('1.1');
    expect(residue('1.1 The parties agree.')).toBe('The parties agree.');
  });

  it('detects bare integer with a dot delimiter (flush)', () => {
    expect(residue('1. Introduction')).toBe('Introduction');
  });

  it('detects lettered limb (a)', () => {
    const span = computeRemovalSpan('(a) first limb');
    expect(span!.token).toBe('(a)');
    expect(residue('(a) first limb')).toBe('first limb');
  });

  it('detects roman limb (iv)', () => {
    const span = computeRemovalSpan('(iv) fourth');
    expect(span!.token).toBe('(iv)');
    expect(residue('(iv) fourth')).toBe('fourth');
  });

  it('detects upper-letter limb (A)', () => {
    expect(residue('(A) Recital')).toBe('Recital');
  });
});

describe('computeRemovalSpan — batch guard (#2)', () => {
  it('returns null for a hanging paragraph with no leading number', () => {
    expect(computeRemovalSpan('provided that the Buyer shall pay')).toBeNull();
  });

  it('does not flag "5 days after" (bare integer needs . ) or tab)', () => {
    expect(computeRemovalSpan('5 days after Completion')).toBeNull();
  });

  it('does not flag a date "30 June 2026"', () => {
    expect(computeRemovalSpan('30 June 2026')).toBeNull();
  });
});

describe('computeRemovalSpan — leading-occurrence-only (#5)', () => {
  it('matches only the leading number, never a later in-body occurrence', () => {
    expect(residue('1.5 see clause 1.5 above')).toBe('see clause 1.5 above');
  });

  it('handles leading whitespace before the number (deletes it — flush)', () => {
    expect(residue('   1.1 indented')).toBe('indented');
  });
});

describe('computeRemovalSpan — whitespace variants & override (#3)', () => {
  it('auto: collapses double-space to flush', () => {
    expect(residue('1.5  body')).toBe('body');
  });

  it('auto: strips a tab delimiter to flush', () => {
    expect(residue('1.5\tbody')).toBe('body');
  });

  it('auto: strips an NBSP trailing run to flush (constraint 3 — trailing run)', () => {
    // Space delimiter then an NBSP run; the whole horizontal run is removed.
    expect(residue('1.5 \u00A0\u00A0body')).toBe('body');
  });

  it('collapse-to-one-space: keeps exactly one whitespace char', () => {
    expect(residue('1.5   body', 'collapse-to-one-space')).toBe(' body');
  });

  it('strip-to-flush: removes all surrounding whitespace', () => {
    expect(residue('1.5   body', 'strip-to-flush')).toBe('body');
  });

  it('keep-one-tab: keeps a single tab when present in the trailing run', () => {
    expect(residue('1.5\t\tbody', 'keep-one-tab')).toBe('\tbody');
  });

  it('keep-one-tab: falls back to one space when no tab present', () => {
    expect(residue('1.5   body', 'keep-one-tab')).toBe(' body');
  });
});
