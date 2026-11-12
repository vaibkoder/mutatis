/**
 * Feature 06 upgrade (ADR-0023) — sequence-anomaly detector (T2, M10/M12).
 * Constraint #4: surface both readings, never auto-resequence.
 */
import { describe, it, expect } from 'vitest';
import { detectSequenceAnomalies } from '@features/numbering/upgrade/sequence-anomaly.js';

describe('detectSequenceAnomalies', () => {
  it('returns nothing for a clean dotted sequence', () => {
    expect(detectSequenceAnomalies(['1.1', '1.2', '1.3'])).toEqual([]);
  });

  it('flags a duplicate with both readings', () => {
    const a = detectSequenceAnomalies(['1.4', '1.5', '1.5']);
    expect(a).toHaveLength(1);
    expect(a[0]!.kind).toBe('duplicate');
    expect(a[0]!.token).toBe('1.5');
    expect(a[0]!.readings.typo).toBeTruthy();
    expect(a[0]!.readings.reserved).toBeTruthy();
  });

  it('flags a gap with both readings', () => {
    const a = detectSequenceAnomalies(['1.5', '1.7']);
    expect(a).toHaveLength(1);
    expect(a[0]!.kind).toBe('gap');
    expect(a[0]!.readings.typo).toBeTruthy();
    expect(a[0]!.readings.reserved).toBeTruthy();
  });

  it('does not flag a level change across groups', () => {
    // "1.1" then "(a)" is a descent into a new level, not a gap.
    expect(detectSequenceAnomalies(['1.1', '(a)', '(b)'])).toEqual([]);
  });

  it('does not flag across a parent change (1.5 → 2.1)', () => {
    expect(detectSequenceAnomalies(['1.5', '2.1'])).toEqual([]);
  });

  it('handles lettered limb sequences (gap (a) → (c))', () => {
    const a = detectSequenceAnomalies(['(a)', '(c)']);
    expect(a).toHaveLength(1);
    expect(a[0]!.kind).toBe('gap');
  });

  it('handles roman limb duplicates ((ii),(ii))', () => {
    const a = detectSequenceAnomalies(['(i)', '(ii)', '(ii)']);
    expect(a).toHaveLength(1);
    expect(a[0]!.kind).toBe('duplicate');
  });

  it('never mutates input (no auto-resequence)', () => {
    const input = ['1.5', '1.5'];
    detectSequenceAnomalies(input);
    expect(input).toEqual(['1.5', '1.5']);
  });
});
