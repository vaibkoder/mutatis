/**
 * C1 — Template types schema tests.
 * Tests: slotCount guard, schema shape, SLOT_COUNT_MAX enforcement.
 */
import { describe, it, expect } from 'vitest';
import {
  guardSlotCount,
  SLOT_COUNT_DEFAULT,
  SLOT_COUNT_MAX,
} from '@features/content/templates/types.js';

describe('C1 — guardSlotCount', () => {
  it('accepts values 1..SLOT_COUNT_MAX', () => {
    for (let i = 1; i <= SLOT_COUNT_MAX; i++) {
      expect(guardSlotCount(i)).toBe(i);
    }
  });

  it('throws on slotCount > SLOT_COUNT_MAX', () => {
    expect(() => guardSlotCount(SLOT_COUNT_MAX + 1)).toThrow(RangeError);
    expect(() => guardSlotCount(SLOT_COUNT_MAX + 1)).toThrow(/cap/i);
  });

  it('throws on non-positive values', () => {
    expect(() => guardSlotCount(0)).toThrow(RangeError);
    expect(() => guardSlotCount(-1)).toThrow(RangeError);
  });

  it('SLOT_COUNT_DEFAULT is 4, SLOT_COUNT_MAX is 6', () => {
    expect(SLOT_COUNT_DEFAULT).toBe(4);
    expect(SLOT_COUNT_MAX).toBe(6);
  });
});
