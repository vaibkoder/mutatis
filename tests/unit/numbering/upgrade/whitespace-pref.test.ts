/**
 * Feature 06 upgrade (ADR-0023) — whitespace-override persistence (T3).
 * Constraint #3: one global override, persisted.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import {
  getWhitespaceOverride,
  setWhitespaceOverride,
  defaultWhitespaceOverride,
} from '@features/numbering/upgrade/whitespace-pref.js';

describe('whitespace-override persistence', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('defaults to "auto" when unset', async () => {
    expect(defaultWhitespaceOverride()).toBe('auto');
    expect(await getWhitespaceOverride()).toBe('auto');
  });

  it('round-trips a set value through storage', async () => {
    await setWhitespaceOverride('keep-one-tab');
    expect(await getWhitespaceOverride()).toBe('keep-one-tab');
  });

  it('rejects an unknown override value', async () => {
    // @ts-expect-error — exercising the runtime guard with a bad value.
    await expect(setWhitespaceOverride('nonsense')).rejects.toThrow();
  });
});
