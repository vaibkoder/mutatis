/**
 * C5 — Conditional show/hide tests.
 * C-T4: non-destructive hide then restore — no content loss.
 * C-T5: destructive delete only on confirm; irreversibility warned.
 *
 * HOST-ONLY coverage notes:
 *   - Logic (evaluateCondition, cache persistence) is fully unit-tested here.
 *   - Office JS getOoxml() / insertOoxml() round-trip is HOST-ONLY — not
 *     reachable in jsdom; tests for those branches verify the code path wiring
 *     via mock rather than live document fidelity.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { evaluateCondition } from '@features/content/templates/conditional.js';
import type { ShowWhenCondition, AnswersMap } from '@features/content/templates/types.js';

describe('C5 — evaluateCondition', () => {
  it('evaluates simple equals condition (string)', () => {
    const cond: ShowWhenCondition = { var: 'INCLUDES_STAFF', equals: 'yes' };
    expect(evaluateCondition(cond, { INCLUDES_STAFF: 'yes' })).toBe(true);
    expect(evaluateCondition(cond, { INCLUDES_STAFF: 'no' })).toBe(false);
  });

  it('evaluates equals condition (boolean true)', () => {
    const cond: ShowWhenCondition = { var: 'INCLUDES_STAFF', equals: true };
    expect(evaluateCondition(cond, { INCLUDES_STAFF: true })).toBe(true);
    expect(evaluateCondition(cond, { INCLUDES_STAFF: 'true' })).toBe(true);
    expect(evaluateCondition(cond, { INCLUDES_STAFF: false })).toBe(false);
    expect(evaluateCondition(cond, { INCLUDES_STAFF: 'false' })).toBe(false);
  });

  it('evaluates equals condition (boolean false)', () => {
    const cond: ShowWhenCondition = { var: 'HAS_WAIVER', equals: false };
    expect(evaluateCondition(cond, { HAS_WAIVER: false })).toBe(true);
    expect(evaluateCondition(cond, { HAS_WAIVER: 'false' })).toBe(true);
    expect(evaluateCondition(cond, { HAS_WAIVER: true })).toBe(false);
  });

  it('evaluates notEquals condition', () => {
    const cond: ShowWhenCondition = { var: 'JURISDICTION', notEquals: 'SG' };
    expect(evaluateCondition(cond, { JURISDICTION: 'EW' })).toBe(true);
    expect(evaluateCondition(cond, { JURISDICTION: 'SG' })).toBe(false);
  });

  it('evaluates AND condition', () => {
    const cond: ShowWhenCondition = {
      and: [
        { var: 'A', equals: true },
        { var: 'B', equals: 'yes' },
      ],
    };
    expect(evaluateCondition(cond, { A: true, B: 'yes' })).toBe(true);
    expect(evaluateCondition(cond, { A: true, B: 'no' })).toBe(false);
    expect(evaluateCondition(cond, { A: false, B: 'yes' })).toBe(false);
  });

  it('evaluates OR condition', () => {
    const cond: ShowWhenCondition = {
      or: [
        { var: 'A', equals: true },
        { var: 'B', equals: 'yes' },
      ],
    };
    expect(evaluateCondition(cond, { A: false, B: 'yes' })).toBe(true);
    expect(evaluateCondition(cond, { A: true, B: 'no' })).toBe(true);
    expect(evaluateCondition(cond, { A: false, B: 'no' })).toBe(false);
  });

  it('handles missing var as falsy (not undefined equality trap)', () => {
    const cond: ShowWhenCondition = { var: 'MISSING', equals: 'x' };
    expect(evaluateCondition(cond, {})).toBe(false);
  });

  it('evaluates nested AND/OR', () => {
    const cond: ShowWhenCondition = {
      or: [
        { and: [{ var: 'A', equals: true }, { var: 'B', equals: true }] },
        { var: 'C', equals: 'override' },
      ],
    };
    const answers: AnswersMap = { A: true, B: true, C: '' };
    expect(evaluateCondition(cond, answers)).toBe(true);
    answers.A = false;
    expect(evaluateCondition(cond, answers)).toBe(false);
    answers.C = 'override';
    expect(evaluateCondition(cond, answers)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// C-T4 / C-T5 logic wiring (HOST-ONLY for actual Office JS round-trip)
// -------------------------------------------------------------------------

vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(),
}));

vi.mock('@shared/storage/index.js', async () => {
  const { createLocalStore } = await import('@shared/storage/local.js');
  const store = createLocalStore();
  return {
    getStorageProvider: () => ({ local: store, docSettings: store }),
  };
});

import { hideConditionalNonDestructive, restoreConditional, deleteConditionalDestructive } from '@features/content/templates/conditional.js';
import { wordRun } from '@shared/office/index.js';

const mockWordRun = wordRun as ReturnType<typeof vi.fn>;

describe('C5 — non-destructive hide / restore wiring (C-T4)', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.clearAllMocks();
  });

  it('caches OOXML and clears content on hide', async () => {
    const cachedOoxml = '<w:body>SECTION_CONTENT</w:body>';

    mockWordRun.mockImplementationOnce(async (fn: (ctx: unknown) => Promise<unknown>) => {
      const ooxmlProxy = { value: cachedOoxml };
      const cc = {
        tag: 'SECTION_TUPE',
        getOoxml: vi.fn(() => ooxmlProxy),
        insertText: vi.fn(),
      };
      const ctx = {
        document: {
          contentControls: {
            load: vi.fn(),
            items: [cc],
          },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
      return undefined;
    });

    const result = await hideConditionalNonDestructive('SECTION_TUPE');
    expect(result.ok).toBe(true);

    // Verify the cache was persisted.
    const store = (await import('@shared/storage/index.js')).getStorageProvider();
    const cache = await store.docSettings.get('mutatis.hiddenSections');
    expect(cache).toBeTruthy();
    expect((cache as Record<string, string>)['SECTION_TUPE']).toBe(cachedOoxml);
  });

  it('restores from cache on re-show', async () => {
    // Pre-populate the hidden cache.
    const store = (await import('@shared/storage/index.js')).getStorageProvider();
    await store.docSettings.set('mutatis.hiddenSections', {
      SECTION_TUPE: '<w:body>RESTORED</w:body>',
    });

    const insertOoxmlSpy = vi.fn();
    mockWordRun.mockImplementationOnce(async (fn: (ctx: unknown) => Promise<unknown>) => {
      const cc = {
        tag: 'SECTION_TUPE',
        insertOoxml: insertOoxmlSpy,
      };
      const ctx = {
        document: {
          contentControls: {
            load: vi.fn(),
            items: [cc],
          },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
      return undefined;
    });

    const result = await restoreConditional('SECTION_TUPE');
    expect(result.ok).toBe(true);
    expect(insertOoxmlSpy).toHaveBeenCalledWith('<w:body>RESTORED</w:body>', expect.anything());

    // Cache entry should be removed after restore.
    const cache = await store.docSettings.get('mutatis.hiddenSections');
    expect((cache as Record<string, string>)['SECTION_TUPE']).toBeUndefined();
  });

  it('is idempotent — already hidden returns ok without re-caching', async () => {
    const store = (await import('@shared/storage/index.js')).getStorageProvider();
    await store.docSettings.set('mutatis.hiddenSections', {
      SECTION_X: '<w:body>EXISTING</w:body>',
    });

    // wordRun should NOT be called for an already-hidden section.
    const result = await hideConditionalNonDestructive('SECTION_X');
    expect(result.ok).toBe(true);
    expect(result.alreadyHidden).toBe(true);
    expect(mockWordRun).not.toHaveBeenCalled();
  });
});

describe('C5 — destructive delete wiring (C-T5)', () => {
  beforeEach(() => {
    localStorage.clear();
    vi.clearAllMocks();
  });

  it('calls control.delete(false) after confirm', async () => {
    const deleteSpy = vi.fn();
    mockWordRun.mockImplementationOnce(async (fn: (ctx: unknown) => Promise<unknown>) => {
      const cc = {
        tag: 'SECTION_TO_DELETE',
        delete: deleteSpy,
      };
      const ctx = {
        document: {
          contentControls: {
            load: vi.fn(),
            items: [cc],
          },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
      return undefined;
    });

    const result = await deleteConditionalDestructive('SECTION_TO_DELETE');
    expect(result.ok).toBe(true);
    // keepContent must be false (irreversible delete).
    expect(deleteSpy).toHaveBeenCalledWith(false);
  });
});
