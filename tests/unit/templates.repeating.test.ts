/**
 * C6 — Fixed-slot repeating tests.
 * C-T6: 2 rows → slots 1–2 filled, 3–6 hidden; re-push with 4 rows re-shows 3–4.
 * C-T7: 7 rows into 6 slots → blocked with honest cap error; no cloning.
 *
 * HOST-ONLY: The actual Office JS insertText and hide calls are mocked.
 * Logic tested: cap enforcement, slot fill construction, hide/show sequencing.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('@shared/office/index.js', () => ({
  wordRun: vi.fn(),
  batchLoad: vi.fn(),
  processInChunks: vi.fn(),
}));

vi.mock('@shared/storage/index.js', async () => {
  const { createLocalStore } = await import('@shared/storage/local.js');
  const store = createLocalStore();
  return {
    getStorageProvider: () => ({ local: store, docSettings: store }),
  };
});

import { fillRepeatingGroup } from '@features/content/templates/repeating.js';
import type { RepeatingGroup } from '@features/content/templates/types.js';
import { SLOT_COUNT_MAX, HIDDEN_SECTIONS_KEY } from '@features/content/templates/types.js';
import { wordRun } from '@shared/office/index.js';
import { getStorageProvider } from '@shared/storage/index.js';

const mockWordRun = wordRun as ReturnType<typeof vi.fn>;

const TEST_GROUP: RepeatingGroup = {
  group: 'DIRECTORS',
  slotCount: 6,
  slotTags: [
    { field: 'name', tagPattern: 'DIRECTOR_{i}_NAME' },
    { field: 'uenOrCh', tagPattern: 'DIRECTOR_{i}_ID' },
  ],
  hideUnusedSlots: true,
};

beforeEach(() => {
  localStorage.clear();
  vi.clearAllMocks();
});

describe('C6 — cap enforcement (C-T7)', () => {
  it('returns a cap error when K > slotCount, no Office calls made', async () => {
    const rows = Array.from({ length: 7 }, (_, i) => ({
      name: `Director ${i + 1}`,
      uenOrCh: `ID${i}`,
    }));

    const result = await fillRepeatingGroup(TEST_GROUP, rows);

    expect(result.ok).toBe(false);
    expect(result.capError).toBeTruthy();
    expect(result.capError).toMatch(/up to 6/i);
    // The error must explain the limit without mentioning cloning as a solution.
    expect(result.capError).not.toMatch(/clone|cloning/i);
    expect(mockWordRun).not.toHaveBeenCalled();
  });

  it('accepts exactly slotCount rows without error', async () => {
    const rows = Array.from({ length: 6 }, (_, i) => ({
      name: `Director ${i + 1}`,
      uenOrCh: `ID${i}`,
    }));

    // wordRun mock that pretends all tags exist.
    mockWordRun.mockResolvedValue(undefined);

    const result = await fillRepeatingGroup(TEST_GROUP, rows);
    expect(result.ok).toBe(true);
    expect(result.capError).toBeUndefined();
  });
});

describe('C6 — cap error message quality', () => {
  it('states the cap clearly and does not mention cloning as an option', async () => {
    const rows = Array.from({ length: 8 }, (_, i) => ({ name: `D${i}`, uenOrCh: `X${i}` }));
    const result = await fillRepeatingGroup(TEST_GROUP, rows);
    expect(result.capError).toContain(`${TEST_GROUP.slotCount}`);
    expect(result.capError).not.toMatch(/clone|copy|duplicate slot/i);
    expect(result.capError).toContain('8'); // states the over-limit count
  });
});

describe('C6 — SLOT_COUNT_MAX constant', () => {
  it('is 6 (spec §C.4.3)', () => {
    expect(SLOT_COUNT_MAX).toBe(6);
  });
});

describe('C6 — fill result (C-T6)', () => {
  it('fills K=2 rows and reports ok with no capError', async () => {
    // Fill is done via a single wordRun call; hide ops use conditional.ts (mocked separately).
    mockWordRun.mockResolvedValue(undefined);

    const rows = [
      { name: 'Lim Wei', uenOrCh: '201912345A' },
      { name: 'S. Patel', uenOrCh: '198700456K' },
    ];

    const result = await fillRepeatingGroup(TEST_GROUP, rows);
    expect(result.ok).toBe(true);
    expect(result.capError).toBeUndefined();
  });

  it('builds correct fill tags for K=2 rows (DIRECTOR_1_NAME, DIRECTOR_2_NAME)', async () => {
    // Track the tag→value map built by the fill function before it calls wordRun.
    const capturedFills: Array<{ tag: string; value: string }> = [];

    mockWordRun.mockImplementationOnce(async (fn: (ctx: unknown) => Promise<unknown>) => {
      const allTags = [
        'DIRECTOR_1_NAME', 'DIRECTOR_1_ID',
        'DIRECTOR_2_NAME', 'DIRECTOR_2_ID',
      ];
      const items = allTags.map(t => ({
        tag: t,
        insertText: vi.fn((val: string) => { capturedFills.push({ tag: t, value: val }); }),
      }));
      const ctx = {
        document: {
          contentControls: { load: vi.fn(), items },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
      return undefined;
    });

    // Remaining wordRun calls (for restoreConditional + hideConditionalNonDestructive).
    mockWordRun.mockResolvedValue(undefined);

    const rows = [
      { name: 'Lim Wei', uenOrCh: '201912345A' },
      { name: 'S. Patel', uenOrCh: '198700456K' },
    ];

    await fillRepeatingGroup(TEST_GROUP, rows);

    const filledTags = capturedFills.map(c => c.tag);
    // Only slots 1 and 2 should have been written with non-empty values.
    expect(filledTags).toContain('DIRECTOR_1_NAME');
    expect(filledTags).toContain('DIRECTOR_2_NAME');
    // Values correct.
    const nameCall = capturedFills.find(c => c.tag === 'DIRECTOR_1_NAME');
    expect(nameCall?.value).toBe('Lim Wei');
  });
});

// ---------------------------------------------------------------------------
// C6 — re-fill cache consistency (2→4→2 stale hidden-cache regression)
// ---------------------------------------------------------------------------

/**
 * Drives fillRepeatingGroup against a stateful fake document so the
 * docSettings.hiddenSections cache (via conditional.ts hide/restore) is
 * genuinely exercised — not mocked away. Guards the stale-cache bug where a
 * slot hidden on one fill stayed cached after being re-filled, corrupting a
 * later hide/restore.
 */
function installFakeDoc(): Map<string, { content: string }> {
  const doc = new Map<string, { content: string }>();
  for (let i = 1; i <= 6; i++) {
    doc.set(`DIRECTOR_${i}_NAME`, { content: `[name ${i}]` });
    doc.set(`DIRECTOR_${i}_ID`, { content: `[id ${i}]` });
  }
  mockWordRun.mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
    const items = [...doc.entries()].map(([tag, slot]) => ({
      tag,
      getOoxml: () => ({ value: slot.content ? `<x>${slot.content}</x>` : '' }),
      insertText: (val: string) => { slot.content = val; },
      insertOoxml: (xml: string) => { slot.content = xml; },
      delete: () => { doc.delete(tag); },
    }));
    const ctx = {
      document: { contentControls: { load: () => undefined, items } },
      sync: async () => undefined,
    };
    return fn(ctx);
  });
  return doc;
}

describe('C6 — re-fill cache consistency (C-T6 2→4→2)', () => {
  const mkRows = (k: number) =>
    Array.from({ length: k }, (_, i) => ({ name: `D${i + 1}`, uenOrCh: `ID${i + 1}` }));

  async function cacheKeys(): Promise<string[]> {
    const cache =
      ((await getStorageProvider().docSettings.get(HIDDEN_SECTIONS_KEY)) as
        | Record<string, string>
        | undefined) ?? {};
    return Object.keys(cache).sort();
  }

  it('clears the hidden cache for a slot once it is re-filled (no stale entries)', async () => {
    const doc = installFakeDoc();

    // K=2: slots 1-2 filled, 3-6 hidden+cached.
    await fillRepeatingGroup(TEST_GROUP, mkRows(2));
    expect(await cacheKeys()).toEqual([
      'DIRECTOR_3_ID', 'DIRECTOR_3_NAME', 'DIRECTOR_4_ID', 'DIRECTOR_4_NAME',
      'DIRECTOR_5_ID', 'DIRECTOR_5_NAME', 'DIRECTOR_6_ID', 'DIRECTOR_6_NAME',
    ]);

    // K=4: slots 3-4 are now filled — their stale cache entries MUST be gone,
    // leaving only the still-hidden 5-6. (The bug left 3-6 all cached.)
    await fillRepeatingGroup(TEST_GROUP, mkRows(4));
    expect(await cacheKeys()).toEqual([
      'DIRECTOR_5_ID', 'DIRECTOR_5_NAME', 'DIRECTOR_6_ID', 'DIRECTOR_6_NAME',
    ]);
    expect(doc.get('DIRECTOR_3_NAME')?.content).toBe('D3');
    expect(doc.get('DIRECTOR_4_NAME')?.content).toBe('D4');

    // K=2 again (shrink): 3-6 hidden again, 1-2 filled and never cached.
    await fillRepeatingGroup(TEST_GROUP, mkRows(2));
    expect(await cacheKeys()).toEqual([
      'DIRECTOR_3_ID', 'DIRECTOR_3_NAME', 'DIRECTOR_4_ID', 'DIRECTOR_4_NAME',
      'DIRECTOR_5_ID', 'DIRECTOR_5_NAME', 'DIRECTOR_6_ID', 'DIRECTOR_6_NAME',
    ]);
    expect(doc.get('DIRECTOR_1_NAME')?.content).toBe('D1');
  });
});
