import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the shared office module so isCapable can be controlled per-test.
vi.mock('@shared/office/index.js', () => ({
  isCapable: vi.fn(() => false),
}));

import { isCapable } from '@shared/office/index.js';
import { applyTracked } from '@features/repair/apply/tracked.js';
import type { RepairFinding } from '@features/repair/types.js';

const mockIsCapable = isCapable as ReturnType<typeof vi.fn>;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeFinding(overrides: Partial<RepairFinding> = {}): RepairFinding {
  return {
    id: 'tracked.changes',
    category: 'tracked',
    severity: 'stabilisation',
    label: 'Accept all tracked changes',
    detail: '3 tracked changes found.',
    count: 3,
    destructive: true,
    requiresConfirm: true,
    selectedByDefault: false,
    action: 'acceptAll',
    result: null,
    ...overrides,
  };
}

/**
 * Build a minimal ctx mock.
 *
 * @param paraColours — per-paragraph font.color values returned after accept.
 *   Pass an empty array for "no paragraphs" (clean document).
 */
function makeCtx(paraColours: Array<string | null> = []): Word.RequestContext {
  const acceptAll = vi.fn();

  // Each paragraph proxy exposes a font.color property.
  const paraItems = paraColours.map((color) => ({
    load: vi.fn(),
    font: { color } as unknown as Word.Font,
  }));

  const paragraphsProxy = {
    load: vi.fn(),
    items: paraItems,
  };

  const trackedProxy = {
    acceptAll,
  };

  // sync call counter — first call = after acceptAll, subsequent calls = paragraph loads.
  const ctx = {
    document: {
      body: {
        getTrackedChanges: vi.fn(() => trackedProxy),
        paragraphs: paragraphsProxy,
        font: { color: null },
        load: vi.fn(),
      },
    },
    sync: vi.fn(async () => undefined),
  } as unknown as Word.RequestContext;

  return ctx;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('applyTracked', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  // 1. wordApi16 true + action acceptAll → acceptAll called, result:'done'
  it('calls acceptAll and returns result:done when wordApi16 is available', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx([]); // no coloured runs
    const finding = makeFinding();

    const results = await applyTracked(ctx, [finding], {});

    expect(
      (ctx.document.body.getTrackedChanges as ReturnType<typeof vi.fn>),
    ).toHaveBeenCalled();
    expect(
      (ctx.document.body.getTrackedChanges() as unknown as { acceptAll: ReturnType<typeof vi.fn> }).acceptAll,
    ).toHaveBeenCalled();
    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');
  });

  // 2. wordApi16 false → result:'skipped', acceptAll never called
  it('returns result:skipped and never calls acceptAll when wordApi16 is unavailable', async () => {
    mockIsCapable.mockReturnValue(false);
    const ctx = makeCtx([]);
    const finding = makeFinding();

    const results = await applyTracked(ctx, [finding], {});

    expect(
      (ctx.document.body.getTrackedChanges as ReturnType<typeof vi.fn>),
    ).not.toHaveBeenCalled();
    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('skipped');
  });

  // 3. stripInlineFormatting:false + coloured runs after accept → detail updated with residual warning
  it('appends residual colour warning when coloured runs exist and stripInlineFormatting is false', async () => {
    mockIsCapable.mockReturnValue(true);
    // Two paragraphs with non-black colours (simulates third-party tool artefacts).
    const ctx = makeCtx(['#FF0000', '#0000FF']);
    const finding = makeFinding({ detail: '2 tracked changes found.' });

    const results = await applyTracked(ctx, [finding], { stripInlineFormatting: false });

    expect(results[0].result).toBe('done');
    expect(results[0].detail).toContain('Warning:');
    expect(results[0].detail).toContain('residual colour');
    expect(results[0].detail).toContain('2 run(s)');
    expect(results[0].detail).toContain('strip inline formatting');
  });

  // 4. stripInlineFormatting:true + coloured runs → no residual warning (T9 will handle it)
  it('does not append residual warning when stripInlineFormatting is true', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx(['#FF0000', '#0000FF']);
    const finding = makeFinding({ detail: '2 tracked changes found.' });

    const results = await applyTracked(ctx, [finding], { stripInlineFormatting: true });

    expect(results[0].result).toBe('done');
    expect(results[0].detail).not.toContain('Warning:');
    expect(results[0].detail).not.toContain('residual colour');
  });

  // 5. Zero coloured runs → no warning appended to detail
  it('does not append a warning when all paragraphs have null/auto/#000000 font colour', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx([null, 'auto', '#000000', '']);
    const originalDetail = '4 tracked changes found.';
    const finding = makeFinding({ detail: originalDetail });

    const results = await applyTracked(ctx, [finding], { stripInlineFormatting: false });

    expect(results[0].result).toBe('done');
    expect(results[0].detail).toBe(originalDetail);
  });

  // 6. GeneralException on acceptAll → result:'error'
  it('returns result:error when acceptAll throws a GeneralException', async () => {
    mockIsCapable.mockReturnValue(true);

    // Override sync to throw on the first call (after acceptAll).
    const ctx = makeCtx([]);
    (ctx.sync as ReturnType<typeof vi.fn>).mockRejectedValueOnce(
      Object.assign(new Error('GeneralException'), { code: 'GeneralException' }),
    );

    const finding = makeFinding();
    const results = await applyTracked(ctx, [finding], {});

    expect(results[0].result).toBe('error');
  });

  // Extra: findings with wrong action are always skipped, regardless of capability.
  it('skips findings whose action is not acceptAll', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx([]);
    const wrongFinding = makeFinding({ action: 'removeComments' });

    const results = await applyTracked(ctx, [wrongFinding], {});

    expect(
      (ctx.document.body.getTrackedChanges as ReturnType<typeof vi.fn>),
    ).not.toHaveBeenCalled();
    expect(results[0].result).toBe('skipped');
  });

  // Extra: mixed findings — only acceptAll ones are processed.
  it('processes only acceptAll findings in a mixed array', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx([]);
    const good = makeFinding({ id: 'tracked.changes', action: 'acceptAll' });
    const bad = makeFinding({ id: 'other', action: 'removeComments' });

    const results = await applyTracked(ctx, [good, bad], {});

    expect(results).toHaveLength(2);
    expect(results.find((r) => r.id === 'tracked.changes')?.result).toBe('done');
    expect(results.find((r) => r.id === 'other')?.result).toBe('skipped');
  });
});
