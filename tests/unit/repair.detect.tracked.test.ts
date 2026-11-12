import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the shared office module so isCapable can be controlled per-test.
vi.mock('@shared/office/index.js', () => ({
  isCapable: vi.fn(() => false),
}));

import { isCapable } from '@shared/office/index.js';
import { detectTracked } from '@features/repair/detect/tracked.js';

/** Build a minimal ctx mock that simulates Word.RequestContext. */
function makeCtx(itemCount: number): Word.RequestContext {
  const items = Array.from({ length: itemCount }, (_, i) => ({ id: `tc-${i}` }));
  const trackedProxy = {
    load: vi.fn(),
    items,
  };
  return {
    document: {
      body: {
        getTrackedChanges: vi.fn(() => trackedProxy),
      },
    },
    sync: vi.fn(async () => undefined),
  } as unknown as Word.RequestContext;
}

const mockIsCapable = isCapable as ReturnType<typeof vi.fn>;

describe('detectTracked', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns null when wordApi16 is not available (gate respected)', async () => {
    mockIsCapable.mockReturnValue(false);
    const ctx = makeCtx(5);
    const result = await detectTracked(ctx);
    expect(result).toBeNull();
    // Should not touch the document at all
    expect((ctx.document.body.getTrackedChanges as ReturnType<typeof vi.fn>)).not.toHaveBeenCalled();
  });

  it('returns null when wordApi16 is available but there are 0 tracked changes', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx(0);
    const result = await detectTracked(ctx);
    expect(result).toBeNull();
  });

  it('returns a finding with correct shape when 47 tracked changes are present', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx(47);
    const result = await detectTracked(ctx);

    expect(result).not.toBeNull();
    expect(result?.id).toBe('tracked.changes');
    expect(result?.category).toBe('tracked');
    expect(result?.severity).toBe('stabilisation');
    expect(result?.count).toBe(47);
    expect(result?.selectedByDefault).toBe(false);
    expect(result?.requiresConfirm).toBe(true);
    expect(result?.destructive).toBe(true);
    expect(result?.action).toBe('acceptAll');
    expect(result?.detail).toContain('47 tracked changes');
  });

  it('uses singular label for exactly 1 tracked change', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx(1);
    const result = await detectTracked(ctx);
    expect(result?.count).toBe(1);
    expect(result?.detail).toContain('1 tracked change ');
    expect(result?.detail).not.toContain('changes');
  });

  it('calls ctx.sync() exactly once when capability is present', async () => {
    mockIsCapable.mockReturnValue(true);
    const ctx = makeCtx(3);
    await detectTracked(ctx);
    expect(ctx.sync).toHaveBeenCalledOnce();
  });
});
