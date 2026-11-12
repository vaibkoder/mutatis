// tests/unit/repair.apply.spacing.test.ts
import { describe, it, expect, vi } from 'vitest';
import { applySpacing } from '../../src/features/repair/apply/spacing.js';
import { SG_COURT_FILING } from '../../src/features/repair/profiles.js';
import type { RepairFinding } from '../../src/features/repair/types.js';

// ---------------------------------------------------------------------------
// Mock helpers
// ---------------------------------------------------------------------------

interface MockParagraph {
  text: string;
  style?: string;
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
  delete: ReturnType<typeof vi.fn>;
  load: ReturnType<typeof vi.fn>;
}

function makePara(text: string, style = 'Normal'): MockParagraph {
  return {
    text,
    style,
    spaceBefore: 6,
    spaceAfter: 6,
    lineSpacing: 12,  // single-spaced in points
    delete: vi.fn(),
    load: vi.fn(),
  };
}

function makeCtx(paragraphs: MockParagraph[]) {
  const syncFn = vi.fn().mockResolvedValue(undefined);
  const ctx = {
    document: {
      body: {
        paragraphs: {
          load: vi.fn(),
          items: paragraphs,
        },
      },
    },
    sync: syncFn,
  } as unknown as Word.RequestContext;
  return { ctx, syncFn };
}

function makeNormaliseFinding(): RepairFinding {
  return {
    id: 'spacing.anomaly',
    category: 'spacing',
    severity: 'stabilisation',
    label: 'Spacing anomalies',
    count: 3,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: false,
    action: 'normaliseSpacing',
  };
}

function makeBlankParaFinding(): RepairFinding {
  return {
    id: 'blankParas.count',
    category: 'blankParas',
    severity: 'stabilisation',
    label: 'Blank paragraphs',
    count: 3,
    destructive: false,
    requiresConfirm: false,
    selectedByDefault: true,
    action: 'deleteBlankParas',
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('applySpacing', () => {
  const profile = SG_COURT_FILING;

  // 1. normaliseSpacing: spaceBefore assigned from profile
  it('normaliseSpacing: assigns spaceBefore from profile to all paragraphs', async () => {
    const paras = [makePara('Para 1'), makePara('Para 2'), makePara('Para 3')];
    const { ctx } = makeCtx(paras);

    const results = await applySpacing(ctx, [makeNormaliseFinding()], profile, {});

    expect(results).toHaveLength(1);
    expect(results[0].result).toBe('done');
    for (const p of paras) {
      expect(p.spaceBefore).toBe(profile.body.spacing.spaceBefore);
    }
  });

  // 2. normaliseSpacing: lineSpacing (in pts) and spaceAfter assigned from profile
  it('normaliseSpacing: assigns lineSpacingRule from profile', async () => {
    const paras = [makePara('A'), makePara('B')];
    const { ctx } = makeCtx(paras);

    await applySpacing(ctx, [makeNormaliseFinding()], profile, {});

    // lineSpacing multiple × 12 = points
    const expectedLineSpacingPt = profile.body.spacing.lineSpacing * 12;
    for (const p of paras) {
      expect(p.lineSpacing).toBe(expectedLineSpacingPt);
      expect(p.spaceAfter).toBe(profile.body.spacing.spaceAfter);
    }
  });

  // 3. deleteBlankParas deleteAll: 3 blank paragraphs in 5 → 3 deleted
  it('deleteBlankParas deleteAll: deletes all blank paragraphs except the last para guard', async () => {
    const paras = [
      makePara('Content 1'),
      makePara(''),            // blank → delete
      makePara('Content 2'),
      makePara(''),            // blank → delete
      makePara(''),            // blank but last → NOT deleted
    ];
    const { ctx } = makeCtx(paras);

    const results = await applySpacing(ctx, [makeBlankParaFinding()], profile, {
      blankParaMode: 'deleteAll',
    });

    expect(results[0].result).toBe('done');
    // paras[1] and paras[3] should be deleted; paras[4] is last → preserved
    expect(paras[1].delete).toHaveBeenCalledOnce();
    expect(paras[3].delete).toHaveBeenCalledOnce();
    expect(paras[4].delete).not.toHaveBeenCalled();
    // paras[0] and paras[2] are non-blank → never deleted
    expect(paras[0].delete).not.toHaveBeenCalled();
    expect(paras[2].delete).not.toHaveBeenCalled();
  });

  // 4. deleteBlankParas collapseMultiple: 3 consecutive blanks → only first retained, 2 deleted
  it('deleteBlankParas collapseMultiple: 3 consecutive blanks → first retained, 2 deleted', async () => {
    const paras = [
      makePara('Content'),
      makePara(''),   // blank run start → retain
      makePara(''),   // 2nd consecutive → delete
      makePara(''),   // 3rd consecutive → delete
      makePara('End'),
    ];
    const { ctx } = makeCtx(paras);

    const results = await applySpacing(ctx, [makeBlankParaFinding()], profile, {
      blankParaMode: 'collapseMultiple',
    });

    expect(results[0].result).toBe('done');
    expect(paras[1].delete).not.toHaveBeenCalled(); // first in run: retained
    expect(paras[2].delete).toHaveBeenCalledOnce(); // 2nd: deleted
    expect(paras[3].delete).toHaveBeenCalledOnce(); // 3rd: deleted
    expect(paras[0].delete).not.toHaveBeenCalled();
    expect(paras[4].delete).not.toHaveBeenCalled();
  });

  // 5. Last paragraph blank → NOT deleted (guard)
  it('deleteBlankParas deleteAll: last-paragraph guard prevents deletion of sole final blank', async () => {
    const paras = [
      makePara('Content'),
      makePara(''),   // last paragraph → must not be deleted
    ];
    const { ctx } = makeCtx(paras);

    await applySpacing(ctx, [makeBlankParaFinding()], profile, { blankParaMode: 'deleteAll' });

    expect(paras[1].delete).not.toHaveBeenCalled();
  });

  // 6. Empty findings list → returns empty array, no crash
  it('returns an empty array and does not crash when findings is empty', async () => {
    const paras = [makePara('Something')];
    const { ctx, syncFn } = makeCtx(paras);

    const results = await applySpacing(ctx, [], profile, {});

    expect(results).toEqual([]);
    expect(syncFn).not.toHaveBeenCalled();
  });

  // 7. Chunking: 250 paragraphs → ctx.sync called multiple times
  it('normaliseSpacing: 250 paragraphs → ctx.sync called multiple times (chunked)', async () => {
    const paras = Array.from({ length: 250 }, (_, i) => makePara(`Para ${i}`));
    const { ctx, syncFn } = makeCtx(paras);

    await applySpacing(ctx, [makeNormaliseFinding()], profile, {});

    // Expect at least 2 syncs from normaliseSpacing (initial load sync + 2 chunk syncs for 250 paras)
    // Chunk size = 200 → chunk 1 (0-199) sync + chunk 2 (200-249) sync + initial load sync = 3 total
    expect(syncFn.mock.calls.length).toBeGreaterThanOrEqual(3);
  });
});
