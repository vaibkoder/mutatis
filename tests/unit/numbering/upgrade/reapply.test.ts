/**
 * Feature 06 upgrade (ADR-0023) — human-driven re-apply engine (T5, M11/M12).
 * Step-2 hanging exclusion: only previously-cleared (unlisted) paragraphs are
 * turned into a list; table cells and existing list items are left alone.
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';

// Isolate reapply's own logic — the apply/review primitives are tested elsewhere.
const applySchemeToList = vi.fn(async () => [] as number[]);
const detectMultipleSections = vi.fn(async () => false);
const reviewCrossReferences = vi.fn(async () => ({ count: 0, items: [], active: true }));

vi.mock('@features/numbering/apply.js', () => ({
  applySchemeToList: (...a: unknown[]) => applySchemeToList(...(a as [])),
  detectMultipleSections: (...a: unknown[]) => detectMultipleSections(...(a as [])),
}));
vi.mock('@features/numbering/review.js', () => ({
  reviewCrossReferences: (...a: unknown[]) => reviewCrossReferences(...(a as [])),
}));

import { startListAndApplyOnSelection } from '@features/numbering/upgrade/reapply.js';
import { UK_PROFILE } from '@shared/config/index.js';
import type { NumberingScheme } from '@features/numbering/types.js';

const SCHEME = { id: 's', name: 'S', levels: [] } as unknown as NumberingScheme;

interface ParaOpts { inList?: boolean; inTable?: boolean }

function makePara(opts: ParaOpts = {}) {
  return {
    parentTableCellOrNullObject: { isNullObject: !opts.inTable },
    listItemOrNullObject: { isNullObject: !opts.inList },
    startNewList: vi.fn(() => ({ id: 42 })),
    attachToList: vi.fn(),
    load: vi.fn(),
  };
}

function makeCtx(paras: ReturnType<typeof makePara>[]) {
  return {
    document: {
      body: {},
      getSelection: () => ({ paragraphs: { items: paras } }),
    },
    load: vi.fn(),
    sync: vi.fn(async () => {}),
  } as unknown as Word.RequestContext;
}

beforeEach(() => {
  applySchemeToList.mockClear();
  detectMultipleSections.mockClear();
  reviewCrossReferences.mockClear();
});

describe('startListAndApplyOnSelection', () => {
  it('starts a new list from cleared paragraphs and applies the scheme', async () => {
    const paras = [makePara(), makePara(), makePara()];
    const ctx = makeCtx(paras);

    const result = await startListAndApplyOnSelection(SCHEME, UK_PROFILE, ctx);

    expect(result).not.toBeNull();
    expect(result!.listed).toBe(3);
    expect(result!.report.listsAffected).toBe(1);
    // First paragraph seeds the list; the rest attach to it.
    expect(paras[0]!.startNewList).toHaveBeenCalledTimes(1);
    expect(paras[1]!.attachToList).toHaveBeenCalledWith(42, 0);
    expect(paras[2]!.attachToList).toHaveBeenCalledWith(42, 0);
    expect(applySchemeToList).toHaveBeenCalledTimes(1);
    expect(reviewCrossReferences).toHaveBeenCalledTimes(1);
  });

  it('excludes paragraphs that already belong to a list (hanging exclusion)', async () => {
    const cleared1 = makePara();
    const listed = makePara({ inList: true });
    const cleared2 = makePara();
    const ctx = makeCtx([cleared1, listed, cleared2]);

    const result = await startListAndApplyOnSelection(SCHEME, UK_PROFILE, ctx);

    expect(result!.listed).toBe(2);
    expect(cleared1.startNewList).toHaveBeenCalledTimes(1);
    expect(cleared2.attachToList).toHaveBeenCalledWith(42, 0);
    expect(listed.startNewList).not.toHaveBeenCalled();
    expect(listed.attachToList).not.toHaveBeenCalled();
  });

  it('excludes table-cell paragraphs and counts them', async () => {
    const cleared = makePara();
    const tableCell = makePara({ inTable: true });
    const ctx = makeCtx([cleared, tableCell]);

    const result = await startListAndApplyOnSelection(SCHEME, UK_PROFILE, ctx);

    expect(result!.listed).toBe(1);
    expect(result!.report.tableParasSkipped).toBe(1);
    expect(tableCell.startNewList).not.toHaveBeenCalled();
  });

  it('returns null when no eligible paragraph is present', async () => {
    const ctx = makeCtx([makePara({ inList: true }), makePara({ inTable: true })]);
    const result = await startListAndApplyOnSelection(SCHEME, UK_PROFILE, ctx);
    expect(result).toBeNull();
    expect(applySchemeToList).not.toHaveBeenCalled();
  });
});
