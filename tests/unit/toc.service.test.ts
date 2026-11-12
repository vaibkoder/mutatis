/**
 * tests/unit/toc.service.test.ts
 *
 * Unit tests for A2, A3, A4: toc-service.ts (Office JS operations).
 *
 * All Office JS and storage dependencies are mocked. Tests verify:
 *  - Switch string is correctly set on field.code (A2)
 *  - Capability-absent path uses insertOoxml + sets no fieldId (A2/A-T5)
 *  - Registry entry is persisted after insert (A2)
 *  - Refresh-all loads fields once and updates each (A3)
 *  - Missing field is pruned from registry after refresh-all (A3)
 *  - Scoped insert appends \b BookmarkName (A4)
 *  - Duplicate bookmark name is counter-suffixed to be unique (A4)
 *
 * HOST-ONLY CAVEATS:
 *  - Actual TOC rendering after updateResult() — needs live Word host.
 *  - field.id availability/uniqueness across hosts — needs live host verification.
 *  - w:fldSimple rendering via insertOoxml — needs live host verification.
 *  - range.insertBookmark behaviour — needs live host verification.
 */

import { describe, it, expect, vi, beforeEach, afterAll } from 'vitest';

// ---------------------------------------------------------------------------
// Global Word stub — provides enum values referenced inside wordRun callbacks.
// Must be installed before any test that invokes service code.
// ---------------------------------------------------------------------------

function installWordStub() {
  (globalThis as unknown as { Word: unknown }).Word = {
    InsertLocation: { before: 'Before', after: 'After', replace: 'Replace', start: 'Start', end: 'End' },
    FieldType: { toc: 'Toc', hyperlink: 'Hyperlink', pageNumber: 'PageNumber' },
    ContentControlType: { plainText: 'PlainText', richText: 'RichText' },
  };
}

installWordStub();

afterAll(() => {
  delete (globalThis as unknown as { Word?: unknown }).Word;
});

// ---------------------------------------------------------------------------
// Mock: @shared/office/index.js
// ---------------------------------------------------------------------------
// We control isCapable and wordRun per test.

const _wordRunImpl = vi.fn(async (fn: (ctx: unknown) => Promise<unknown>) => fn(_mockCtx));

vi.mock('@shared/office/index.js', () => ({
  isCapable: vi.fn(() => true),       // default: wordApi15 present
  wordRun: (fn: (ctx: unknown) => Promise<unknown>) => _wordRunImpl(fn),
  batchLoad: vi.fn(),
  processInChunks: vi.fn(),
  detectCapabilities: vi.fn(),
  getCapabilities: vi.fn(),
  _resetCapabilities: vi.fn(),
  waitForOffice: vi.fn(),
  _resetReady: vi.fn(),
}));

// ---------------------------------------------------------------------------
// Mock: @shared/storage/index.js
// ---------------------------------------------------------------------------

const _docSettingsStore: Map<string, unknown> = new Map();

vi.mock('@shared/storage/index.js', () => ({
  getStorageProvider: vi.fn(() => ({
    docSettings: {
      get: vi.fn(async (key: string) => _docSettingsStore.get(key)),
      set: vi.fn(async (key: string, value: unknown) => { _docSettingsStore.set(key, value); }),
      delete: vi.fn(async (key: string) => { _docSettingsStore.delete(key); }),
      list: vi.fn(async () => []),
      export: vi.fn(async () => new Blob(['{}'])),
      import: vi.fn(async () => {}),
    },
    local: {
      get: vi.fn(), set: vi.fn(), delete: vi.fn(), list: vi.fn(async () => []),
      export: vi.fn(), import: vi.fn(),
    },
  })),
}));

// ---------------------------------------------------------------------------
// Fake Word context object
// ---------------------------------------------------------------------------

function makeField(idVal = 'field-001') {
  const resultProxy = {
    text: 'Clause 1\t1\nClause 2\t2',
    select: vi.fn(),
  };
  return {
    id: idVal,
    code: '',
    result: resultProxy,
    updateResult: vi.fn(),
    load: vi.fn(),
  };
}

const _mockFields = {
  items: [] as ReturnType<typeof makeField>[],
  load: vi.fn(),
};

const _mockBookmarks = {
  items: [] as Array<{ name: string }>,
  load: vi.fn(),
};

const _mockBody = {
  getRange: vi.fn(() => _mockRange),
};

const _mockSelection = {
  getRange: vi.fn(() => _mockRange),
  insertBookmark: vi.fn(),
};

const _mockRange = {
  insertField: vi.fn(() => makeField()),
  insertOoxml: vi.fn(),
  insertBookmark: vi.fn(),
};

const _mockCtx = {
  document: {
    body: _mockBody,
    getSelection: vi.fn(() => _mockSelection),
    fields: {
      getByTypes: vi.fn(() => _mockFields),
    },
    bookmarks: _mockBookmarks,
  },
  sync: vi.fn(async () => {}),
};

// Helper to reset the mock context between tests.
function resetMockCtx() {
  _docSettingsStore.clear();
  _mockFields.items = [];
  _mockFields.load.mockClear();
  _mockBookmarks.items = [];
  _mockBookmarks.load.mockClear();
  _mockBody.getRange.mockClear();
  _mockRange.insertField.mockReset();
  _mockRange.insertOoxml.mockClear();
  _mockRange.insertBookmark.mockClear();
  _mockCtx.document.getSelection.mockClear();
  _mockCtx.document.fields.getByTypes.mockClear();
  _mockCtx.sync.mockClear();

  // Default insertField returns a fresh field.
  _mockRange.insertField.mockImplementation(() => makeField());
}

// ---------------------------------------------------------------------------
// Import under test (after mocks are declared)
// ---------------------------------------------------------------------------

import { isCapable } from '@shared/office/index.js';

import {
  insertToc,
  insertScopedToc,
  insertScheduleIndexField,
  refreshAllTocs,
  loadRegistry,
  clearRegistry,
  generateBookmarkName,
  buildFldSimpleOoxml,
} from '@features/content/toc/toc-service.js';
import { HOUSE_DEFAULT_PROFILE } from '@features/content/toc/config.js';
import type { ScheduleIndexConfig } from '@features/content/toc/types.js';

// ---------------------------------------------------------------------------
// A2: F9-fallback OOXML escaping (buildFldSimpleOoxml)
// ---------------------------------------------------------------------------

describe('buildFldSimpleOoxml — XML-attribute escaping', () => {
  it('escapes &, <, > and " in the instruction so the w:instr attribute is well-formed', () => {
    // A firm style name with an ampersand in the switch string.
    const ooxml = buildFldSimpleOoxml(' TOC \\t "Heading & <Title>,1" \\h ');
    expect(ooxml).toContain('&amp;');
    expect(ooxml).toContain('&lt;Title&gt;');
    expect(ooxml).toContain('&quot;');
    // No raw metacharacters survive inside the attribute value.
    const attr = ooxml.slice(ooxml.indexOf('w:instr="') + 9, ooxml.indexOf('"', ooxml.indexOf('w:instr="') + 9));
    expect(attr).not.toMatch(/[<>]/);
    expect(attr).not.toMatch(/&(?!amp;|lt;|gt;|quot;|apos;)/);
  });

  it('leaves a plain instruction unchanged', () => {
    const ooxml = buildFldSimpleOoxml(' TOC \\o "1-3" \\h ');
    expect(ooxml).toContain('w:instr=" TOC \\o &quot;1-3&quot; \\h "');
  });
});

// ---------------------------------------------------------------------------
// A2: insertToc — native path
// ---------------------------------------------------------------------------

describe('insertToc — native path (wordApi15 present)', () => {
  beforeEach(() => {
    resetMockCtx();
    vi.mocked(isCapable).mockReturnValue(true);
  });

  it('sets the correct field.code on the inserted field', async () => {
    let capturedCode = '';
    _mockRange.insertField.mockImplementation(() => {
      const f = makeField();
      // Capture code assignment at the end of the test by checking what was set.
      Object.defineProperty(f, 'code', {
        set(val: string) { capturedCode = val; },
        get() { return capturedCode; },
      });
      return f;
    });

    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'TOC' });
    expect(capturedCode).toContain('TOC');
    expect(capturedCode).toContain('\\t');
    expect(capturedCode).toContain('\\h');
  });

  it('persists a registry entry with fieldId from field.id', async () => {
    await clearRegistry();
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'Main TOC' });

    const reg = await loadRegistry();
    expect(reg).toHaveLength(1);
    expect(reg[0]!.scopeLabel).toBe('Main TOC');
    expect(reg[0]!.fieldId).toBe('field-001');
  });

  it('returns usedNativeField:true and usedFallback:false', async () => {
    const result = await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    expect(result.usedNativeField).toBe(true);
    expect(result.usedFallback).toBe(false);
  });

  it('inserts at topOfDocument when placement is topOfDocument', async () => {
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'topOfDocument', scopeLabel: 'x' });
    expect(_mockBody.getRange).toHaveBeenCalled();
    expect(_mockCtx.document.getSelection).not.toHaveBeenCalled();
  });

  it('inserts at cursor when placement is atCursor', async () => {
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    expect(_mockCtx.document.getSelection).toHaveBeenCalled();
  });

  it('calls updateResult() on the field', async () => {
    let updateCalled = false;
    _mockRange.insertField.mockImplementation(() => {
      const f = makeField();
      f.updateResult = vi.fn(() => { updateCalled = true; });
      return f;
    });
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    expect(updateCalled).toBe(true);
  });

  it('appends bookmark switch when opts.bookmark is provided', async () => {
    let capturedCode = '';
    _mockRange.insertField.mockImplementation(() => {
      const f = makeField();
      Object.defineProperty(f, 'code', {
        set(val: string) { capturedCode = val; },
        get() { return capturedCode; },
      });
      return f;
    });

    await insertToc(HOUSE_DEFAULT_PROFILE, {
      placement: 'atCursor',
      scopeLabel: 'Part 1',
      bookmark: 'TOC_SCOPE_part1_1',
    });

    expect(capturedCode).toContain('\\b TOC_SCOPE_part1_1');
  });
});

// ---------------------------------------------------------------------------
// A2: insertToc — F9 fallback path (A-T5)
// ---------------------------------------------------------------------------

describe('insertToc — F9 fallback path (wordApi15 absent)', () => {
  beforeEach(() => {
    resetMockCtx();
    vi.mocked(isCapable).mockReturnValue(false);
  });

  it('uses insertOoxml (not insertField) when wordApi15 is absent', async () => {
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    expect(_mockRange.insertField).not.toHaveBeenCalled();
    expect(_mockRange.insertOoxml).toHaveBeenCalled();
  });

  it('returns usedFallback:true when wordApi15 is absent', async () => {
    const result = await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    expect(result.usedFallback).toBe(true);
    expect(result.usedNativeField).toBe(false);
  });

  it('inserts a w:fldSimple element in the OOXML fragment', async () => {
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'x' });
    const ooxmlArg = _mockRange.insertOoxml.mock.calls[0]?.[0] as string | undefined;
    expect(ooxmlArg).toContain('w:fldSimple');
    expect(ooxmlArg).toContain('TOC');
  });

  it('persists a registry entry with empty fieldId on fallback path', async () => {
    await clearRegistry();
    await insertToc(HOUSE_DEFAULT_PROFILE, { placement: 'atCursor', scopeLabel: 'Fallback TOC' });
    const reg = await loadRegistry();
    expect(reg).toHaveLength(1);
    expect(reg[0]!.fieldId).toBe('');
    expect(reg[0]!.scopeLabel).toBe('Fallback TOC');
  });
});

// ---------------------------------------------------------------------------
// A3: refreshAllTocs — batched update + reconciliation
// ---------------------------------------------------------------------------

describe('refreshAllTocs — batched update and reconciliation', () => {
  beforeEach(() => {
    resetMockCtx();
    vi.mocked(isCapable).mockReturnValue(true);
  });

  it('calls getByTypes([FieldType.toc]) once', async () => {
    _mockFields.items = [makeField('f1'), makeField('f2')];
    await refreshAllTocs();
    expect(_mockCtx.document.fields.getByTypes).toHaveBeenCalledTimes(1);
  });

  it('calls updateResult() on every field (batched — no per-item sync)', async () => {
    const f1 = makeField('f1');
    const f2 = makeField('f2');
    _mockFields.items = [f1, f2];
    await refreshAllTocs();
    expect(f1.updateResult).toHaveBeenCalled();
    expect(f2.updateResult).toHaveBeenCalled();
  });

  it('prunes registry entries whose fieldId no longer exists', async () => {
    // Seed registry with f1 (still live) and f-gone (deleted).
    _docSettingsStore.set('toc.registry', [
      { fieldId: 'f1', scopeLabel: 'Live TOC' },
      { fieldId: 'f-gone', scopeLabel: 'Deleted TOC' },
    ]);
    _mockFields.items = [makeField('f1')];  // only f1 in the document

    await refreshAllTocs();

    const reg = await loadRegistry();
    expect(reg.map(e => e.fieldId)).toEqual(['f1']);
    expect(reg.find(e => e.fieldId === 'f-gone')).toBeUndefined();
  });

  it('keeps registry entries with empty fieldId (F9 fallback entries)', async () => {
    _docSettingsStore.set('toc.registry', [
      { fieldId: '', scopeLabel: 'Fallback TOC' },
      { fieldId: 'f1', scopeLabel: 'Live TOC' },
    ]);
    _mockFields.items = [makeField('f1')];

    await refreshAllTocs();

    const reg = await loadRegistry();
    expect(reg).toHaveLength(2);
    expect(reg.find(e => e.fieldId === '')).toBeDefined();
  });

  it('does nothing when wordApi15 is absent', async () => {
    vi.mocked(isCapable).mockReturnValue(false);
    await refreshAllTocs();
    expect(_mockCtx.document.fields.getByTypes).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// A4: generateBookmarkName — collision avoidance
// ---------------------------------------------------------------------------

describe('generateBookmarkName — scoped TOC bookmark naming', () => {
  it('generates TOC_SCOPE_<slug>_1 when no collision', () => {
    expect(generateBookmarkName('Part 1 — Finance', [])).toBe('TOC_SCOPE_part_1_finance_1');
  });

  it('increments counter on collision', () => {
    const existing = ['TOC_SCOPE_part1_1'];
    expect(generateBookmarkName('Part1', existing)).toBe('TOC_SCOPE_part1_2');
  });

  it('increments counter until a unique name is found', () => {
    const existing = ['TOC_SCOPE_part1_1', 'TOC_SCOPE_part1_2', 'TOC_SCOPE_part1_3'];
    expect(generateBookmarkName('Part1', existing)).toBe('TOC_SCOPE_part1_4');
  });

  it('truncates slug to 20 characters', () => {
    const name = generateBookmarkName('A very long label that exceeds twenty characters', []);
    // slug = 'a_very_long_label_th' (20 chars), so name = TOC_SCOPE_a_very_long_label_th_1
    expect(name).toContain('TOC_SCOPE_');
    const slug = name.replace('TOC_SCOPE_', '').replace('_1', '');
    expect(slug.length).toBeLessThanOrEqual(20);
  });

  it('contains _1 suffix even when no collisions', () => {
    const name = generateBookmarkName('intro', []);
    expect(name.endsWith('_1')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// A4: insertScopedToc — appends \b to switch string
// ---------------------------------------------------------------------------

describe('insertScopedToc — scoped TOC with bookmark', () => {
  beforeEach(() => {
    resetMockCtx();
    vi.mocked(isCapable).mockReturnValue(true);
    _mockBookmarks.items = [];
  });

  it('appends \\b TOC_SCOPE_… to the field code', async () => {
    let capturedCode = '';
    _mockRange.insertField.mockImplementation(() => {
      const f = makeField();
      Object.defineProperty(f, 'code', {
        set(val: string) { capturedCode = val; },
        get() { return capturedCode; },
      });
      return f;
    });

    await insertScopedToc(HOUSE_DEFAULT_PROFILE, 'Part 1');
    expect(capturedCode).toContain('\\b TOC_SCOPE_');
  });

  it('persists the bookmark name in the registry entry', async () => {
    await clearRegistry();
    await insertScopedToc(HOUSE_DEFAULT_PROFILE, 'Part 1');
    const reg = await loadRegistry();
    expect(reg[0]!.bookmark).toMatch(/^TOC_SCOPE_/);
  });
});

// ---------------------------------------------------------------------------
// A5: insertScheduleIndexField — Strategy 1 (A-T4)
// ---------------------------------------------------------------------------

describe('insertScheduleIndexField — Strategy 1 (A-T4)', () => {
  const scheduleConfig: ScheduleIndexConfig = {
    enabled: true,
    label: 'Index of Schedules',
    scheduleStyles: ['Schedule Heading', 'Schedule Title'],
    level: 1,
    separateField: true,
    placement: 'afterMainToc',
  };

  beforeEach(() => {
    resetMockCtx();
    vi.mocked(isCapable).mockReturnValue(true);
  });

  it('sets field.code with \\t for schedule styles only (no \\o)', async () => {
    let capturedCode = '';
    _mockRange.insertField.mockImplementation(() => {
      const f = makeField();
      Object.defineProperty(f, 'code', {
        set(val: string) { capturedCode = val; },
        get() { return capturedCode; },
      });
      return f;
    });

    await insertScheduleIndexField(scheduleConfig, 'topOfDocument');
    expect(capturedCode).toContain('\\t');
    expect(capturedCode).toContain('Schedule Heading');
    expect(capturedCode).not.toContain('\\o');
  });

  it('returns usedFallback:true when wordApi15 absent (A-T5)', async () => {
    vi.mocked(isCapable).mockReturnValue(false);
    const result = await insertScheduleIndexField(scheduleConfig, 'atCursor');
    expect(result.usedFallback).toBe(true);
  });
});
