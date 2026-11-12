/**
 * tests/unit/toc.schedule-index.test.ts
 *
 * Unit tests for A5: schedule-index.ts (Index of Schedules, Strategies 1 and 2).
 *
 * Acceptance check A-T4: Strategy 1 field lists exactly the Schedule-styled
 * paragraphs and excludes body headings (verified by checking the switch
 * string forwarded to insertScheduleIndexField).
 *
 * Strategy 2 acceptance: snapshot HTML contains the schedule text and is
 * labelled correctly; body heading paragraphs are excluded.
 *
 * HOST-ONLY CAVEATS:
 *  - Actual TOC field rendering for Strategy 1 — needs live Word host.
 *  - Bookmark insertion and HTML rendering for Strategy 2 — needs live host.
 *  - range.insertHtml output fidelity — needs live host.
 */

import { describe, it, expect, vi, beforeEach, afterAll } from 'vitest';

// ---------------------------------------------------------------------------
// Global Word stub
// ---------------------------------------------------------------------------

function installWordStub() {
  (globalThis as unknown as { Word: unknown }).Word = {
    InsertLocation: { before: 'Before', after: 'After', replace: 'Replace', start: 'Start', end: 'End' },
    FieldType: { toc: 'Toc' },
  };
}

installWordStub();

afterAll(() => {
  delete (globalThis as unknown as { Word?: unknown }).Word;
});

// ---------------------------------------------------------------------------
// Mock: @shared/office/index.js
// ---------------------------------------------------------------------------

let _isCapableResult = true;

// Capture the HTML string passed to insertHtml in Strategy 2.
const _insertHtmlCalls: string[] = [];
const _insertBookmarkCalls: string[] = [];

// Stateful document model for Strategy-2 idempotency (Regenerate) tests:
// content controls and bookmarks survive across Word.run calls so a second
// generate sees what the first one inserted.
interface MockContentControl { tag: string; title: string; deleted: boolean; delete: (keepContent: boolean) => void }
const _docContentControls: MockContentControl[] = [];
const _docBookmarks: Array<{ name: string }> = [];

const _mockRange = {
  insertField: vi.fn(() => ({
    id: 'f-sched-001',
    code: '',
    result: { text: 'Schedule 1\t1\nSchedule 2\t2' },
    updateResult: vi.fn(),
    load: vi.fn(),
  })),
  insertOoxml: vi.fn(),
  // insertHtml returns the inserted range so callers can wrap it in a CC.
  insertHtml: vi.fn((html: string) => { _insertHtmlCalls.push(html); return _mockRange; }),
  insertBookmark: vi.fn((name: string) => {
    _insertBookmarkCalls.push(name);
    _docBookmarks.push({ name });
  }),
  insertContentControl: vi.fn(() => {
    const cc: MockContentControl = {
      tag: '',
      title: '',
      deleted: false,
      delete(_keepContent: boolean) { this.deleted = true; },
    };
    _docContentControls.push(cc);
    return cc;
  }),
};

const _mockCtx = {
  document: {
    body: { getRange: vi.fn(() => _mockRange) },
    getSelection: vi.fn(() => ({ getRange: vi.fn(() => _mockRange) })),
    fields: {
      getByTypes: vi.fn(() => ({ items: [], load: vi.fn() })),
    },
    contentControls: {
      getByTag: vi.fn((tag: string) => ({
        items: _docContentControls.filter(c => c.tag === tag && !c.deleted),
        load: vi.fn(),
      })),
    },
    bookmarks: { items: _docBookmarks, load: vi.fn() },
  },
  sync: vi.fn(async () => {}),
};

vi.mock('@shared/office/index.js', () => ({
  isCapable: vi.fn(() => _isCapableResult),
  wordRun: vi.fn(async (fn: (ctx: unknown) => Promise<unknown>) => fn(_mockCtx)),
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

const _store = new Map<string, unknown>();

vi.mock('@shared/storage/index.js', () => ({
  getStorageProvider: vi.fn(() => ({
    docSettings: {
      get: vi.fn(async (k: string) => _store.get(k)),
      set: vi.fn(async (k: string, v: unknown) => { _store.set(k, v); }),
      delete: vi.fn(async (k: string) => { _store.delete(k); }),
      list: vi.fn(async () => []),
      export: vi.fn(async () => new Blob(['{}'])),
      import: vi.fn(async () => {}),
    },
    local: { get: vi.fn(), set: vi.fn(), delete: vi.fn(), list: vi.fn(async () => []) },
  })),
}));

// ---------------------------------------------------------------------------
// Mock: @shared/doc-model/index.js
// ---------------------------------------------------------------------------

const _mockParagraphs: Array<{ id: string; text: string; styleName: string; getRange: (ctx: unknown) => typeof _mockRange }> = [];

vi.mock('@shared/doc-model/index.js', () => ({
  DocModel: {
    getParagraphsByStyle: vi.fn((styleNames: string[]) => {
      const wanted = new Set(styleNames.map((s: string) => s.toLowerCase()));
      return _mockParagraphs.filter(p => wanted.has(p.styleName.toLowerCase()));
    }),
    paragraphs: { value: [] },
    allParagraphs: { value: [] },
    footnotes: { value: [] },
    load: vi.fn(async () => {}),
  },
}));

// ---------------------------------------------------------------------------
// Import under test
// ---------------------------------------------------------------------------

import { isCapable } from '@shared/office/index.js';
import {
  generateScheduleIndexStrategy1,
  generateScheduleIndexStrategy2,
} from '@features/content/toc/schedule-index.js';
import type { ScheduleIndexConfig } from '@features/content/toc/types.js';

// ---------------------------------------------------------------------------
// Test fixtures
// ---------------------------------------------------------------------------

const SCHEDULE_CONFIG: ScheduleIndexConfig = {
  enabled: true,
  label: 'Index of Schedules',
  scheduleStyles: ['Schedule Heading', 'Schedule Title'],
  level: 1,
  separateField: true,
  placement: 'afterMainToc',
};

function makeSchedulePara(id: string, text: string, style = 'Schedule Heading') {
  return {
    id,
    text,
    styleName: style,
    getRange: (_ctx: unknown) => _mockRange,
  };
}

// ---------------------------------------------------------------------------
// Strategy 1 tests (A-T4 scope)
// ---------------------------------------------------------------------------

describe('generateScheduleIndexStrategy1 — Strategy 1 (live TOC field)', () => {
  beforeEach(() => {
    _isCapableResult = true;
    vi.mocked(isCapable).mockImplementation(() => _isCapableResult);
    _store.clear();
    _mockRange.insertField.mockClear();
    _mockRange.insertOoxml.mockClear();
    _insertHtmlCalls.length = 0;
    _insertBookmarkCalls.length = 0;
    _mockCtx.sync.mockClear();
    // Reset to valid field mock.
    _mockRange.insertField.mockReturnValue({
      id: 'f-sched-001',
      code: '',
      result: { text: 'Schedule 1\t1\nSchedule 2\t2' },
      updateResult: vi.fn(),
      load: vi.fn(),
    });
  });

  it('returns strategy:1', async () => {
    const result = await generateScheduleIndexStrategy1(SCHEDULE_CONFIG);
    expect(result.strategy).toBe(1);
  });

  it('uses native field when wordApi15 is present', async () => {
    const result = await generateScheduleIndexStrategy1(SCHEDULE_CONFIG);
    expect(result.usedNativeField).toBe(true);
    expect(result.usedFallback).toBe(false);
  });

  it('uses fallback (insertOoxml) when wordApi15 is absent (A-T5)', async () => {
    _isCapableResult = false;
    vi.mocked(isCapable).mockImplementation(() => false);
    const result = await generateScheduleIndexStrategy1(SCHEDULE_CONFIG, 'atCursor');
    expect(result.usedFallback).toBe(true);
    expect(result.usedNativeField).toBe(false);
    expect(_mockRange.insertOoxml).toHaveBeenCalled();
  });

  it('calls insertField with FieldType.toc (A-T4)', async () => {
    await generateScheduleIndexStrategy1(SCHEDULE_CONFIG);
    expect(_mockRange.insertField).toHaveBeenCalledWith(
      expect.anything(),
      'Toc',   // Word.FieldType.toc stubbed to 'Toc'
    );
  });

  it('the field code contains only \\t for schedule styles — no \\o (A-T4)', async () => {
    let capturedCode = '';
    _mockRange.insertField.mockReturnValue({
      id: 'f1',
      get code() { return capturedCode; },
      set code(v: string) { capturedCode = v; },
      result: { text: 'Schedule 1' },
      updateResult: vi.fn(),
      load: vi.fn(),
    });

    await generateScheduleIndexStrategy1(SCHEDULE_CONFIG);

    expect(capturedCode).toContain('\\t');
    expect(capturedCode).toContain('Schedule Heading');
    expect(capturedCode).toContain('Schedule Title');
    expect(capturedCode).not.toContain('\\o');
  });
});

// ---------------------------------------------------------------------------
// Strategy 2 tests
// ---------------------------------------------------------------------------

describe('generateScheduleIndexStrategy2 — Strategy 2 (snapshot)', () => {
  beforeEach(() => {
    _isCapableResult = true;
    vi.mocked(isCapable).mockImplementation(() => _isCapableResult);
    _store.clear();
    _mockRange.insertField.mockClear();
    _mockRange.insertOoxml.mockClear();
    _mockRange.insertHtml.mockClear();
    _mockRange.insertBookmark.mockClear();
    _mockRange.insertContentControl.mockClear();
    _insertHtmlCalls.length = 0;
    _insertBookmarkCalls.length = 0;
    _docContentControls.length = 0;
    _docBookmarks.length = 0;
    _mockCtx.sync.mockClear();
    _mockParagraphs.length = 0;
  });

  it('returns strategy:2', async () => {
    _mockParagraphs.push(makeSchedulePara('p1', 'Schedule 1'));
    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    expect(result.strategy).toBe(2);
  });

  it('returns isEmpty:true when no schedule-styled paragraphs exist', async () => {
    // _mockParagraphs is empty.
    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    expect(result.isEmpty).toBe(true);
    expect(result.entries).toHaveLength(0);
  });

  it('collects exactly the paragraphs matching schedule styles (A-T4)', async () => {
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1 — Definitions', 'Schedule Heading'),
      makeSchedulePara('body', 'Clause 1 — General', 'Heading 1'),   // should be excluded
      makeSchedulePara('sch2', 'Schedule 2 — IP', 'Schedule Title'),
    );

    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);

    expect(result.isEmpty).toBe(false);
    expect(result.entries).toHaveLength(2);
    expect(result.entries.map(e => e.paragraphId)).toEqual(['sch1', 'sch2']);
  });

  it('excludes body heading paragraphs (A-T4: only schedule styles)', async () => {
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1', 'Schedule Heading'),
      makeSchedulePara('h1', 'Clause 1', 'Heading 1'),
    );

    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    const ids = result.entries.map(e => e.paragraphId);
    expect(ids).toContain('sch1');
    expect(ids).not.toContain('h1');
  });

  it('inserts HTML with the label "Index of Schedules"', async () => {
    _mockParagraphs.push(makeSchedulePara('sch1', 'Schedule 1'));
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    const html = _insertHtmlCalls[0] ?? '';
    expect(html).toContain('Index of Schedules');
  });

  it('labels the HTML as a snapshot requiring regeneration', async () => {
    _mockParagraphs.push(makeSchedulePara('sch1', 'Schedule 1'));
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    const html = _insertHtmlCalls[0] ?? '';
    expect(html.toLowerCase()).toContain('snapshot');
    expect(html.toLowerCase()).toContain('regenerate');
  });

  it('HTML-escapes special characters in schedule text', async () => {
    _mockParagraphs.push(makeSchedulePara('sch1', 'Schedule 1 — <Test> & "IP"'));
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    const html = _insertHtmlCalls[0] ?? '';
    // Should NOT contain raw < > & "
    expect(html).toContain('&lt;Test&gt;');
    expect(html).toContain('&amp;');
    expect(html).toContain('&quot;');
  });

  it('inserts a bookmark for each schedule paragraph', async () => {
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1'),
      makeSchedulePara('sch2', 'Schedule 2'),
    );
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    // One bookmark per paragraph.
    expect(_mockRange.insertBookmark).toHaveBeenCalledTimes(2);
  });

  it('snapshot label comes from config (not hardcoded "Table of Authorities")', async () => {
    const customConfig: ScheduleIndexConfig = {
      ...SCHEDULE_CONFIG,
      label: 'Index of Schedules',
    };
    _mockParagraphs.push(makeSchedulePara('sch1', 'Schedule 1'));
    await generateScheduleIndexStrategy2(customConfig);
    const html = _insertHtmlCalls[0] ?? '';
    expect(html).toContain('Index of Schedules');
    expect(html).not.toContain('Table of Authorities');
  });

  it('wraps the snapshot in a content control tagged IDX_SCHEDULE_SNAPSHOT', async () => {
    _mockParagraphs.push(makeSchedulePara('sch1', 'Schedule 1'));
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    expect(_mockRange.insertContentControl).toHaveBeenCalledTimes(1);
    expect(_docContentControls).toHaveLength(1);
    expect(_docContentControls[0]?.tag).toBe('IDX_SCHEDULE_SNAPSHOT');
  });

  it('does not reuse a bookmark name already present in the document (collision check, §A.8)', async () => {
    // Pre-seed a user bookmark that the legacy fixed-name scheme would clobber.
    _docBookmarks.push({ name: 'IDX_SCHEDULE_1' });
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1'),
      makeSchedulePara('sch2', 'Schedule 2'),
    );

    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    const assigned = result.entries.map(e => e.bookmark);

    expect(assigned).not.toContain('IDX_SCHEDULE_1'); // user's bookmark untouched
    expect(new Set(assigned).size).toBe(assigned.length); // all unique
  });

  it('Regenerate after all Schedules removed clears the prior snapshot (no stale list)', async () => {
    // First generate with two schedules.
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1'),
      makeSchedulePara('sch2', 'Schedule 2'),
    );
    await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    expect(_docContentControls.filter(c => !c.deleted)).toHaveLength(1);

    // Now the document has no Schedule-styled paragraphs; Regenerate.
    _mockParagraphs.length = 0;
    const result = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);

    expect(result.isEmpty).toBe(true);
    // The prior snapshot was removed rather than left stale.
    const liveSnapshots = _docContentControls.filter(
      c => c.tag === 'IDX_SCHEDULE_SNAPSHOT' && !c.deleted,
    );
    expect(liveSnapshots).toHaveLength(0);
  });

  it('Regenerate replaces in place: exactly ONE snapshot remains and bookmark names stay unique', async () => {
    _mockParagraphs.push(
      makeSchedulePara('sch1', 'Schedule 1 — Definitions'),
      makeSchedulePara('sch2', 'Schedule 2 — IP'),
    );

    // First generate.
    const first = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);
    // Second generate (Regenerate) — same paragraphs, same tag.
    const second = await generateScheduleIndexStrategy2(SCHEDULE_CONFIG);

    // Exactly one live (non-deleted) snapshot content control survives.
    const liveSnapshots = _docContentControls.filter(
      c => c.tag === 'IDX_SCHEDULE_SNAPSHOT' && !c.deleted,
    );
    expect(liveSnapshots).toHaveLength(1);

    // The prior snapshot was deleted, not left to stack.
    const deletedSnapshots = _docContentControls.filter(
      c => c.tag === 'IDX_SCHEDULE_SNAPSHOT' && c.deleted,
    );
    expect(deletedSnapshots).toHaveLength(1);

    // Bookmark names across both cycles are all unique (counter-suffixed).
    const allNames = [...first.entries, ...second.entries].map(e => e.bookmark);
    expect(new Set(allNames).size).toBe(allNames.length);
    // The second cycle stepped past the first cycle's names rather than colliding.
    expect(second.entries.map(e => e.bookmark)).not.toEqual(
      first.entries.map(e => e.bookmark),
    );
  });
});
