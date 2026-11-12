import { describe, it, expect, afterEach } from 'vitest';
import { loadDocSnapshot } from '@shared/doc-model/loader.js';

// ---------------------------------------------------------------------------
// Minimal Word mock — just enough surface for the single-pass loader.
// Each spec is [text, cell] where cell is null for body prose or {row, col}
// for a table-cell paragraph (parentTableCellOrNullObject coordinates).
// ---------------------------------------------------------------------------

interface ParaSpec {
  text: string;
  cell: { row: number; col: number } | null;
}

function mockParagraph(spec: ParaSpec) {
  return {
    text: spec.text,
    style: 'Normal',
    listItem: null,
    parentTableCellOrNullObject: spec.cell
      ? { isNullObject: false, rowIndex: spec.cell.row, cellIndex: spec.cell.col }
      : { isNullObject: true, rowIndex: undefined, cellIndex: undefined },
    load() { /* properties already present on the mock */ },
    getRange() { return {} as unknown; },
  };
}

function installWordMock(specs: ParaSpec[]) {
  const paragraphItems = specs.map(mockParagraph);
  const ctx = {
    document: {
      body: {
        paragraphs: { load() {}, items: paragraphItems },
        footnotes: { load() {}, items: [] as unknown[] },
      },
    },
    sync: async () => {},
  };
  (globalThis as unknown as { Word: unknown }).Word = {
    run: async (fn: (c: typeof ctx) => unknown) => fn(ctx),
  };
}

afterEach(() => {
  delete (globalThis as unknown as { Word?: unknown }).Word;
});

describe('loadDocSnapshot — table-cell tagging', () => {
  it('tags body prose as body and table-cell paragraphs with coordinates', async () => {
    installWordMock([
      { text: 'Introductory prose.', cell: null },
      { text: '"Term One"', cell: { row: 0, col: 0 } },
      { text: 'means the first thing.', cell: { row: 0, col: 1 } },
    ]);

    const snap = await loadDocSnapshot();
    const [intro, term, defn] = snap.paragraphs;

    expect(intro?.container).toBe('body');
    expect(intro?.tableInfo).toBeUndefined();

    expect(term?.container).toBe('tableCell');
    expect(term?.tableInfo).toEqual({ tableId: 'table-0', row: 0, col: 0 });
    expect(defn?.tableInfo).toEqual({ tableId: 'table-0', row: 0, col: 1 });
  });

  it('keeps body-para-{index} ids unchanged for table-cell paragraphs', async () => {
    installWordMock([
      { text: 'Prose.', cell: null },
      { text: 'Cell.', cell: { row: 0, col: 0 } },
    ]);
    const snap = await loadDocSnapshot();
    expect(snap.paragraphs[0]?.id).toBe('body-para-0');
    expect(snap.paragraphs[1]?.id).toBe('body-para-1');
  });

  it('groups contiguous cells under one tableId and starts a new id for a second table', async () => {
    installWordMock([
      { text: '"A"', cell: { row: 0, col: 0 } },
      { text: 'def A', cell: { row: 0, col: 1 } },
      { text: '"B"', cell: { row: 1, col: 0 } },
      { text: 'def B', cell: { row: 1, col: 1 } },
      { text: 'Separating prose.', cell: null },
      { text: '"C"', cell: { row: 0, col: 0 } },
      { text: 'def C', cell: { row: 0, col: 1 } },
    ]);

    const snap = await loadDocSnapshot();
    const ids = snap.paragraphs.map(p => p.tableInfo?.tableId ?? null);
    expect(ids).toEqual([
      'table-0', 'table-0', 'table-0', 'table-0',
      null,
      'table-1', 'table-1',
    ]);
  });

  it('starts a new table when two tables are adjacent (row/col resets)', async () => {
    installWordMock([
      { text: '"A"', cell: { row: 0, col: 0 } },
      { text: 'def A', cell: { row: 1, col: 0 } },
      // second table immediately follows, row index resets to 0
      { text: '"B"', cell: { row: 0, col: 0 } },
      { text: 'def B', cell: { row: 0, col: 1 } },
    ]);
    const snap = await loadDocSnapshot();
    const ids = snap.paragraphs.map(p => p.tableInfo?.tableId);
    expect(ids).toEqual(['table-0', 'table-0', 'table-1', 'table-1']);
  });
});
