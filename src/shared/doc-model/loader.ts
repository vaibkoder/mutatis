import { wordRun } from '@shared/office/batch.js';
import type { ParaNode, DocSnapshot } from './types.js';

/**
 * Single-pass document loader.
 * Reads all paragraphs and footnotes in one batched Word.run, building the
 * in-memory DocSnapshot that every feature queries.
 * (spec/ARCHITECTURE.md §Document Model / Paragraph Index)
 */
export async function loadDocSnapshot(): Promise<DocSnapshot> {
  return wordRun(async ctx => {
    const body = ctx.document.body;
    const paras = body.paragraphs;
    paras.load('items');

    // WordApi 1.5: load footnote bodies
    let footnoteParas: Word.Paragraph[] = [];
    let hasFootnotes = false;
    try {
      const footnotes = body.footnotes;
      footnotes.load('items');
      await ctx.sync();
      hasFootnotes = true;
      footnoteParas = footnotes.items.flatMap(fn => {
        fn.body.paragraphs.load('items');
        return [];
      });
      if (hasFootnotes) {
        // reload after loading paragraph items inside each footnote body
        for (const fn of footnotes.items) {
          fn.body.paragraphs.load('items');
        }
        await ctx.sync();
        footnoteParas = footnotes.items.flatMap(fn => fn.body.paragraphs.items);
      }
    } catch {
      // Footnotes not supported on this host (pre-1.5); degrade gracefully.
    }

    // Load all paragraph properties. Split into two syncs so that a failure
    // loading the optional parentTableCellOrNullObject properties (not supported
    // on all Word Online hosts) never prevents text from being read.
    await ctx.sync();
    const rawParas = paras.items;

    // Essential properties — must succeed. `listItem` is deliberately NOT loaded
    // here: it is a navigation object and loading it directly (rather than its
    // sub-properties) throws on Word Online/Safari, which would prevent text from
    // ever being read. List info is loaded separately and guarded below.
    const allParas = [...rawParas, ...footnoteParas];
    for (const p of allParas) {
      p.load('text,style');
    }
    await ctx.sync();

    // Optional table-cell coordinates — degrade gracefully if unsupported.
    let cellCoordsLoaded = false;
    try {
      for (const p of rawParas) {
        p.load('parentTableCellOrNullObject/rowIndex,parentTableCellOrNullObject/cellIndex');
      }
      await ctx.sync();
      cellCoordsLoaded = true;
    } catch {
      // parentTableCellOrNullObject not available on this host; all paragraphs
      // will be tagged as 'body' container.
    }

    // Optional list info — load via the null-object variant so a non-list
    // paragraph (or a host without list support) degrades gracefully instead of
    // throwing. Mirrors the parentTableCellOrNullObject block above.
    let listInfoLoaded = false;
    try {
      for (const p of allParas) {
        p.listItemOrNullObject.load('listString,level');
      }
      await ctx.sync();
      listInfoLoaded = true;
    } catch {
      // List items not available on this host; paragraphs carry no list info.
    }

    const buildNode = (
      p: Word.Paragraph,
      index: number,
      container: ParaNode['container'],
      tableInfo?: ParaNode['tableInfo'],
    ): ParaNode => {
      const id = `${container === 'tableCell' ? 'body' : container}-para-${index}`;
      let listString: string | undefined;
      let listLevel: number | undefined;

      if (listInfoLoaded) {
        try {
          const li = p.listItemOrNullObject;
          if (li && !li.isNullObject) {
            listString = li.listString ?? undefined;
            listLevel = li.level;
          }
        } catch { /* not a list item */ }
      }

      const capturedPara = p;

      const node: ParaNode = {
        id,
        index,
        text: p.text,
        styleName: p.style,
        container,
        getRange: (_ctx: Word.RequestContext) => capturedPara.getRange(),
      };
      if (listString !== undefined) node.listString = listString;
      if (listLevel !== undefined) node.listLevel = listLevel;
      if (tableInfo !== undefined) node.tableInfo = tableInfo;
      return node;
    };

    // Read a paragraph's table-cell coordinates, or null if it is body prose or
    // if the optional cell-coordinates load was skipped/failed.
    const cellCoords = (p: Word.Paragraph): { row: number; col: number } | null => {
      if (!cellCoordsLoaded) return null;
      try {
        const cell = p.parentTableCellOrNullObject;
        if (!cell || cell.isNullObject) return null;
        return { row: cell.rowIndex, col: cell.cellIndex };
      } catch {
        return null;
      }
    };

    // Walk body paragraphs in document order, tagging table-cell paragraphs and
    // synthesising a stable tableId. body.paragraphs returns a table's cell
    // paragraphs contiguously and in (row, col) order, so a new table begins
    // whenever a cell paragraph's (row, col) resets lexicographically below the
    // previous cell paragraph (or follows a non-table paragraph).
    let tableSeq = -1;
    let prev: { row: number; col: number } | null = null;
    const paragraphs: ParaNode[] = rawParas.map((p, i) => {
      const coords = cellCoords(p);
      if (!coords) {
        prev = null;
        return buildNode(p, i, 'body');
      }
      const isNewTable =
        prev === null || coords.row < prev.row ||
        (coords.row === prev.row && coords.col < prev.col);
      if (isNewTable) tableSeq += 1;
      prev = coords;
      return buildNode(p, i, 'tableCell', {
        tableId: `table-${tableSeq}`,
        row: coords.row,
        col: coords.col,
      });
    });
    const footnotes: ParaNode[] = footnoteParas.map((p, i) => buildNode(p, i, 'footnote'));

    return { paragraphs, footnotes, loadedAt: Date.now() };
  });
}
