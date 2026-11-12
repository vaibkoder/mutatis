/**
 * Table-structure diff — OOXML (Word) front-end.
 * Spec: spec/features/08-compare.md §9.3, §9.6.
 *
 * Parses `<w:tbl>` structures into the format-agnostic {@link RowModel} grid and
 * delegates the row-then-cell LCS alignment to {@link ./tables-core.ts}. A genuinely
 * added column surfaces as `structural` insert ops for the new cells, not a
 * wholesale rewrite. Pure — author/date injected.
 *
 * The alignment core is shared with the PDF approximate table-region path
 *; this file owns only OOXML parsing + the multi-table loop.
 */
import { XMLParser } from 'fast-xml-parser';

import type { DiffOp } from '../model.js';

import { fnv1a } from '../preprocess/hash.js';
import { normaliseText } from '../preprocess/normalise.js';

import type { CellModel, RowModel } from './tables-core.js';
import { diffGrid, rowOp, resolveGridOptions } from './tables-core.js';

const PARSER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  suppressEmptyNode: true,
} as const;

type PoNode = Record<string, unknown>;

/**
 * Diff the tables of two document parts. Tables are aligned by position in
 * document order; within each aligned pair, rows then cells are LCS-aligned by
 * the shared grid core.
 */
export function diffTables(
  origDocXml: string,
  revDocXml: string,
  author: string,
  date: string,
): DiffOp[] {
  const origTables = extractTables(origDocXml);
  const revTables = extractTables(revDocXml);

  const ops: DiffOp[] = [];
  const counter = { next: 0 };
  const opts = resolveGridOptions({ author, date, counter });
  const pairCount = Math.min(origTables.length, revTables.length);

  for (let t = 0; t < pairCount; t++) {
    ops.push(...diffGrid(origTables[t]!, revTables[t]!, { author, date, counter }));
  }

  // Whole tables added / removed beyond the aligned pairs → structural row ops.
  for (let t = pairCount; t < origTables.length; t++) {
    const rows = origTables[t]!;
    for (let r = 0; r < rows.length; r++) {
      ops.push(rowOp('delete', r, rows[r]!, opts));
    }
  }
  for (let t = pairCount; t < revTables.length; t++) {
    const rows = revTables[t]!;
    for (let r = 0; r < rows.length; r++) {
      ops.push(rowOp('insert', r, rows[r]!, opts));
    }
  }

  return ops;
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

/** Extract all tables (in document order) as RowModel[][]. */
function extractTables(documentXml: string): RowModel[][] {
  const parser = new XMLParser(PARSER_OPTS);
  const tree = parser.parse(documentXml) as PoNode[];

  const tables: RowModel[][] = [];
  const walk = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:tbl') {
        tables.push(parseTable(node));
        continue; // nested tables handled within parseTable's row walk if present.
      }
      const value = node[tag];
      if (Array.isArray(value)) walk(value as PoNode[]);
    }
  };
  walk(tree);
  return tables;
}

/** Parse a <w:tbl> node into rows of cells with hashes. */
function parseTable(tblNode: PoNode): RowModel[] {
  const rows: RowModel[] = [];
  const tblChildren = tblNode['w:tbl'];
  if (!Array.isArray(tblChildren)) return rows;

  for (const child of tblChildren as PoNode[]) {
    if (tagOf(child) !== 'w:tr') continue;
    const cells = parseRow(child);
    const rowHash = fnv1a(cells.map((c) => c.hash).join('|'));
    rows.push({ cells, hash: rowHash });
  }
  return rows;
}

/** Parse a <w:tr> node into its cells. */
function parseRow(trNode: PoNode): CellModel[] {
  const cells: CellModel[] = [];
  const trChildren = trNode['w:tr'];
  if (!Array.isArray(trChildren)) return cells;

  for (const child of trChildren as PoNode[]) {
    if (tagOf(child) !== 'w:tc') continue;
    const text = normaliseText(cellText(child));
    cells.push({ text, hash: fnv1a(text) });
  }
  return cells;
}

/** Concatenate all <w:t> text inside a <w:tc> cell node. */
function cellText(tcNode: PoNode): string {
  const parts: string[] = [];
  const collect = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:t') {
        const kids = node['w:t'];
        if (Array.isArray(kids)) {
          for (const k of kids as PoNode[]) {
            const t = (k as Record<string, unknown>)['#text'];
            if (typeof t === 'string') parts.push(t);
          }
        }
        continue;
      }
      if (tag === 'w:tab' || tag === 'w:br' || tag === 'w:cr') {
        parts.push(' ');
        continue;
      }
      const value = node[tag];
      if (Array.isArray(value)) collect(value as PoNode[]);
    }
  };
  const value = tcNode['w:tc'];
  if (Array.isArray(value)) collect(value as PoNode[]);
  return parts.join('');
}

/** The element tag of a preserveOrder node (single non-attrs key), or undefined. */
function tagOf(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@') continue;
    return key;
  }
  return undefined;
}
