import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { DefinitionsMap, DefinitionRecord, TextAnchor } from './types.js';
import { mergeDefinition } from './definitions.js';

interface Cell { text: string; paraId: string; }
type Row = Map<number, Cell>;       // col → cell
type Table = Map<number, Row>;      // row → Row

const NUMBERING = /^\s*\d+(?:\.\d+)*\.?\s*$/;
const ENUMERATOR = /^\s*(?:\([a-z0-9ivxlcdm]+\)|\d+\.)\s*/i;
const HEADER = /^(term|defined term|definition|meaning|interpretation)s?$/i;
const QUOTED = /["'“‘]/;

/** Reconstruct tables from per-paragraph tableInfo (group by tableId → row → col). */
function reconstruct(tableParas: ReadonlyArray<ParaNode>): Map<string, Table> {
  const tables = new Map<string, Table>();
  for (const p of tableParas) {
    const ti = p.tableInfo;
    if (!ti) continue;
    let table = tables.get(ti.tableId);
    if (!table) { table = new Map(); tables.set(ti.tableId, table); }
    let row = table.get(ti.row);
    if (!row) { row = new Map(); table.set(ti.row, row); }
    const existing = row.get(ti.col);
    if (existing) {
      existing.text = `${existing.text} ${p.text}`.trim();
    } else {
      row.set(ti.col, { text: p.text, paraId: p.id });
    }
  }
  return tables;
}

interface TableShape { ok: boolean; firstContentCol: number; }

/** Decide whether a reconstructed table is a two-column definition table (spec §6.2). */
function classifyTable(table: Table): TableShape {
  const rows = [...table.keys()].sort((a, b) => a - b).map(r => table.get(r)!);
  if (rows.length < 1) return { ok: false, firstContentCol: 0 };

  // Leading pure-numbering column tolerated → content shifts right by one.
  const col0 = rows.map(r => r.get(0)?.text.trim() ?? '');
  const numberingCol = col0.every(v => v === '' || NUMBERING.test(v)) && col0.some(v => /^\d/.test(v));
  const firstContentCol = numberingCol ? 1 : 0;
  const L = firstContentCol;
  const R = firstContentCol + 1;

  let qualifying = 0;
  let anyQuoted = false;
  for (const r of rows) {
    const left = (r.get(L)?.text ?? '').trim();
    const right = (r.get(R)?.text ?? '').trim();
    if (!left && !right) continue;
    if (HEADER.test(left.replace(/["'“”‘’]/g, ''))) continue; // skip header row
    const words = left.split(/\s+/).length;
    const leftShort = left.length <= 80 && words <= 12;
    const leftTermish = QUOTED.test(left[0] ?? '') || /^[A-Z]/.test(left);
    if (leftShort && leftTermish && right.length > 0) {
      qualifying += 1;
      if (QUOTED.test(left[0] ?? '')) anyQuoted = true;
    }
  }

  // Err toward capturing: two qualifying rows, or one with a quoted headword.
  const ok = qualifying >= 2 || (qualifying >= 1 && anyQuoted);
  return { ok, firstContentCol };
}

/** Split a term cell into one or more raw terms (multi-term "and"/"," splits). */
function extractTerms(cellText: string): string[] {
  const stripped = cellText.replace(ENUMERATOR, '').trim();
  // Multiple quoted runs → one term each (e.g. "Buyer" and "Purchaser").
  const quoted = [...stripped.matchAll(/["'“‘]([^"'“”‘’]{1,120})["'”’]/g)].map(m => m[1] ?? '');
  if (quoted.length >= 1) return quoted.filter(t => /[a-z0-9]/i.test(t));
  // Unquoted single term cell.
  if (/[a-z]/i.test(stripped) && stripped.length > 0) return [stripped];
  return [];
}

/**
 * Reconstruct two-column definition tables and merge their term/definition rows
 * into the shared DefinitionsMap (source = "table", verb = null). (spec §6)
 */
export function mergeTableDefinitions(
  map: DefinitionsMap,
  tableParas: ReadonlyArray<ParaNode>,
  _profile: JurisdictionProfile,
): void {
  const tables = reconstruct(tableParas);
  for (const table of tables.values()) {
    const shape = classifyTable(table);
    if (!shape.ok) continue;
    const L = shape.firstContentCol;
    const R = L + 1;

    for (const row of table.values()) {
      const termCell = row.get(L);
      const defnCell = row.get(R);
      if (!termCell || !defnCell) continue;
      const left = termCell.text.trim();
      if (!left || HEADER.test(left.replace(/["'“”‘’]/g, '')) || !defnCell.text.trim()) continue;

      for (const raw of extractTerms(termCell.text)) {
        const startInCell = Math.max(0, termCell.text.indexOf(raw));
        const anchor: TextAnchor = {
          paragraphId: termCell.paraId,
          startOffset: startInCell,
          endOffset: startInCell + raw.length,
        };
        const record: DefinitionRecord = {
          source: 'table',
          verb: null,
          shape: 'table',
          isCrossRef: false,
          isCrossDoc: false,
          crossRefTarget: null,
          anchor,
          styling: { bold: false, italic: false },
          rawQuote: QUOTED.test(termCell.text.trim()[0] ?? '') ? 'double' : 'none',
          contextScope: null,
        };
        mergeDefinition(map, raw, record);
      }
    }
  }
}
