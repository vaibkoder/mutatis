/**
 * Group a body part's flat paragraph records into Word table units (ADR-0046).
 * Spec: spec/features/08-compare.md §4.1, §9.3, §9.6.
 *
 * `extractParaRecords` walks every `<w:p>` in document order, cells included,
 * and `assignNumberLabels` runs its counters over that flat sequence. This
 * post-pass runs AFTER both, so every record — and every numbering label — is
 * byte-identical to before; it only regroups. Each top-level `<w:tbl>` becomes
 * ONE record that keeps the table's XML (`rawXml`), its structure
 * (`wordTable`) and its cell paragraphs, so an equal table is re-emitted
 * verbatim and a changed one is revised inside its cells.
 *
 * The walk below mirrors `extractParaRecords`' traversal exactly — a `<w:p>` is
 * consumed and never entered, every other element is entered in child order —
 * and consumes the flat records in that same order. If the two ever disagree
 * on how many paragraphs there are, the flat records are returned unchanged:
 * the document is then authored exactly as before ADR-0046, never mis-grouped.
 */
import type { ParaRecord, WordTableCell, WordTableModel, WordTableRow } from '../model.js';
import type { RowModel } from '../diff/tables-core.js';
import { attrOf, childrenOf, parsePreserveOrder, tagOf, type PoNode } from '@shared/ooxml/po-xml.js';

import { fnv1a } from './hash.js';
import { normaliseTextAggressive } from './normalise.js';
import { buildNodes } from './inline-units.js';

/** Block-level elements a cell may hold that the model leaves out harmlessly. */
const IGNORABLE_IN_CELL = new Set(['w:tcPr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:proofErr', 'w:permStart', 'w:permEnd']);
/** Row children the model represents (or may leave out harmlessly). */
const ROW_CHILDREN = new Set(['w:tblPrEx', 'w:trPr', 'w:tc', 'w:bookmarkStart', 'w:bookmarkEnd']);
/** Table children the model represents (or may leave out harmlessly). */
const TABLE_CHILDREN = new Set(['w:tblPr', 'w:tblGrid', 'w:tr', 'w:bookmarkStart', 'w:bookmarkEnd']);

/** The separator a table's readable text uses between rows (the Markdown path's, ADR-0027). */
const ROW_SEPARATOR = '  ¶  ';

/**
 * Regroup `flat` (the records of ONE body part, in document order) so every
 * top-level table is a single unit. `cleanXml` is the part XML the records were
 * extracted from. Indices are renumbered: top-level by position, each cell's
 * content by position in that cell.
 */
export function groupTableUnits(flat: ParaRecord[], cleanXml: string): ParaRecord[] {
  if (!/<w:tbl[\s>]/.test(cleanXml)) return flat;
  const tree = parsePreserveOrder(cleanXml);
  let cursor = 0;
  let consistent = true;

  // Cell content renumbering, applied only once the whole walk is known to be
  // consistent: a walk that gives up must hand back the records untouched, or
  // the flat path reads their ranges from the wrong indices.
  const renumber: (() => void)[] = [];

  const nextRecord = (): ParaRecord | undefined => {
    const rec = flat[cursor++];
    if (rec === undefined) consistent = false;
    return rec;
  };

  /** Walk `nodes` exactly as `extractParaRecords` does, appending what it finds to `sink`. */
  const walk = (nodes: PoNode[], sink: ParaRecord[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:p') {
        const rec = nextRecord();
        if (rec !== undefined) sink.push(rec);
        continue;
      }
      if (tag === 'w:tbl') {
        if (strayParagraphs(node)) walk(childrenOf(node), sink);
        else sink.push(buildUnit(node));
        continue;
      }
      walk(childrenOf(node), sink);
    }
  };

  // A wrapper the model does not represent — a repeating-section content
  // control round `<w:tr>`, custom XML round `<w:tc>` — is TRANSPARENT: the rows
  // and cells inside it are read like any others, so their words are in the
  // model and the hash and an edit inside one is reported (ADR-0008: skipping
  // them made such an edit report nothing). The table is `supported: false`,
  // so a changed one is authored whole from its XML, wrapper and all. (A table
  // with a paragraph outside any cell never gets here — see `strayParagraphs`.)
  const buildUnit = (tbl: PoNode): ParaRecord => {
    let supported = true;
    let tblPrXml = '';
    let tblGridXml = '';
    const rows: WordTableRow[] = [];
    const visit = (nodes: PoNode[], wrapped: boolean): void => {
      for (const child of nodes) {
        const tag = tagOf(child);
        if (tag === undefined) continue;
        if (tag === 'w:tblPr' && !wrapped) tblPrXml = buildNodes([child]);
        else if (tag === 'w:tblGrid' && !wrapped) tblGridXml = buildNodes([child]);
        else if (tag === 'w:tr') {
          const row = buildRow(child);
          if (!row.supported) supported = false;
          rows.push(row.row);
        } else if (wrapped || !TABLE_CHILDREN.has(tag)) {
          supported = false;
          visit(childrenOf(child), true);
        }
      }
    };
    visit(childrenOf(tbl), false);
    const wordTable: WordTableModel = { tblPrXml, tblGridXml, rows, supported };
    return unitRecord(buildNodes([tbl]), wordTable, flat[0]!);
  };

  const buildRow = (tr: PoNode): { row: WordTableRow; supported: boolean } => {
    let supported = true;
    let tblPrExXml = '';
    let trPrXml = '';
    const cells: WordTableCell[] = [];
    const visit = (nodes: PoNode[], wrapped: boolean): void => {
      for (const child of nodes) {
        const tag = tagOf(child);
        if (tag === undefined) continue;
        if (tag === 'w:tblPrEx' && !wrapped) tblPrExXml = buildNodes([child]);
        else if (tag === 'w:trPr' && !wrapped) trPrXml = buildNodes([child]);
        else if (tag === 'w:tc') {
          const cell = buildCell(child);
          if (!cell.supported) supported = false;
          cells.push(cell.cell);
        } else if (wrapped || !ROW_CHILDREN.has(tag)) {
          supported = false;
          visit(childrenOf(child), true);
        }
      }
    };
    visit(childrenOf(tr), false);
    const hash = fnv1a(cells.map((c) => `${c.gridSpan}:${c.vMerge ?? ''}:${c.hash}`).join('|'));
    return { row: { tblPrExXml, trPrXml, cells, hash }, supported };
  };

  const buildCell = (tc: PoNode): { cell: WordTableCell; supported: boolean } => {
    let supported = true;
    let tcPrXml = '';
    let gridSpan = 1;
    let vMerge: WordTableCell['vMerge'] = null;
    const content: ParaRecord[] = [];
    for (const child of childrenOf(tc)) {
      const tag = tagOf(child);
      if (tag === undefined) continue;
      if (tag === 'w:p' || tag === 'w:tbl') {
        walk([child], content);
        continue;
      }
      if (tag === 'w:tcPr') {
        tcPrXml = buildNodes([child]);
        for (const prop of childrenOf(child)) {
          const ptag = tagOf(prop);
          if (ptag === 'w:gridSpan') gridSpan = Number(attrOf(prop, '@_w:val') ?? '1') || 1;
          if (ptag === 'w:vMerge') vMerge = attrOf(prop, '@_w:val') === 'restart' ? 'restart' : 'continue';
        }
        continue;
      }
      if (!IGNORABLE_IN_CELL.has(tag)) {
        // A block-level wrapper (a content control, custom XML) round a cell's
        // paragraphs: its paragraphs are still read, so the text is compared,
        // but a changed table is not re-assembled from a model that drops it.
        supported = false;
        walk(childrenOf(child), content);
      }
    }
    renumber.push(() =>
      content.forEach((rec, i) => {
        rec.index = i;
      }),
    );
    const hash = fnv1a(content.map((r) => (r.wordTable ? 'T' : 'P') + r.hash).join('|'));
    return { cell: { tcPrXml, gridSpan, vMerge, content, hash }, supported };
  };

  const grouped: ParaRecord[] = [];
  walk(tree, grouped);
  if (!consistent || cursor !== flat.length) return flat;
  for (const apply of renumber) apply();
  grouped.forEach((rec, i) => {
    rec.index = i;
  });
  return grouped;
}

/**
 * The flat paragraph sequence a grouped record list stands for: every table
 * unit replaced by its cells' records, row by row, nested tables included.
 * The inverse of {@link groupTableUnits} as far as text goes — what a reader
 * reads, in reading order.
 */
export function flattenTableUnits(records: readonly ParaRecord[]): ParaRecord[] {
  const out: ParaRecord[] = [];
  for (const rec of records) {
    if (rec.wordTable === undefined) {
      out.push(rec);
      continue;
    }
    for (const row of rec.wordTable.rows) {
      for (const cell of row.cells) out.push(...flattenTableUnits(cell.content));
    }
  }
  return out;
}

/** A cell's readable text: its content records' text, one space apart. */
export function cellText(cell: WordTableCell): string {
  return cell.content
    .map((r) => r.normText)
    .filter((t) => t !== '')
    .join(' ');
}

/** True when any record in `records` (nested tables included) is a signature-block paragraph. */
function anySignature(records: readonly ParaRecord[]): boolean {
  return records.some((r) => r.isSignatureBlock);
}

/**
 * Whether anything a reader reads sits anywhere but in a cell of a row: a
 * paragraph or table outside every cell, or a cell outside every row — reached
 * through a wrapper, or written where the schema allows neither. Word does not
 * write it, but a generator can, and a unit has no place for it: its style,
 * objects and position would be in no hash, and the walk would not consume it.
 * Such a table is left as flat records and compared paragraph by paragraph, as
 * before ADR-0046 — every change reported, the table's structure not kept.
 * Wrappers and properties are searched too: a paragraph inside `w:tblPr` is
 * stray like any other.
 */
function strayParagraphs(tbl: PoNode): boolean {
  const search = (nodes: PoNode[], inRow: boolean): boolean =>
    nodes.some((node) => {
      const tag = tagOf(node);
      if (tag === undefined) return false;
      if (tag === 'w:tc') return !inRow;
      if (tag === 'w:p' || tag === 'w:tbl') return true;
      return search(childrenOf(node), inRow || tag === 'w:tr');
    });
  return search(childrenOf(tbl), false);
}

/** Build the unit record for one table, in the part `like` (any record of it) belongs to. */
function unitRecord(rawXml: string, wordTable: WordTableModel, like: ParaRecord): ParaRecord {
  const grid: RowModel[] = wordTable.rows.map((row) => ({
    cells: row.cells.map((c) => ({ text: cellText(c), hash: c.hash })),
    hash: row.hash,
  }));
  const normText = grid.map((r) => r.cells.map((c) => c.text).join(' | ')).join(ROW_SEPARATOR);
  const cells = wordTable.rows.flatMap((r) => r.cells);
  return {
    index: 0,
    partId: like.partId,
    rawXml,
    normText,
    normTextAggressive: normaliseTextAggressive(normText),
    numberingPrefix: '',
    hash: fnv1a(`tbl|${wordTable.rows.map((r) => r.hash).join('/')}`),
    styleId: null,
    numbering: null,
    isTableCell: true,
    isSignatureBlock: cells.some((c) => anySignature(c.content)),
    isHeaderFooter: like.isHeaderFooter,
    isFootnote: like.isFootnote,
    table: grid,
    wordTable,
  };
}
