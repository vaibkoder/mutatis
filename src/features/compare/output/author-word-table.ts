/**
 * Word table redline author (ADR-0046). Spec: spec/features/08-compare.md §5.5,
 * §9.3, §9.6.
 *
 * Keeps a Word table a table. A changed pair is re-assembled from the revised
 * table's own XML — `tblPr`, `tblGrid`, each row's `tblPrEx`/`trPr`, each cell's
 * `tcPr` — and only what changed carries a revision:
 *
 *  - a matched cell's content is authored by the BODY author over that cell's
 *    own ops (ADR-0045 objects, ADR-0043 mark order and all); identical content
 *    is re-emitted verbatim;
 *  - an inserted or deleted row carries `<w:ins/>`/`<w:del/>` in its `trPr` and
 *    every paragraph in it is wholly inserted or deleted (a deleted row keeps the
 *    ORIGINAL's row and cell properties).
 *
 * The Markdown author (`author-table.ts`, ADR-0027) is not used: it rebuilds
 * cells from strings, which would discard every table and run property here.
 *
 * Self-check (ADR-0008): accepting every revision in a re-assembled table must
 * spell the revised table, cell by cell, and rejecting every one the original.
 * If not, the pair is authored as the whole original table deleted followed by
 * the whole revised table inserted — never less markup, never a lost cell.
 *
 * Pure: ids, author and date come from the injected context.
 */
import type { ChangeClass, DiffOp, ParaRecord, WordTableCell, WordTableModel, WordTableRow } from '../model.js';
import { flattenTrackedChanges } from '../preprocess/flatten.js';
import { buildNodes, parseNodes } from '../preprocess/inline-units.js';
import { attrOf, childrenOf, tagOf, type PoNode } from '@shared/ooxml/po-xml.js';

import type { AuthorContext } from './author.js';

/** What the body author lends this module (passed in, so the two do not import each other). */
export interface WordTableAuthorDeps {
  /** Author a cell's content from its ops — the body author's block output, no `<w:body>`. */
  authorBlocks: (ops: DiffOp[], original: ParaRecord[], revised: ParaRecord[]) => string;
  /** Wholly insert or delete one paragraph's runs (ADR-0045-aware). */
  markParagraph: (paraXml: string, kind: 'ins' | 'del') => string;
}

/**
 * A changed Word table pair (an op carrying `tableDiff`). Falls back to the
 * whole original deleted + the whole revised inserted when either table holds
 * a wrapper the model does not represent, or when the self-check fails.
 */
export function authorWordTableChange(
  op: DiffOp,
  orig: ParaRecord,
  rev: ParaRecord,
  ctx: AuthorContext,
  deps: WordTableAuthorDeps,
): string {
  const o = orig.wordTable;
  const r = rev.wordTable;
  if (o === undefined || r === undefined || op.tableDiff === undefined || !o.supported || !r.supported) {
    return wholeTablePair(orig, rev, ctx, deps);
  }
  // Dry run on a scratch counter, so a result that fails mints no ids.
  let scratch = 1;
  const dryCtx: AuthorContext = { author: ctx.author, date: ctx.date, nextId: () => scratch++ };
  const dry = assemble(op, o, r, dryCtx, deps);
  if (!readsAs(dry, orig.rawXml, rev.rawXml)) return wholeTablePair(orig, rev, ctx, deps);
  return assemble(op, o, r, ctx, deps, op.classes);
}

/** A whole Word table inserted or deleted: every row marked, every paragraph wholly revised. */
export function authorWordTableWhole(
  rec: ParaRecord,
  kind: 'ins' | 'del',
  ctx: AuthorContext,
  deps: WordTableAuthorDeps,
): string {
  return markWhole(rec.rawXml, kind, ctx, deps);
}

/**
 * The whole original deleted, then the whole revised inserted. Two adjacent
 * tables read as one in Word: struck rows above inserted rows, which accepts to
 * the revised table and rejects to the original.
 */
function wholeTablePair(orig: ParaRecord, rev: ParaRecord, ctx: AuthorContext, deps: WordTableAuthorDeps): string {
  return markWhole(orig.rawXml, 'del', ctx, deps) + markWhole(rev.rawXml, 'ins', ctx, deps);
}

/** Re-assemble the revised table with each row and cell slot authored. */
function assemble(
  op: DiffOp,
  o: WordTableModel,
  r: WordTableModel,
  ctx: AuthorContext,
  deps: WordTableAuthorDeps,
  classes?: readonly ChangeClass[],
): string {
  let rows = '';
  for (const slot of op.tableDiff!.rows) {
    if (slot.kind === 'insert') {
      rows += markWhole(rowXml(r.rows[slot.r]!), 'ins', ctx, deps);
      continue;
    }
    if (slot.kind === 'delete') {
      rows += markWhole(rowXml(o.rows[slot.o]!), 'del', ctx, deps);
      continue;
    }
    const orow = o.rows[slot.o]!;
    const rrow = r.rows[slot.r]!;
    let cells = '';
    for (const cell of slot.cells) {
      if (cell.kind === 'insert') {
        cells += markWholeCell(rrow.cells[cell.r]!, 'ins', ctx, deps);
        continue;
      }
      if (cell.kind === 'delete') {
        cells += markWholeCell(orow.cells[cell.o]!, 'del', ctx, deps);
        continue;
      }
      const oc = orow.cells[cell.o]!;
      const rc = rrow.cells[cell.r]!;
      const content =
        cell.ops.length === 0 ? contentXml(rc) : deps.authorBlocks(cell.ops, oc.content, rc.content);
      // The nested ops tagged their own revisions; what follows is this op's.
      if (classes !== undefined) ctx.beginOp?.(classes);
      cells += `<w:tc>${rc.tcPrXml}${endsWithParagraph(content)}</w:tc>`;
    }
    rows += `<w:tr>${rrow.tblPrExXml}${rrow.trPrXml}${cells}</w:tr>`;
  }
  return `<w:tbl>${r.tblPrXml}${gridXml(op.tableDiff!, o, r)}${rows}</w:tbl>`;
}

/**
 * The authored `tblGrid`: the revised one, unless a column was added or
 * removed — then one `gridCol` per column slot, a removed column keeping its
 * original width, so every authored row spans the grid it is drawn on. A grid
 * that does not have one column per cell is left as the revised table wrote it.
 */
function gridXml(diff: NonNullable<DiffOp['tableDiff']>, o: WordTableModel, r: WordTableModel): string {
  if (diff.columns === undefined) return r.tblGridXml;
  const cols = (grid: string): string[] => grid.match(/<w:gridCol\b[^>]*?(?:\/>|>[\s\S]*?<\/w:gridCol>)/g) ?? [];
  const oc = cols(o.tblGridXml);
  const rc = cols(r.tblGridXml);
  // The map's own widths, not the first row's: a heading merged across row 0
  // is the commonest shape a column map now runs past (task 10).
  const wo = diff.columns.filter((c) => c.kind !== 'insert').length;
  const wr = diff.columns.filter((c) => c.kind !== 'delete').length;
  if (oc.length !== wo || rc.length !== wr) return r.tblGridXml;
  const built = diff.columns.map((c) => (c.kind === 'delete' ? oc[c.o]! : rc[c.r]!)).join('');
  return `<w:tblGrid>${built}</w:tblGrid>`;
}

/** A row rebuilt from its model (the model keeps every element a supported row holds that matters). */
function rowXml(row: WordTableRow): string {
  return `<w:tr>${row.tblPrExXml}${row.trPrXml}${row.cells.map((c) => `<w:tc>${c.tcPrXml}${contentXml(c)}</w:tc>`).join('')}</w:tr>`;
}

function contentXml(cell: WordTableCell): string {
  return endsWithParagraph(cell.content.map((rec) => rec.rawXml).join(''));
}

/** A cell must end with a paragraph (a nested table may not close it, and an empty cell is invalid). */
function endsWithParagraph(content: string): string {
  return /<\/w:p>\s*$|<w:p(\s[^>]*)?\/>\s*$/.test(content) ? content : `${content}<w:p/>`;
}

/** A cell added to or removed from an existing row: `w:cellIns`/`w:cellDel` and its content wholly revised. */
function markWholeCell(cell: WordTableCell, kind: 'ins' | 'del', ctx: AuthorContext, deps: WordTableAuthorDeps): string {
  const mark = revisionMark(kind === 'ins' ? 'cellIns' : 'cellDel', ctx);
  const tcPr = cell.tcPrXml === '' ? `<w:tcPr>${mark}</w:tcPr>` : addToProps(cell.tcPrXml, 'w:tcPr', mark, 'w:tcPrChange');
  const content = markWhole(`<w:tc>${contentXml(cell)}</w:tc>`, kind, ctx, deps);
  return content.replace(/^<w:tc>/, `<w:tc>${tcPr}`);
}

/**
 * Mark every row (`trPr` `ins`/`del`) and every paragraph (runs wholly
 * revised) in a fragment, nested tables included, in document order.
 */
function markWhole(xml: string, kind: 'ins' | 'del', ctx: AuthorContext, deps: WordTableAuthorDeps): string {
  const walk = (nodes: PoNode[]): string =>
    nodes
      .map((node) => {
        const tag = tagOf(node);
        if (tag === 'w:p') return deps.markParagraph(buildNodes([node]), kind);
        if (tag === 'w:tr') return markRow(node, kind, ctx, walk);
        if (tag === undefined || childrenOf(node).length === 0) return buildNodes([node]);
        return wrapChildren(node, walk(childrenOf(node)));
      })
      .join('');
  return walk(parseNodes(xml));
}

/** A `w:tr` with the row mark added to its `trPr` (created when absent) and its cells walked. */
function markRow(tr: PoNode, kind: 'ins' | 'del', ctx: AuthorContext, walk: (nodes: PoNode[]) => string): string {
  const mark = revisionMark(kind, ctx);
  const children = childrenOf(tr);
  let out = '';
  let marked = false;
  for (const child of children) {
    const tag = tagOf(child);
    if (tag === 'w:trPr') {
      out += addToProps(buildNodes([child]), 'w:trPr', mark, 'w:trPrChange');
      marked = true;
      continue;
    }
    if (!marked && tag !== 'w:tblPrEx') {
      out += `<w:trPr>${mark}</w:trPr>`;
      marked = true;
    }
    out += tag === 'w:tblPrEx' ? buildNodes([child]) : walk([child]);
  }
  if (!marked) out += `<w:trPr>${mark}</w:trPr>`;
  return wrapChildren(tr, out);
}

/** Rebuild `node`'s element around new inner XML, keeping its attributes. */
function wrapChildren(node: PoNode, inner: string): string {
  const tag = tagOf(node)!;
  const shell = buildNodes([{ [tag]: [], ...(node[':@'] !== undefined ? { ':@': node[':@'] } : {}) }]);
  // `suppressEmptyNode` writes an empty element self-closed: `<w:tr …/>`.
  return shell.replace(/\/>$/, `>${inner}</${tag}>`);
}

/**
 * Put `mark` into a properties element: before its `*PrChange` child if it has
 * one (the revision marks precede it in the schema), else at its end.
 */
function addToProps(propsXml: string, tag: string, mark: string, changeTag: string): string {
  if (new RegExp(`^<${tag}(\\s[^>]*)?/>$`).test(propsXml)) return propsXml.replace(/\/>$/, `>${mark}</${tag}>`);
  const change = propsXml.indexOf(`<${changeTag}`);
  if (change >= 0) return propsXml.slice(0, change) + mark + propsXml.slice(change);
  const close = propsXml.lastIndexOf(`</${tag}>`);
  return propsXml.slice(0, close) + mark + propsXml.slice(close);
}

function escapeAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** A self-closing revision mark: `<w:ins …/>`, `<w:cellDel …/>`. */
function revisionMark(tag: string, ctx: AuthorContext): string {
  return `<w:${tag} w:id="${ctx.nextId()}" w:author="${escapeAttr(ctx.author)}" w:date="${escapeAttr(ctx.date)}"/>`;
}

// ---------------------------------------------------------------------------
// Self-check
// ---------------------------------------------------------------------------

/**
 * Accept-all of `authored` reads as `revisedXml` and reject-all as
 * `originalXml`: the same rows, the same cells, the same words in each cell
 * (whitespace aside — a paragraph join is not a lost word) and the same objects
 * that carry no text.
 */
function readsAs(authored: string, originalXml: string, revisedXml: string): boolean {
  return (
    signature(flattenTrackedChanges(authored, 'acceptAll')) === signature(revisedXml) &&
    signature(flattenTrackedChanges(authored, 'rejectAll')) === signature(originalXml)
  );
}

const OBJECT_TAGS = new Map([
  ['w:footnoteReference', '#fn'],
  ['w:endnoteReference', '#en'],
  ['w:drawing', '#dr'],
  ['w:pict', '#pi'],
  ['m:oMath', '#eq'],
  ['w:hyperlink', '#hl'],
  ['w:fldSimple', '#fs'],
]);

/** Tables, rows and cells (nesting included), each cell's merges, its text with whitespace removed, and its objects. */
export function signature(xml: string): string {
  const out: string[] = [];
  const walk = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      // A property change's record of the OLD properties is not the table.
      if (tag.endsWith('PrChange')) continue;
      if (tag === 'w:tbl') out.push('T');
      else if (tag === 'w:tr') out.push('/');
      else if (tag === 'w:tc') out.push('|');
      else if (tag === 'w:fldChar') out.push('#fc');
      // A cell's merges are its shape: two rows with the same words in
      // differently merged cells are different tables.
      else if (tag === 'w:gridSpan' && (attrOf(node, '@_w:val') ?? '1') !== '1') out.push(`~s${attrOf(node, '@_w:val')}`);
      else if (tag === 'w:vMerge') out.push(attrOf(node, '@_w:val') === 'restart' ? '~vr' : '~vc');
      const object = OBJECT_TAGS.get(tag);
      if (object !== undefined) out.push(object);
      if (tag === 'w:t' || tag === 'm:t') {
        for (const child of childrenOf(node)) {
          const text = (child as Record<string, unknown>)['#text'];
          if (typeof text === 'string') out.push(text.replace(/\s+/g, ''));
        }
        continue;
      }
      walk(childrenOf(node));
    }
  };
  walk(parseNodes(xml));
  return out.join('');
}
