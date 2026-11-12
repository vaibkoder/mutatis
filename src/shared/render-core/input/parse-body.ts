/**
 * Parse `authoredBodyXml` (and footnote/cell block lists) into RenderCore's
 * resolved block model — host-free, pure. Handles native revision wrappers
 * (`<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>`), runs, footnote references, image
 * placeholders, and basic tables. Reuses the shared `po-xml` preserveOrder
 * helpers and the style cascade.
 */
import { type PoNode, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';
import type {
  RcBlock, RcParagraph, RcRun, RcRow, RcCell, RcTable, EffectiveRPr, RcRevision, RcRevisionType,
} from '../types.js';
import { type StyleTable, resolveParagraph, applyRPr } from './styles.js';

const TWIP = 20;

export interface ParseCtx {
  styles: StyleTable;
  /** Numbering references collected in document order (paraId assigned by sequence). */
  numberingRefs: Array<{ para: RcParagraph; numId: string; ilvl: number; paraId: string }>;
  /** Footnote ids referenced by the body, in encounter order. */
  footnoteRefs: string[];
  seq: { n: number };
}

export function newCtx(styles: StyleTable): ParseCtx {
  return { styles, numberingRefs: [], footnoteRefs: [], seq: { n: 0 } };
}

const REV_TAGS: Record<string, RcRevisionType> = {
  'w:ins': 'ins',
  'w:del': 'del',
  'w:moveFrom': 'moveFrom',
  'w:moveTo': 'moveTo',
};

function revisionOf(node: PoNode, type: RcRevisionType): RcRevision {
  const rev: RcRevision = { type };
  const id = attrOf(node, '@_w:id');
  const author = attrOf(node, '@_w:author');
  const date = attrOf(node, '@_w:date');
  if (id !== undefined) rev.changeId = id;
  if (author !== undefined) rev.author = author;
  if (date !== undefined) rev.date = date;
  return rev;
}

/** Text content of a `<w:t>`/`<w:delText>` node (its #text children, spaces preserved). */
function textOf(node: PoNode): string {
  let s = '';
  for (const c of childrenOf(node)) {
    const txt = (c as Record<string, unknown>)['#text'];
    if (typeof txt === 'string') s += txt;
    else if (typeof txt === 'number') s += String(txt);
  }
  return s;
}

/** Build runs from a `<w:r>` node, given the inherited base rPr + revision context. */
function runsFromR(r: PoNode, baseRpr: EffectiveRPr, rev: RcRevision | undefined): RcRun[] {
  const kids = childrenOf(r);
  let rprNodes: PoNode[] | undefined;
  for (const k of kids) if (tagOf(k) === 'w:rPr') rprNodes = childrenOf(k);
  const rpr = applyRPr(baseRpr, rprNodes);

  const out: RcRun[] = [];
  const pushText = (text: string): void => {
    if (text.length === 0) return;
    const run: RcRun = { text, rpr };
    if (rev) run.revision = rev;
    out.push(run);
  };

  for (const k of kids) {
    const t = tagOf(k);
    switch (t) {
      case 'w:t':
      case 'w:delText':
      case 'w:instrText':
        if (t !== 'w:instrText') pushText(textOf(k)); // skip field instruction text
        break;
      case 'w:tab': pushText('\t'); break;
      case 'w:br': pushText('\n'); break;
      case 'w:cr': pushText('\n'); break;
      case 'w:noBreakHyphen': pushText('‑'); break;
      case 'w:footnoteReference': {
        const id = attrOf(k, '@_w:id') ?? '';
        const run: RcRun = { text: '', rpr: { ...rpr, superscript: true }, footnoteId: id };
        if (rev) run.revision = rev;
        out.push(run);
        break;
      }
      case 'w:drawing':
      case 'w:pict':
      case 'w:object': {
        const run: RcRun = { text: '￼', rpr, imagePlaceholder: true };
        if (rev) run.revision = rev;
        out.push(run);
        break;
      }
      default:
        break;
    }
  }
  return out;
}

/** Walk a run-container child list (paragraph or revision wrapper), appending runs. */
function collectRuns(
  nodes: PoNode[],
  baseRpr: EffectiveRPr,
  rev: RcRevision | undefined,
  out: RcRun[],
): void {
  for (const node of nodes) {
    const t = tagOf(node);
    if (t === 'w:r') {
      out.push(...runsFromR(node, baseRpr, rev));
    } else if (t && REV_TAGS[t]) {
      collectRuns(childrenOf(node), baseRpr, revisionOf(node, REV_TAGS[t]), out);
    } else if (t === 'w:hyperlink' || t === 'w:smartTag' || t === 'w:sdt') {
      // transparent containers — recurse, preserving revision context
      const inner = t === 'w:sdt' ? sdtContent(node) : childrenOf(node);
      collectRuns(inner, baseRpr, rev, out);
    }
  }
}

function sdtContent(sdt: PoNode): PoNode[] {
  for (const c of childrenOf(sdt)) if (tagOf(c) === 'w:sdtContent') return childrenOf(c);
  return [];
}

function paragraph(p: PoNode, ctx: ParseCtx): RcParagraph {
  const kids = childrenOf(p);
  let pPrNodes: PoNode[] = [];
  for (const k of kids) if (tagOf(k) === 'w:pPr') pPrNodes = childrenOf(k);

  // style + paragraph-mark rPr + numPr from pPr
  let styleId: string | null = null;
  let markRPrNodes: PoNode[] = [];
  let numId: string | undefined;
  let ilvl = 0;
  for (const n of pPrNodes) {
    const t = tagOf(n);
    if (t === 'w:pStyle') styleId = attrOf(n, '@_w:val') ?? null;
    else if (t === 'w:rPr') markRPrNodes = childrenOf(n);
    else if (t === 'w:numPr') {
      for (const np of childrenOf(n)) {
        if (tagOf(np) === 'w:numId') numId = attrOf(np, '@_w:val');
        else if (tagOf(np) === 'w:ilvl') ilvl = Number(attrOf(np, '@_w:val') ?? '0') || 0;
      }
    }
  }

  const { ppr, baseRpr, isHeading } = resolveParagraph(ctx.styles, styleId, pPrNodes, markRPrNodes);

  const runs: RcRun[] = [];
  collectRuns(kids, baseRpr, undefined, runs);
  for (const r of runs) if (r.footnoteId !== undefined) ctx.footnoteRefs.push(r.footnoteId);

  const para: RcParagraph = { runs, ppr, markRpr: baseRpr, styleId, isHeading };
  if (numId !== undefined && numId !== '0') {
    const paraId = `p${ctx.seq.n++}`;
    ctx.numberingRefs.push({ para, numId, ilvl, paraId });
  }
  return para;
}

function table(tbl: PoNode, ctx: ParseCtx): RcTable {
  const gridColsPt: number[] = [];
  const rows: RcRow[] = [];
  for (const node of childrenOf(tbl)) {
    const t = tagOf(node);
    if (t === 'w:tblGrid') {
      for (const gc of childrenOf(node)) {
        if (tagOf(gc) === 'w:gridCol') {
          const w = Number(attrOf(gc, '@_w:w') ?? '0');
          gridColsPt.push((Number.isFinite(w) ? w : 0) / TWIP);
        }
      }
    } else if (t === 'w:tr') {
      rows.push(row(node, ctx));
    }
  }
  return { rows, gridColsPt };
}

function row(tr: PoNode, ctx: ParseCtx): RcRow {
  let cantSplit = false;
  const cells: RcCell[] = [];
  for (const node of childrenOf(tr)) {
    const t = tagOf(node);
    if (t === 'w:trPr') {
      for (const p of childrenOf(node)) if (tagOf(p) === 'w:cantSplit') cantSplit = true;
    } else if (t === 'w:tc') {
      cells.push(cell(node, ctx));
    }
  }
  return { cells, cantSplit };
}

function cell(tc: PoNode, ctx: ParseCtx): RcCell {
  let gridSpan = 1;
  const blockNodes: PoNode[] = [];
  for (const node of childrenOf(tc)) {
    const t = tagOf(node);
    if (t === 'w:tcPr') {
      for (const p of childrenOf(node)) {
        if (tagOf(p) === 'w:gridSpan') gridSpan = Number(attrOf(p, '@_w:val') ?? '1') || 1;
      }
    } else if (t === 'w:p' || t === 'w:tbl') {
      blockNodes.push(node);
    }
  }
  return { blocks: parseBlocks(blockNodes, ctx), gridSpan };
}

/** Parse a list of block-level nodes (body, table cell, or footnote body). */
export function parseBlocks(nodes: PoNode[], ctx: ParseCtx): RcBlock[] {
  const out: RcBlock[] = [];
  for (const node of nodes) {
    const t = tagOf(node);
    if (t === 'w:p') out.push({ kind: 'paragraph', paragraph: paragraph(node, ctx) });
    else if (t === 'w:tbl') out.push({ kind: 'table', table: table(node, ctx) });
  }
  return out;
}

/** Find the trailing body-level `<w:sectPr>` child list, if any. */
export function findSectPr(bodyChildren: PoNode[]): PoNode[] | undefined {
  for (let i = bodyChildren.length - 1; i >= 0; i--) {
    const n = bodyChildren[i];
    if (n && tagOf(n) === 'w:sectPr') return childrenOf(n);
  }
  return undefined;
}
