import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { attrOf, childrenOf, parsePreserveOrder, tagOf, type PoNode } from '@shared/ooxml/po-xml.js';
import {
  applyLvlText,
  formatInt,
  parseNumberingSchemeXml,
  parseStyleNumberingSource,
  renderListNumbers,
  resolveLevel,
  resolveParaNumPr,
  type NumberingScheme,
  type ParaNumPr,
  type StyleNumberingSource,
} from '@shared/numbering/index.js';

/**
 * Read a .docx into the paragraph list the contract-structure analyser works on.
 *
 * ONE entry per body `<w:p>` in document order, table cells and content controls
 * included, deleted paragraphs included. That is the "markup" view: a paragraph
 * whose text was wholly struck still has a row, with an empty `text` and its old
 * words in `textOriginal`. Collapsing to one view first — which is what Compare's
 * flattener does, correctly, for Compare — would lose the clause a reviewer is
 * most likely to ask about.
 *
 * Numbers are what Word PAINTS, from the engine's own counter machine
 * (`@shared/numbering`). Nothing here re-implements a counter; what this module
 * adds is deciding which paragraphs take part in the count: a paragraph absent
 * from the accepted view consumes no number, nor does an empty paragraph that
 * only carries a section break, nor a vertical-merge continuation cell.
 */

export interface TableCoord {
  readonly tableId: number;
  readonly row: number;
  readonly col: number;
  readonly parentTableId: number | null;
  /** Columns in this row — a one-cell row is a full-width title or a box. */
  readonly rowCells: number;
}

export interface DocxPara {
  readonly index: number;
  /** Accepted view (every insertion kept, every deletion gone). */
  readonly text: string;
  /** Reject-all view. */
  readonly textOriginal: string;
  readonly inAccepted: boolean;
  readonly inOriginal: boolean;
  readonly hasRevisions: boolean;
  readonly styleId: string | null;
  readonly styleName: string | null;
  /** 0-based outline level from the paragraph or its style chain. */
  readonly outlineLevel: number | null;
  readonly isHeadingStyle: boolean;
  readonly isToc: boolean;
  readonly isTitleStyle: boolean;
  /** Word's painted number in the accepted view ("4.2", "(b)"), or null. */
  readonly label: string | null;
  readonly labelOriginal: string | null;
  /** The list level of `label` (0-based), when numbered. */
  readonly listLevel: number | null;
  /** `${numId}:${ilvl}` — which list and level painted the number. */
  readonly listKey: string | null;
  /** The abstract definition behind the list, so restarted instances group. */
  readonly listGroup: string | null;
  /** Numbered by a bullet level. A level that paints nothing is not a bullet: a typed number may follow. */
  readonly bullet: boolean;
  readonly numFmt: string | null;
  readonly table: TableCoord | null;
  readonly vMergeContinue: boolean;
  /** The paragraph carries a section break (`w:pPr/w:sectPr`). */
  readonly sectionBreak: boolean;
  /** A page break before or inside the paragraph. */
  readonly pageBreak: boolean;
  /** Every accepted-text character is bold. */
  readonly allBold: boolean;
  /** The leading bold run(s), when the paragraph starts bold and goes on in plain text. */
  readonly boldLead: string;
  readonly allCaps: boolean;
  readonly underlineLead: string;
  readonly centred: boolean;
  /** Left indent in twips, from the paragraph or its style. */
  readonly indent: number;
}

export interface DocxRead {
  readonly paragraphs: DocxPara[];
  readonly revisionsPresent: boolean;
}

interface StyleInfo {
  readonly name: string | null;
  readonly basedOn: string | null;
  readonly outline: number | null;
  readonly bold: boolean | null;
  readonly caps: boolean | null;
  readonly indent: number | null;
  readonly jc: string | null;
  readonly type: string | null;
}

interface Styles {
  readonly byId: Map<string, StyleInfo>;
  readonly defaultPara: string | null;
  readonly numbering: StyleNumberingSource;
}

const MAX_CHAIN = 24;

function boolProp(node: PoNode | undefined): boolean | null {
  if (!node) return null;
  const v = attrOf(node, '@_w:val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

function child(nodes: readonly PoNode[], tag: string): PoNode | undefined {
  return nodes.find((n) => tagOf(n) === tag);
}

function readStyles(xml: string | undefined): Styles {
  const byId = new Map<string, StyleInfo>();
  let defaultPara: string | null = null;
  const root = xml ? parsePreserveOrder(xml).find((n) => tagOf(n) === 'w:styles') : undefined;
  for (const node of root ? childrenOf(root) : []) {
    if (tagOf(node) !== 'w:style') continue;
    const id = attrOf(node, '@_w:styleId');
    if (id === undefined) continue;
    const kids = childrenOf(node);
    const pPr = childrenOf(child(kids, 'w:pPr') ?? {});
    const rPr = childrenOf(child(kids, 'w:rPr') ?? {});
    const outlineNode = child(pPr, 'w:outlineLvl');
    const outlineVal = outlineNode ? Number(attrOf(outlineNode, '@_w:val')) : NaN;
    const ind = child(pPr, 'w:ind');
    const indVal = ind ? Number(attrOf(ind, '@_w:left') ?? attrOf(ind, '@_w:start')) : NaN;
    const jc = child(pPr, 'w:jc');
    const type = attrOf(node, '@_w:type') ?? null;
    byId.set(id, {
      name: attrOf(child(kids, 'w:name') ?? {}, '@_w:val') ?? null,
      basedOn: attrOf(child(kids, 'w:basedOn') ?? {}, '@_w:val') ?? null,
      outline: Number.isFinite(outlineVal) && outlineVal < 9 ? outlineVal : null,
      bold: boolProp(child(rPr, 'w:b')),
      caps: boolProp(child(rPr, 'w:caps')),
      indent: Number.isFinite(indVal) ? indVal : null,
      jc: jc ? (attrOf(jc, '@_w:val') ?? null) : null,
      type,
    });
    if (defaultPara === null && type === 'paragraph') {
      const flag = attrOf(node, '@_w:default');
      if (flag !== undefined && flag !== '0' && flag !== 'false' && flag !== 'off') defaultPara = id;
    }
  }
  return { byId, defaultPara, numbering: parseStyleNumberingSource(xml) };
}

/** Walk a style chain leaf to root and return the first defined value. */
function inherit<T>(styles: Styles, styleId: string | null, pick: (s: StyleInfo) => T | null): T | null {
  let id = styleId ?? styles.defaultPara;
  for (let guard = 0; id && guard < MAX_CHAIN; guard += 1) {
    const s = styles.byId.get(id);
    if (!s) return null;
    const v = pick(s);
    if (v !== null) return v;
    id = s.basedOn;
  }
  return null;
}

interface Run {
  readonly text: string;
  readonly bold: boolean;
  readonly underline: boolean;
  readonly caps: boolean;
}

interface Collected {
  accepted: Run[];
  original: string[];
  revisions: boolean;
  pageBreak: boolean;
}

/** Inline containers whose children are runs (or more containers). */
const TRANSPARENT = new Set([
  'w:hyperlink',
  'w:smartTag',
  'w:fldSimple',
  'w:customXml',
  'w:bdo',
  'w:dir',
  'w:sdtContent',
  'w:moveFromRangeStart',
]);

function collectRuns(
  nodes: readonly PoNode[],
  styles: Styles,
  paraStyle: string | null,
  state: { del: boolean; ins: boolean },
  out: Collected,
): void {
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue;
    switch (tag) {
      case 'w:r':
        collectRun(node, styles, paraStyle, state, out);
        break;
      case 'w:ins':
      case 'w:moveTo':
        out.revisions = true;
        collectRuns(childrenOf(node), styles, paraStyle, { ...state, ins: true }, out);
        break;
      case 'w:del':
      case 'w:moveFrom':
        out.revisions = true;
        collectRuns(childrenOf(node), styles, paraStyle, { ...state, del: true }, out);
        break;
      case 'w:sdt': {
        const content = child(childrenOf(node), 'w:sdtContent');
        if (content) collectRuns(childrenOf(content), styles, paraStyle, state, out);
        break;
      }
      case 'mc:AlternateContent': {
        const kids = childrenOf(node);
        const pick = child(kids, 'mc:Choice') ?? child(kids, 'mc:Fallback');
        if (pick) collectRuns(childrenOf(pick), styles, paraStyle, state, out);
        break;
      }
      default:
        if (TRANSPARENT.has(tag)) collectRuns(childrenOf(node), styles, paraStyle, state, out);
    }
  }
}

function textOfNode(node: PoNode): string {
  let s = '';
  for (const kid of childrenOf(node)) {
    const t = (kid as Record<string, unknown>)['#text'];
    if (typeof t === 'string') s += t;
  }
  return s;
}

function collectRun(
  run: PoNode,
  styles: Styles,
  paraStyle: string | null,
  state: { del: boolean; ins: boolean },
  out: Collected,
): void {
  const kids = childrenOf(run);
  const rPr = childrenOf(child(kids, 'w:rPr') ?? {});
  const rStyle = attrOf(child(rPr, 'w:rStyle') ?? {}, '@_w:val') ?? null;
  const bold =
    boolProp(child(rPr, 'w:b')) ??
    (rStyle ? inherit(styles, rStyle, (s) => s.bold) : null) ??
    inherit(styles, paraStyle, (s) => s.bold) ??
    false;
  const caps =
    boolProp(child(rPr, 'w:caps')) ??
    (rStyle ? inherit(styles, rStyle, (s) => s.caps) : null) ??
    inherit(styles, paraStyle, (s) => s.caps) ??
    false;
  const u = child(rPr, 'w:u');
  const underline = u !== undefined && attrOf(u, '@_w:val') !== 'none';
  if (child(rPr, 'w:ins') || child(rPr, 'w:del')) out.revisions = true;

  const push = (text: string): void => {
    if (text === '') return;
    if (!state.del) out.accepted.push({ text, bold, underline, caps });
    if (!state.ins) out.original.push(text);
  };

  let inInstr = false;
  for (const node of kids) {
    const tag = tagOf(node);
    switch (tag) {
      case 'w:t':
        if (!inInstr) push(textOfNode(node));
        break;
      case 'w:delText':
        if (!state.ins && !inInstr) out.original.push(textOfNode(node));
        break;
      case 'w:tab':
      case 'w:ptab':
        push('\t');
        break;
      case 'w:br':
        // A page break separates pages, not words: Word's own text has nothing there.
        if (attrOf(node, '@_w:type') === 'page') out.pageBreak = true;
        else push('\n');
        break;
      case 'w:cr':
        push('\n');
        break;
      case 'w:lastRenderedPageBreak':
        break;
      case 'w:noBreakHyphen':
        push('-');
        break;
      case 'w:sym':
        push('￼');
        break;
      case 'w:fldChar':
        // Field codes are instrText, which is never read; the result is plain runs.
        inInstr = false;
        break;
      // Text boxes are their own story, not words of the paragraph that anchors them.
      default:
        break;
    }
  }
}

interface RawPara {
  node: PoNode;
  table: TableCoord | null;
  vMergeContinue: boolean;
}

/** Body paragraphs in document order, descending into tables and content controls. */
function bodyParagraphs(body: readonly PoNode[]): RawPara[] {
  const out: RawPara[] = [];
  let tableSeq = 0;
  const walk = (
    nodes: readonly PoNode[],
    cell: { table: TableCoord; vMergeContinue: boolean } | null,
  ): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === 'w:p') {
        out.push({ node, table: cell?.table ?? null, vMergeContinue: cell?.vMergeContinue ?? false });
      } else if (tag === 'w:tbl') {
        const tableId = tableSeq++;
        let row = 0;
        for (const tr of childrenOf(node)) {
          if (tagOf(tr) !== 'w:tr') continue;
          const cells = rowCells(childrenOf(tr));
          let col = 0;
          for (const tc of cells) {
            const tcPr = childrenOf(child(childrenOf(tc), 'w:tcPr') ?? {});
            const vMerge = child(tcPr, 'w:vMerge');
            const cont = vMerge !== undefined && attrOf(vMerge, '@_w:val') !== 'restart';
            walk(childrenOf(tc), {
              table: {
                tableId,
                row,
                col,
                parentTableId: cell?.table.tableId ?? null,
                rowCells: cells.length,
              },
              vMergeContinue: cont,
            });
            col += 1;
          }
          row += 1;
        }
      } else if (tag === 'w:sdt') {
        const content = child(childrenOf(node), 'w:sdtContent');
        if (content) walk(childrenOf(content), cell);
      } else if (
        tag === 'w:customXml' ||
        tag === 'w:ins' ||
        tag === 'w:del' ||
        tag === 'w:moveTo' ||
        tag === 'w:moveFrom'
      ) {
        walk(childrenOf(node), cell);
      } else if (tag === 'mc:AlternateContent') {
        const kids = childrenOf(node);
        const pick = child(kids, 'mc:Choice') ?? child(kids, 'mc:Fallback');
        if (pick) walk(childrenOf(pick), cell);
      }
    }
  };
  walk(body, null);
  return out;
}

/** The cells of a row, looking through row-level content controls. */
function rowCells(nodes: readonly PoNode[]): PoNode[] {
  const cells: PoNode[] = [];
  for (const n of nodes) {
    const tag = tagOf(n);
    if (tag === 'w:tc') cells.push(n);
    else if (tag === 'w:sdt') {
      const content = child(childrenOf(n), 'w:sdtContent');
      if (content) cells.push(...rowCells(childrenOf(content)));
    } else if (tag === 'w:customXml') cells.push(...rowCells(childrenOf(n)));
  }
  return cells;
}

function joinRuns(runs: readonly Run[]): string {
  return runs.map((r) => r.text).join('');
}

/** The leading run text that satisfies `pick`, provided plain text follows it. */
function leadWhere(runs: readonly Run[], pick: (r: Run) => boolean): string {
  let lead = '';
  let i = 0;
  for (; i < runs.length; i += 1) {
    const r = runs[i]!;
    if (pick(r)) lead += r.text;
    else if (r.text.trim() === '' && lead === '') continue;
    else break;
  }
  return lead.trim() === '' ? '' : lead;
}

export function readDocx(bytes: Uint8Array | ArrayBuffer): DocxRead {
  const bundle = new PizZipBundle(bytes);
  const documentXml = bundle.part('word/document.xml');
  if (documentXml === undefined) throw new Error('Not a Word document: word/document.xml is missing');
  const styles = readStyles(bundle.part('word/styles.xml'));
  const scheme = parseNumberingSchemeXml(bundle.part('word/numbering.xml'));

  const docRoot = parsePreserveOrder(documentXml).find((n) => tagOf(n) === 'w:document');
  const bodyNode = docRoot ? child(childrenOf(docRoot), 'w:body') : undefined;
  const raws = bodyParagraphs(bodyNode ? childrenOf(bodyNode) : []);

  interface Draft {
    text: string;
    textOriginal: string;
    inAccepted: boolean;
    inOriginal: boolean;
    revisions: boolean;
    styleId: string | null;
    numAccepted: { numId: string; ilvl: number } | null;
    numOriginal: { numId: string; ilvl: number } | null;
    counts: boolean;
    countsOriginal: boolean;
    raw: RawPara;
    runs: Run[];
    sectionBreak: boolean;
    pageBreak: boolean;
    outline: number | null;
    indent: number;
    jc: string | null;
  }

  let revisionsPresent = false;
  const drafts: Draft[] = raws.map((raw) => {
    const kids = childrenOf(raw.node);
    const pPrNode = child(kids, 'w:pPr');
    const pPr = childrenOf(pPrNode ?? {});
    const styleId = attrOf(child(pPr, 'w:pStyle') ?? {}, '@_w:val') ?? null;
    const markRPr = childrenOf(child(pPr, 'w:rPr') ?? {});
    const markDel = child(markRPr, 'w:del') !== undefined || child(markRPr, 'w:moveFrom') !== undefined;
    const markIns = child(markRPr, 'w:ins') !== undefined || child(markRPr, 'w:moveTo') !== undefined;
    const pPrChange = child(pPr, 'w:pPrChange');
    const oldPPr = pPrChange ? childrenOf(child(childrenOf(pPrChange), 'w:pPr') ?? {}) : null;

    const out: Collected = { accepted: [], original: [], revisions: false, pageBreak: false };
    collectRuns(kids.filter((k) => tagOf(k) !== 'w:pPr'), styles, styleId, { del: false, ins: false }, out);
    const text = joinRuns(out.accepted);
    const textOriginal = out.original.join('');
    const revisions = out.revisions || markDel || markIns || pPrChange !== undefined;
    if (revisions) revisionsPresent = true;

    // A paragraph whose mark was deleted and whose words are all gone has left the
    // accepted view; one whose mark was inserted is not in the original.
    const inAccepted = !(markDel && text.trim() === '');
    const inOriginal = !(markIns && textOriginal.trim() === '');
    const sectionBreak = child(pPr, 'w:sectPr') !== undefined;
    const blankBoth = text.trim() === '' && textOriginal.trim() === '';
    // A deleted paragraph mark joins this paragraph to the next one: it paints no number of its own.
    const counts = inAccepted && !markDel && !raw.vMergeContinue && !(sectionBreak && blankBoth);
    const countsOriginal = inOriginal && !raw.vMergeContinue && !(sectionBreak && blankBoth);

    const numAccepted = resolveParaNumPr(styles.numbering, styleId, pPr);
    let numOriginal = numAccepted;
    if (oldPPr !== null) {
      const oldStyle = attrOf(child(oldPPr, 'w:pStyle') ?? {}, '@_w:val') ?? styleId;
      numOriginal = resolveParaNumPr(styles.numbering, oldStyle, oldPPr);
    }

    const outlineNode = child(pPr, 'w:outlineLvl');
    const ownOutline = outlineNode ? Number(attrOf(outlineNode, '@_w:val')) : NaN;
    const outline = Number.isFinite(ownOutline) && ownOutline < 9 ? ownOutline : inherit(styles, styleId, (s) => s.outline);
    const ind = child(pPr, 'w:ind');
    const ownIndent = ind ? Number(attrOf(ind, '@_w:left') ?? attrOf(ind, '@_w:start')) : NaN;
    const indent = Number.isFinite(ownIndent) ? ownIndent : (inherit(styles, styleId, (s) => s.indent) ?? 0);
    const jcNode = child(pPr, 'w:jc');
    const jc = jcNode ? (attrOf(jcNode, '@_w:val') ?? null) : inherit(styles, styleId, (s) => s.jc);
    const pageBefore = boolProp(child(pPr, 'w:pageBreakBefore')) === true;

    return {
      text,
      textOriginal,
      inAccepted,
      inOriginal,
      revisions,
      styleId,
      numAccepted,
      numOriginal,
      counts,
      countsOriginal,
      raw,
      runs: out.accepted,
      sectionBreak,
      pageBreak: out.pageBreak || pageBefore,
      outline,
      indent,
      jc,
    };
  });

  const accepted = paint(
    scheme,
    drafts.map((d, i) => (d.counts && d.numAccepted ? { paraId: String(i), ...d.numAccepted } : null)),
  );
  const original = paint(
    scheme,
    drafts.map((d, i) => (d.countsOriginal && d.numOriginal ? { paraId: String(i), ...d.numOriginal } : null)),
  );

  const paragraphs: DocxPara[] = drafts.map((d, index) => {
    const styleName = d.styleId ? (styles.byId.get(d.styleId)?.name ?? null) : null;
    const lowerName = (styleName ?? '').toLowerCase();
    // The list a paragraph belongs to: its accepted numbering, or for a deleted one, its original.
    const num = d.counts ? d.numAccepted : d.countsOriginal && !d.inAccepted ? d.numOriginal : null;
    const level = num ? resolveLevel(scheme, num.numId, num.ilvl) : undefined;
    const painted = accepted.get(String(index)) ?? null;
    const bullet = level !== undefined && level.numFmt === 'bullet';
    const textRuns = d.runs.filter((r) => r.text.trim() !== '');
    return {
      index,
      text: d.text,
      textOriginal: d.textOriginal,
      inAccepted: d.inAccepted,
      inOriginal: d.inOriginal,
      hasRevisions: d.revisions,
      styleId: d.styleId,
      styleName,
      outlineLevel: d.outline,
      isHeadingStyle: /^heading [1-9]$/.test(lowerName) || (d.outline !== null && d.outline < 9 && lowerName !== ''),
      isToc: /^toc( |\d|$)|^table of contents/.test(lowerName),
      isTitleStyle: lowerName === 'title' || lowerName === 'subtitle',
      label: d.counts && level && level.numFmt !== 'bullet' ? painted : null,
      labelOriginal: original.get(String(index)) ?? null,
      listLevel: num ? num.ilvl : null,
      listKey: num ? `${num.numId}:${num.ilvl}` : null,
      listGroup: num ? (scheme.nums.get(num.numId)?.abstractNumId ?? num.numId) : null,
      bullet,
      numFmt: level?.numFmt ?? null,
      table: d.raw.table,
      vMergeContinue: d.raw.vMergeContinue,
      sectionBreak: d.sectionBreak,
      pageBreak: d.pageBreak,
      allBold: textRuns.length > 0 && textRuns.every((r) => r.bold),
      boldLead: leadWhere(d.runs, (r) => r.bold),
      allCaps: textRuns.length > 0 && textRuns.every((r) => r.caps),
      underlineLead: leadWhere(d.runs, (r) => r.underline),
      centred: d.jc === 'center',
      indent: d.indent,
    };
  });

  return { paragraphs, revisionsPresent };
}

/**
 * Paint labels for the paragraphs that take part in the count. The engine's
 * counter machine supplies the counters; the TOKENS are re-rendered here where
 * the engine is known to differ from Word: under `<w:isLgl/>` Word turns the
 * ancestor levels decimal but keeps the paragraph's own format ("Section 1.01",
 * not "1.1"), and the text formats ("First", "Second") render nothing upstream.
 */
function paint(schemeIn: NumberingScheme, refs: readonly (ParaNumPr | null)[]): Map<string, string> {
  const present = refs.filter((r): r is ParaNumPr => r !== null);
  const scheme = withoutHiddenRestarts(schemeIn, present);
  // Word keeps one counter per abstract definition AS REFERENCED: a list that borrows its levels
  // through `numStyleLink` still counts on its own. The engine merges the two, so each group is
  // run through the counter machine separately.
  const groups = new Map<string, ParaNumPr[]>();
  for (const r of present) {
    const num = scheme.nums.get(r.numId);
    const own = num?.abstractNumId ?? r.numId;
    (groups.get(own) ?? groups.set(own, []).get(own)!).push(r);
  }
  const painted = new Map<string, { listString: string; path: readonly number[] }>();
  for (const group of groups.values()) for (const [id, n] of renderListNumbers(scheme, group)) painted.set(id, n);
  const out = new Map<string, string>();
  const byId = new Map(present.map((r) => [r.paraId, r]));
  for (const [id, n] of painted) {
    let label = n.listString.trim();
    const ref = byId.get(id);
    const level = ref ? resolveLevel(scheme, ref.numId, ref.ilvl) : undefined;
    if (ref && level && n.path.length === ref.ilvl + 1 && (level.isLgl || needsOwnFormat(level.numFmt) || n.path.some((c) => c > 26))) {
      const tokens = n.path.map((counter, k) => {
        if (k === ref.ilvl) return formatToken(counter, level.numFmt);
        if (level.isLgl) return String(counter);
        return formatToken(counter, resolveLevel(scheme, ref.numId, k)?.numFmt ?? 'decimal');
      });
      label = applyLvlText(level.lvlText, tokens).trim();
    }
    if (label !== '') out.set(id, label);
  }
  return out;
}

/**
 * Drop a `startOverride` on an ANCESTOR level that none of the instance's paragraphs displays.
 * The engine restarts every level an instance touches, ancestors included, so one "(c)" restarted
 * at level 3 knocked the NIH licence's clause 6 back to 2; Word leaves the hidden levels alone.
 */
function withoutHiddenRestarts(scheme: NumberingScheme, refs: readonly ParaNumPr[]): NumberingScheme {
  const drop = new Map<string, Set<number> | null>(); // null: some paragraph shows a restarted ancestor
  const seen = new Set<string>();
  for (let k = 0; k < refs.length; k += 1) {
    const r = refs[k]!;
    const num = scheme.nums.get(r.numId);
    if (!num || drop.get(r.numId) === null) continue;
    const first = !seen.has(r.numId);
    seen.add(r.numId);
    if (!first) continue;
    // only a restart that would wind back a counter the rest of the list is using
    const ancestors = [...num.lvlOverrides]
      .filter(([l, o]) => o.startOverride !== undefined && l < r.ilvl && countBefore(scheme, refs, k, l) > o.startOverride)
      .map(([l]) => l);
    if (ancestors.length === 0) continue;
    const text = resolveLevel(scheme, r.numId, r.ilvl)?.lvlText ?? '';
    if (ancestors.some((l) => text.includes(`%${l + 1}`))) {
      drop.set(r.numId, null);
      continue;
    }
    const set = drop.get(r.numId) ?? new Set<number>();
    for (const l of ancestors) set.add(l);
    drop.set(r.numId, set);
  }
  if (![...drop.values()].some((v) => v !== null && v.size > 0)) return scheme;
  const nums = new Map(scheme.nums);
  for (const [numId, levels] of drop) {
    const num = nums.get(numId);
    if (!num || !levels) continue;
    const lvlOverrides = new Map(num.lvlOverrides);
    for (const l of levels) {
      const o = lvlOverrides.get(l)!;
      if (o.lvl) lvlOverrides.set(l, { lvl: o.lvl });
      else lvlOverrides.delete(l);
    }
    nums.set(numId, { ...num, lvlOverrides });
  }
  return { ...scheme, nums };
}

/** Paragraphs of the same list at level `l` before `refs[k]`, since the last one above `l`. */
function countBefore(scheme: NumberingScheme, refs: readonly ParaNumPr[], k: number, l: number): number {
  const group = scheme.nums.get(refs[k]!.numId)?.abstractNumId;
  let n = 0;
  for (let j = k - 1; j >= 0; j -= 1) {
    const q = refs[j]!;
    if (scheme.nums.get(q.numId)?.abstractNumId !== group) continue;
    if (q.ilvl < l) break;
    if (q.ilvl === l) n += 1;
  }
  return n;
}

function needsOwnFormat(numFmt: string): boolean {
  return numFmt !== 'bullet' && numFmt !== 'none' && formatInt(1, numFmt) === '' && formatToken(1, numFmt) !== '';
}

const ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve',
  'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
const ORDINAL_IRREGULAR: Record<string, string> = {
  One: 'First', Two: 'Second', Three: 'Third', Five: 'Fifth', Eight: 'Eighth', Nine: 'Ninth', Twelve: 'Twelfth',
};

function cardinalWords(n: number): string {
  if (n < 20) return ONES[n] ?? String(n);
  if (n < 100) return TENS[Math.floor(n / 10)]! + (n % 10 ? '-' + ONES[n % 10]!.toLowerCase() : '');
  if (n < 1000) return `${ONES[Math.floor(n / 100)]} Hundred${n % 100 ? ' and ' + cardinalWords(n % 100).toLowerCase() : ''}`;
  return String(n);
}

function ordinalWords(n: number): string {
  const words = cardinalWords(n);
  const lastSep = Math.max(words.lastIndexOf('-'), words.lastIndexOf(' '));
  const head = words.slice(0, lastSep + 1);
  const tail = words.slice(lastSep + 1);
  const cap = tail.charAt(0).toUpperCase() + tail.slice(1);
  let ord = ORDINAL_IRREGULAR[cap] ?? (cap.endsWith('y') ? cap.slice(0, -1) + 'ieth' : cap + 'th');
  if (lastSep >= 0) ord = ord.toLowerCase();
  return head + ord;
}

function formatToken(n: number, numFmt: string): string {
  switch (numFmt) {
    case 'ordinalText':
      return ordinalWords(n);
    case 'cardinalText':
      return cardinalWords(n);
    case 'upperLetter':
    case 'lowerLetter': {
      // Word repeats the letter past Z: 27 is "AA", 29 is "CC" (not "AC")
      if (n < 1) return formatInt(n, numFmt);
      const letter = String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
      return numFmt === 'lowerLetter' ? letter.toLowerCase() : letter;
    }
    case 'ordinal': {
      const s = ['th', 'st', 'nd', 'rd'];
      const v = n % 100;
      return `${n}${s[(v - 20) % 10] ?? s[v] ?? s[0]}`;
    }
    default:
      return formatInt(n, numFmt);
  }
}
