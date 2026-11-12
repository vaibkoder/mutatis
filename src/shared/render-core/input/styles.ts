/**
 * Style-cascade resolution (feature 16 §4 hard-part #2) — host-free, pure.
 *
 * Computes effective rPr/pPr by the OOXML cascade:
 *   docDefaults → basedOn style chain (root→leaf) → paragraph-mark/direct pPr/rPr.
 * Numbering-level pPr is folded in by the caller (parse-body) via the paragraph's
 * own w:ind where present (numbering-level indents are a noted deferred bound).
 *
 * Toggle properties (b, i, strike, caps, …) follow ECMA-376 §17.7.3: an implied
 * toggle (`<w:b/>` or `w:val` true) XORs the inherited value; an explicit false
 * (`w:val="0"/false/off`) sets it absolutely. So a bold style + a direct `<w:b/>`
 * yields NON-bold (Word's documented surprise), while `<w:b w:val="0"/>` is a
 * hard off. Non-toggle properties (size, family, indents, spacing) are
 * last-writer-wins down the chain.
 */
import { type PoNode, parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';
import type { EffectiveRPr, EffectivePPr, LineRule } from '../types.js';

const HALF_PT = 2; // w:sz is in half-points
const TWIP = 20; // 20 twips = 1pt

function val(node: PoNode): string | undefined {
  return attrOf(node, '@_w:val');
}
/** Toggle reading: undefined = absent; true = implied-on; false = explicit-off. */
function toggleVal(node: PoNode): boolean {
  const v = val(node);
  if (v === undefined) return true; // <w:b/> implies on
  return !(v === '0' || v === 'false' || v === 'off');
}
function num(node: PoNode, attr = '@_w:val'): number | undefined {
  const v = attrOf(node, attr);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

export function defaultRPr(): EffectiveRPr {
  return {
    fontFamily: 'Calibri', // modern Word default (→ Carlito clone); overridden by docDefaults/theme
    sizePt: 11,
    bold: false,
    italic: false,
    underline: false,
    superscript: false,
    subscript: false,
  };
}
export function defaultPPr(): EffectivePPr {
  return {
    justification: 'left',
    indentLeftPt: 0,
    indentRightPt: 0,
    firstLinePt: 0,
    hangingPt: 0,
    spaceBeforePt: 0,
    spaceAfterPt: 0,
    lineRule: 'auto',
    lineValue: 0,
    keepNext: false,
    keepLines: false,
    widowControl: true, // Word default-on
    pageBreakBefore: false,
  };
}

/** Apply an `<w:rPr>` child list onto an accumulating EffectiveRPr (mutates a copy). */
export function applyRPr(state: EffectiveRPr, rPr: PoNode[] | undefined): EffectiveRPr {
  if (!rPr) return state;
  const out = { ...state };
  for (const node of rPr) {
    const t = tagOf(node);
    switch (t) {
      case 'w:rFonts': {
        const ascii = attrOf(node, '@_w:ascii') ?? attrOf(node, '@_w:hAnsi');
        if (ascii) out.fontFamily = ascii;
        break;
      }
      case 'w:sz': {
        const n = num(node);
        if (n !== undefined) out.sizePt = n / HALF_PT;
        break;
      }
      case 'w:b': out.bold = combineToggle(out.bold, node); break;
      case 'w:i': out.italic = combineToggle(out.italic, node); break;
      case 'w:u': {
        const v = val(node);
        out.underline = v !== undefined && v !== 'none';
        break;
      }
      case 'w:vertAlign': {
        const v = val(node);
        out.superscript = v === 'superscript';
        out.subscript = v === 'subscript';
        break;
      }
      default:
        break;
    }
  }
  return out;
}

function combineToggle(inherited: boolean, node: PoNode): boolean {
  const v = val(node);
  if (v === '0' || v === 'false' || v === 'off') return false; // absolute off
  return !inherited; // implied/true → XOR
}

/** Apply a `<w:pPr>` child list onto an accumulating EffectivePPr (mutates a copy). */
export function applyPPr(state: EffectivePPr, pPr: PoNode[] | undefined): EffectivePPr {
  if (!pPr) return state;
  const out = { ...state };
  for (const node of pPr) {
    const t = tagOf(node);
    switch (t) {
      case 'w:jc': {
        const v = val(node);
        if (v === 'both' || v === 'distribute') out.justification = 'both';
        else if (v === 'center') out.justification = 'center';
        else if (v === 'right' || v === 'end') out.justification = 'right';
        else if (v === 'left' || v === 'start') out.justification = 'left';
        break;
      }
      case 'w:ind': {
        const left = num(node, '@_w:left') ?? num(node, '@_w:start');
        const right = num(node, '@_w:right') ?? num(node, '@_w:end');
        const firstLine = num(node, '@_w:firstLine');
        const hanging = num(node, '@_w:hanging');
        if (left !== undefined) out.indentLeftPt = left / TWIP;
        if (right !== undefined) out.indentRightPt = right / TWIP;
        if (hanging !== undefined) { out.hangingPt = hanging / TWIP; out.firstLinePt = 0; }
        else if (firstLine !== undefined) { out.firstLinePt = firstLine / TWIP; out.hangingPt = 0; }
        break;
      }
      case 'w:spacing': {
        const before = num(node, '@_w:before');
        const after = num(node, '@_w:after');
        const line = num(node, '@_w:line');
        const lineRule = attrOf(node, '@_w:lineRule');
        if (before !== undefined) out.spaceBeforePt = before / TWIP;
        if (after !== undefined) out.spaceAfterPt = after / TWIP;
        if (line !== undefined) {
          out.lineValue = line;
          out.lineRule = (lineRule === 'exact' || lineRule === 'atLeast') ? (lineRule as LineRule) : 'auto';
        }
        break;
      }
      case 'w:keepNext': out.keepNext = toggleVal(node); break;
      case 'w:keepLines': out.keepLines = toggleVal(node); break;
      case 'w:widowControl': out.widowControl = toggleVal(node); break;
      case 'w:pageBreakBefore': out.pageBreakBefore = toggleVal(node); break;
      default:
        break;
    }
  }
  return out;
}

// ─── Style table ───────────────────────────────────────────────────────────────

interface StyleDef {
  id: string;
  type: string;
  name: string;
  basedOn: string | null;
  pPr: PoNode[];
  rPr: PoNode[];
  outlineLevel: number | null;
}

export interface StyleTable {
  styles: Map<string, StyleDef>;
  docDefaultRPr: PoNode[];
  docDefaultPPr: PoNode[];
  defaultParaStyleId: string | null;
}

/** rPr/pPr children of a `<w:rPr>`/`<w:pPr>` element found among `parent`'s children. */
function childPr(parent: PoNode[], wrap: 'w:rPr' | 'w:pPr'): PoNode[] {
  for (const n of parent) if (tagOf(n) === wrap) return childrenOf(n);
  return [];
}

export function parseStyles(stylesXml: string | undefined): StyleTable {
  const table: StyleTable = {
    styles: new Map(),
    docDefaultRPr: [],
    docDefaultPPr: [],
    defaultParaStyleId: null,
  };
  if (!stylesXml) return table;
  const roots = parsePreserveOrder(stylesXml);
  const stylesRoot = roots.find((n) => tagOf(n) === 'w:styles');
  if (!stylesRoot) return table;

  for (const node of childrenOf(stylesRoot)) {
    const t = tagOf(node);
    if (t === 'w:docDefaults') {
      for (const dd of childrenOf(node)) {
        if (tagOf(dd) === 'w:rPrDefault') table.docDefaultRPr = childPr(childrenOf(dd), 'w:rPr');
        if (tagOf(dd) === 'w:pPrDefault') table.docDefaultPPr = childPr(childrenOf(dd), 'w:pPr');
      }
    } else if (t === 'w:style') {
      const id = attrOf(node, '@_w:styleId');
      if (!id) continue;
      const kids = childrenOf(node);
      let name = '';
      let basedOn: string | null = null;
      let outlineLevel: number | null = null;
      for (const k of kids) {
        const kt = tagOf(k);
        if (kt === 'w:name') name = attrOf(k, '@_w:val') ?? '';
        else if (kt === 'w:basedOn') basedOn = attrOf(k, '@_w:val') ?? null;
      }
      const pPr = childPr(kids, 'w:pPr');
      for (const p of pPr) if (tagOf(p) === 'w:outlineLvl') outlineLevel = num(p) ?? null;
      table.styles.set(id, {
        id,
        type: attrOf(node, '@_w:type') ?? 'paragraph',
        name,
        basedOn,
        pPr,
        rPr: childPr(kids, 'w:rPr'),
        outlineLevel,
      });
      if (attrOf(node, '@_w:default') && attrOf(node, '@_w:type') === 'paragraph' && !table.defaultParaStyleId) {
        table.defaultParaStyleId = id;
      }
    }
  }
  return table;
}

/** The basedOn chain for a style, root ancestor first, ending at `styleId`. */
function styleChain(table: StyleTable, styleId: string | null): StyleDef[] {
  const chain: StyleDef[] = [];
  const seen = new Set<string>();
  let cur = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const s = table.styles.get(cur);
    if (!s) break;
    chain.unshift(s);
    cur = s.basedOn;
  }
  return chain;
}

export interface ResolvedParaFormat {
  ppr: EffectivePPr;
  /** Effective run props for the paragraph (style + paragraph-mark rPr); runs layer on top. */
  baseRpr: EffectiveRPr;
  isHeading: boolean;
}

/**
 * Resolve effective pPr + base rPr for a paragraph with the given style and
 * direct pPr/paragraph-mark rPr. Run-level rPr is applied later via applyRPr.
 */
export function resolveParagraph(
  table: StyleTable,
  styleId: string | null,
  directPPr: PoNode[],
  paraMarkRPr: PoNode[],
): ResolvedParaFormat {
  const effStyleId = styleId ?? table.defaultParaStyleId;
  let ppr = applyPPr(defaultPPr(), table.docDefaultPPr);
  let rpr = applyRPr(defaultRPr(), table.docDefaultRPr);

  const chain = styleChain(table, effStyleId);
  for (const s of chain) {
    ppr = applyPPr(ppr, s.pPr);
    rpr = applyRPr(rpr, s.rPr);
  }
  ppr = applyPPr(ppr, directPPr);
  rpr = applyRPr(rpr, paraMarkRPr);

  const leaf = effStyleId ? table.styles.get(effStyleId) : undefined;
  const isHeading =
    chain.some((s) => /^heading\s*\d|^heading$/i.test(s.name) || /^Heading\d/.test(s.id)) ||
    (leaf?.outlineLevel !== null && leaf?.outlineLevel !== undefined);

  return { ppr, baseRpr: rpr, isHeading };
}
