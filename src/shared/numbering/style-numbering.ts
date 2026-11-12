/**
 * Shared numbering — a paragraph's EFFECTIVE `<w:numPr>`, including the one it
 * never states: the one it inherits from its paragraph style.
 *
 * Word renders a clause number from `(numId, ilvl)`, and a paragraph may name
 * neither. A firm's house template typically puts `<w:numPr>` on the STYLE
 * (`Subclause2`, `Heading2`) and leaves every paragraph carrying nothing
 * but `<w:pStyle>`. A reader who reads only `<w:pPr><w:numPr>` therefore sees an
 * entire contract as unnumbered — which is exactly what both consumers did
 * before this module: Compare's `ParaRecord` (so a re-levelled clause compared
 * EQUAL) and RenderCore's `parse-body` (so a real label could not be painted).
 *
 * ECMA-376 §17.3.1.19: the numbering reference resolves through the style
 * hierarchy, and a direct `<w:pPr><w:numPr>` overrides it. `numId="0"` is the
 * documented "no numbering" sentinel and is honoured as a removal, not as a
 * reference to instance zero.
 *
 * Deliberately structural, not nominal: `StyleNumberingSource` is the SHAPE this
 * resolution needs, and `parseStyles`' `StyleTable` (render-core `input/styles`)
 * already satisfies it. So there is one styles.xml parser, one basedOn walk per
 * concern, and NO import from `@shared/numbering` into `@shared/render-core`,
 * which would invert the layering (render-core consumes numbering).
 *
 * Host-NEUTRAL: no `Office.*`/`Word.*` token.
 */
import { type PoNode, parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';

/** A paragraph's numbering reference: the instance and the 0-based level. */
export interface NumPrRef {
  numId: string;
  ilvl: number;
}

/**
 * The minimum a style table must expose for numbering resolution: each style's
 * `basedOn` parent and its `<w:pPr>` children, plus the document's default
 * paragraph style (a paragraph with no `<w:pStyle>` inherits from it).
 */
export interface StyleNumberingSource {
  styles: ReadonlyMap<string, { basedOn: string | null; pPr: readonly PoNode[] }>;
  defaultParaStyleId: string | null;
}

/**
 * Read `<w:numPr>` out of a `<w:pPr>`'s children. Either half may be absent:
 * `<w:numPr><w:ilvl w:val="2"/></w:numPr>` re-levels a paragraph while keeping
 * the instance it inherits, and is the shape an outdent takes when the author
 * used the Decrease-Indent button rather than restyling.
 *
 * Returns null when there is no `<w:numPr>` at all — distinct from a `<w:numPr>`
 * that is present but carries no `numId` of its own.
 */
export function readNumPr(pPr: readonly PoNode[]): { numId?: string; ilvl?: number } | null {
  for (const node of pPr) {
    if (tagOf(node) !== 'w:numPr') continue;
    const out: { numId?: string; ilvl?: number } = {};
    for (const child of childrenOf(node)) {
      const tag = tagOf(child);
      if (tag === 'w:numId') {
        const v = attrOf(child, '@_w:val');
        if (v !== undefined) out.numId = v;
      } else if (tag === 'w:ilvl') {
        // Strict: the attribute must be PRESENT and parse to a non-negative
        // integer. `Number('')` is 0, so a valueless `<w:ilvl/>` would otherwise
        // silently re-level the paragraph to level 0 and override its style —
        // while `<w:ilvl w:val="abc"/>`, no more malformed, was correctly
        // ignored. Two equally broken inputs, handled opposite ways.
        const raw = attrOf(child, '@_w:val');
        if (raw !== undefined && /^\d+$/.test(raw.trim())) out.ilvl = Number(raw.trim());
      }
    }
    return out;
  }
  return null;
}

/**
 * The `basedOn` chain for a style, LEAF FIRST. Leaf-first because numbering
 * resolution wants the nearest definition and stops at it, where the formatting
 * cascade wants root-first and applies all of them. Cycle-safe.
 */
function chainFromLeaf(source: StyleNumberingSource, styleId: string | null): PoNode[][] {
  const out: PoNode[][] = [];
  const seen = new Set<string>();
  let cur = styleId;
  while (cur !== null && cur !== undefined && !seen.has(cur)) {
    seen.add(cur);
    const style = source.styles.get(cur);
    if (!style) break;
    out.push([...style.pPr]);
    cur = style.basedOn;
  }
  return out;
}

/**
 * Resolve a paragraph's effective numbering reference.
 *
 * `numId` and `ilvl` are resolved **independently**, each from its own nearest
 * declaration walking direct formatting → `pStyle` → `basedOn` → the default
 * paragraph style (an unstyled paragraph still inherits from `w:default="1"`).
 * They are two leaf properties that happen to share a container, and ECMA-376
 * §17.7.2 merges the style hierarchy leaf by leaf — `readNumPr`'s own doc
 * comment says so for the half this function used to get right.
 *
 * **An earlier version treated `<w:numPr>` as atomic in one direction only**, and
 * both halves of that asymmetry were defects found by review:
 *
 *  - a style declaring only `<w:ilvl>` — `basedOn` a numbered base style,
 *    overriding nothing but the level, which is ordinary template construction —
 *    was skipped entirely, so `Sub` at level 2 resolved to level 0 and a list
 *    re-levelling the reader sees as `(a) (b)` → `(i) (ii)` produced **no op at
 *    all** (ADR-0008 suppression, measured `summary.total` 0);
 *  - and a direct `<w:numId>` with no `<w:ilvl>` discarded the level its style
 *    declared, resolving to 0 and reporting a `formatting` change on a pair of
 *    documents that differ only in whether a redundant `<w:ilvl>` is spelled out.
 *
 * `numId="0"` is the documented removal sentinel: when the NEAREST `numId`
 * declaration is `"0"` the paragraph is not numbered, whatever a further
 * ancestor says. Returns null then, and when no declaration numbers it at all.
 */
export function resolveParaNumPr(
  source: StyleNumberingSource,
  styleId: string | null,
  directPPr: readonly PoNode[],
): NumPrRef | null {
  const chain = chainFromLeaf(source, styleId ?? source.defaultParaStyleId);
  // Direct formatting first, then leaf style to root: nearest declaration wins,
  // per property. A `<w:numPr>` with only one of the two children declares only
  // that one and is transparent to the other — it is not a barrier.
  const declarations = [readNumPr(directPPr), ...chain.map(readNumPr)];

  let numId: string | undefined;
  let ilvl: number | undefined;
  for (const dec of declarations) {
    if (dec === null) continue;
    if (numId === undefined && dec.numId !== undefined) numId = dec.numId;
    if (ilvl === undefined && dec.ilvl !== undefined) ilvl = dec.ilvl;
    if (numId !== undefined && ilvl !== undefined) break;
  }

  if (numId === undefined || numId === '0') return null;
  return { numId, ilvl: ilvl ?? 0 };
}

/**
 * Build a `StyleNumberingSource` from a raw `styles.xml`.
 *
 * For callers that do NOT already hold a style table — Compare reads its parts
 * as strings and must not take a dependency on RenderCore's input layer to get
 * at three fields (that edge would invert the layering: render-core consumes
 * compare, not the reverse). RenderCore itself never calls this: its
 * `parseStyles` table already satisfies `StyleNumberingSource` structurally, so
 * it passes that straight to `resolveParaNumPr` and parses styles.xml once.
 *
 * `defaultParaStyleId` is the FIRST `<w:style w:type="paragraph">` whose
 * `w:default` is an ECMA-376 TRUE (`1`/`true`/`on`). Note that render-core's
 * `parseStyles` treats the attribute's mere presence as true, so the two
 * disagree on the malformed `w:default="0"`; this one is right, and the
 * divergence is deliberate rather than fixed here, because
 * `parseStyles` feeds a gated deliverable's layout.
 *
 * Absent/empty part → an empty source (every paragraph then resolves to
 * un-numbered, exactly as before this module existed).
 */
export function parseStyleNumberingSource(
  stylesXml: string | null | undefined,
): StyleNumberingSource {
  const styles = new Map<string, { basedOn: string | null; pPr: PoNode[] }>();
  const source: StyleNumberingSource = { styles, defaultParaStyleId: null };
  if (!stylesXml) return source;

  const root = parsePreserveOrder(stylesXml).find((n) => tagOf(n) === 'w:styles');
  if (!root) return source;

  let defaultParaStyleId: string | null = null;
  for (const node of childrenOf(root)) {
    if (tagOf(node) !== 'w:style') continue;
    const id = attrOf(node, '@_w:styleId');
    if (id === undefined) continue;

    let basedOn: string | null = null;
    let pPr: PoNode[] = [];
    for (const kid of childrenOf(node)) {
      const tag = tagOf(kid);
      if (tag === 'w:basedOn') basedOn = attrOf(kid, '@_w:val') ?? null;
      else if (tag === 'w:pPr') pPr = childrenOf(kid);
    }
    styles.set(id, { basedOn, pPr });

    // ECMA-376 booleans are `1/0/true/false/on/off`, so `w:default="0"` means
    // NOT the default. Treating presence as truth made the first
    // `w:default="0"` paragraph style the document default, and every unstyled
    // paragraph then inherited its numbering.
    if (defaultParaStyleId === null && attrOf(node, '@_w:type') === 'paragraph') {
      const flag = attrOf(node, '@_w:default');
      if (flag !== undefined && flag !== '0' && flag !== 'false' && flag !== 'off') {
        defaultParaStyleId = id;
      }
    }
  }
  return { styles, defaultParaStyleId };
}
