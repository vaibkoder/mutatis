/**
 * host-web — OOXML → DocSnapshot loader (ADR-0022, spec/features/12 §3, §4, phase 2).
 *
 * Builds the SAME in-memory document model the Office single-pass loader produces
 * (`@shared/doc-model`), but from a parsed `.docx` package instead of a live
 * `Word.run`. This is what lets the unchanged feature checkers (factual-consistency
 * is the phase-2 proof) run identically in the standalone web app.
 *
 * It deliberately MIRRORS the proven `fast-xml-parser` (`preserveOrder`) walk in
 * `src/features/compare/preprocess/pararecord.ts` (run-text concatenation, style
 * extraction, table-cell tracking) rather than refactoring that module — Compare is
 * host-coupled, allow-listed and heavily tested; a cross-feature extraction is
 * out-of-scope risk for this PoC. We reuse the technique and the library here.
 *
 * Fidelity rules that keep web findings identical to the add-in's:
 *  - `ParaNode.text` is the FULL paragraph text (run concatenation; tab/br/cr → a
 *    single space) — NOT Compare's normalised/prefix-stripped `normText`. Offsets
 *    and detections must line up with the Office loader.
 *  - Ids follow the canonical scheme: `body-para-{index}` for body + table-cell
 *    paragraphs (one running index in document order), `footnote-para-{index}` for
 *    footnote-body paragraphs.
 *  - `listString` / `listLevel` (phase 3a): the loader extracts each paragraph's
 *    `<w:numPr>` (`ParaNumPr`, document order, list paras only) and, WHEN list paras
 *    AND a `word/numbering.xml` are both present, calls `renderListNumbers` to merge
 *    rendered numbers onto the matching `ParaNode`s. The renderer (the counter state
 *    machine) is a STUB today (throws), so the call is GUARDED: a document with no
 *    list paragraphs / no `numbering.xml` never invokes it and stays fully working
 *    (the factual-consistency PoC needs no list numbers). A document that DOES have
 *    list paras will currently surface the stub error — by design until the renderer
 *    lands (its integration test is `.todo`).
 *
 * Host-NEUTRAL by construction: this file references no `Office.*`/`Word.*` token,
 * so `host-web` stays off the `no-host-imports` allow-list. The `ParaNode.getRange`
 * stub throws — navigation on web goes through `WebHostPort.reveal(loc)`, never
 * `getRange` (which only the Office adapter can satisfy).
 */
import type { DocxBundle } from '@shared/ooxml/index.js';
import type { ParaNode, DocSnapshot } from '@shared/doc-model/index.js';

import { type PoNode, parsePreserveOrder, tagOf, childrenOf, attrOf } from '@shared/ooxml/po-xml.js';
import { parseNumberingScheme } from './numbering-scheme.js';
import { renderListNumbers, type ParaNumPr } from './render-list-numbers.js';

/** A `ParaNode.getRange` that can never be honoured outside a live Word host. */
function noLiveRange(): never {
  throw new Error('host-web: no live Word.Range; navigate via HostPort.reveal().');
}

/**
 * Parse a `.docx` bundle into the shared `DocSnapshot` model.
 * Reads `word/document.xml` (body + table-cell paragraphs) and, when present,
 * `word/footnotes.xml` (footnote-body paragraphs, skipping separator footnotes).
 */
export function loadDocSnapshotFromOoxml(bundle: DocxBundle): DocSnapshot {
  const documentXml = bundle.part('word/document.xml');
  const footnotesXml = bundle.part('word/footnotes.xml');

  const { paragraphs, numPrs } = documentXml
    ? extractBodyParagraphs(documentXml)
    : { paragraphs: [] as ParaNode[], numPrs: [] as ParaNumPr[] };
  const footnotes = footnotesXml ? extractFootnoteParagraphs(footnotesXml) : [];

  // Phase 3a: render list numbers only when there is something to render AND a
  // numbering.xml to render it from. The no-list path (factual-consistency PoC and
  // every current test) never reaches the (still-stubbed) renderer.
  if (numPrs.length > 0 && bundle.hasPart('word/numbering.xml')) {
    const scheme = parseNumberingScheme(bundle);
    const rendered = renderListNumbers(scheme, numPrs);
    const byId = new Map(paragraphs.map(p => [p.id, p]));
    for (const [paraId, ln] of rendered) {
      const para = byId.get(paraId);
      if (para) {
        para.listString = ln.listString;
        para.listLevel = ln.listLevel;
      }
    }
  }

  return { paragraphs, footnotes, loadedAt: Date.now() };
}

/**
 * Collect every list paragraph's numbering reference (`ParaNumPr`) from a bundle,
 * in document order — the exact input `renderListNumbers` consumes. Non-list paras
 * (no `<w:numPr>`, or `numId` "0"/missing) are excluded. Absent `document.xml` → [].
 *
 * Exposed so the renderer (and its tests) can drive off the same extraction the
 * loader uses, without re-implementing the walk.
 */
export function collectParaNumPrs(bundle: DocxBundle): ParaNumPr[] {
  const documentXml = bundle.part('word/document.xml');
  return documentXml ? extractBodyParagraphs(documentXml).numPrs : [];
}

// ---------------------------------------------------------------------------
// Body (document.xml) — paragraphs in document order, table-cell aware
// ---------------------------------------------------------------------------

interface CellCtx {
  tableId: string;
  row: number;
  col: number;
}

function extractBodyParagraphs(documentXml: string): {
  paragraphs: ParaNode[];
  numPrs: ParaNumPr[];
} {
  const tree = parsePreserveOrder(documentXml);
  const out: ParaNode[] = [];
  const numPrs: ParaNumPr[] = [];
  let index = 0;
  let tableCounter = 0;

  const push = (paraNode: PoNode, cell: CellCtx | null): void => {
    const id = `body-para-${index}`;
    const base = {
      id,
      index,
      text: extractRunText(paraNode),
      styleName: extractStyleId(paraNode) ?? 'Normal',
      getRange: noLiveRange,
    };
    out.push(
      cell
        ? { ...base, container: 'tableCell', tableInfo: { ...cell } }
        : { ...base, container: 'body' },
    );

    // Collect the paragraph's numbering reference (list paras only, document order).
    const numPr = extractNumPr(paraNode);
    if (numPr) numPrs.push({ paraId: id, numId: numPr.numId, ilvl: numPr.ilvl });

    index += 1;
  };

  // Walk in document order. `cell` carries the current table-cell context (null at
  // top level); tables encountered open a new synthesised tableId and 0-based row/col.
  const walk = (nodes: PoNode[], cell: CellCtx | null): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;

      if (tag === 'w:p') {
        push(node, cell);
        continue; // a <w:p> never contains a nested <w:p>
      }

      if (tag === 'w:tbl') {
        walkTable(node, `table-${tableCounter++}`);
        continue;
      }

      const value = node[tag];
      if (Array.isArray(value)) walk(value as PoNode[], cell);
    }
  };

  // NOTE (phase 3/4 re-validation): this synthesises tableId/row/col directly from
  // <w:tbl>/<w:tr>/<w:tc> nesting, whereas the Office loader INFERS table boundaries
  // from row/col regressions in its flat paragraph list. The two can disagree on
  // nested tables / multi-paragraph cells. Harmless now (factual-consistency never
  // reads tableInfo), but when a tableInfo-consuming feature (cross-ref/compare) is
  // ported to web, re-validate that table coordinates match the add-in's.
  const walkTable = (tblNode: PoNode, tableId: string): void => {
    const rows = childrenOf(tblNode);
    let rowIdx = -1;
    for (const rowNode of rows) {
      if (tagOf(rowNode) !== 'w:tr') continue;
      rowIdx += 1;
      let colIdx = -1;
      for (const cellNode of childrenOf(rowNode)) {
        if (tagOf(cellNode) !== 'w:tc') continue;
        colIdx += 1;
        walk(childrenOf(cellNode), { tableId, row: rowIdx, col: colIdx });
      }
    }
  };

  walk(tree, null);
  return { paragraphs: out, numPrs };
}

// ---------------------------------------------------------------------------
// numPr extraction — <w:pPr><w:numPr> (<w:ilvl>, <w:numId>)
// ---------------------------------------------------------------------------

/**
 * Read a paragraph's numbering reference from `<w:pPr><w:numPr>`.
 * Returns undefined when the paragraph is not a list item: no `<w:numPr>`, or a
 * `numId` of "0" / missing (Word's "no list" sentinel). `ilvl` defaults to 0 when
 * `<w:ilvl>` is absent (the OOXML default).
 *
 * Direct `<w:pPr><w:numPr>` takes precedence over any style-linked numbering
 * (`resolveStyleLinkedNumPr`).
 */
function extractNumPr(paraNode: PoNode): { numId: string; ilvl: number } | undefined {
  const pPr = childrenOf(paraNode).find(c => tagOf(c) === 'w:pPr');
  if (!pPr) return resolveStyleLinkedNumPr(paraNode);

  const numPr = childrenOf(pPr).find(c => tagOf(c) === 'w:numPr');
  if (!numPr) return resolveStyleLinkedNumPr(paraNode);

  let numId: string | undefined;
  let ilvl = 0;
  for (const child of childrenOf(numPr)) {
    const tag = tagOf(child);
    if (tag === 'w:numId') {
      numId = attrOf(child, '@_w:val');
    } else if (tag === 'w:ilvl') {
      const v = attrOf(child, '@_w:val');
      if (v !== undefined) {
        const n = parseInt(v, 10);
        if (!Number.isNaN(n)) ilvl = n;
      }
    }
  }

  if (numId === undefined || numId === '0') return undefined;
  return { numId, ilvl };
}

/**
 * DEFERRED HOOK — style-linked numbering. A paragraph can inherit its `<w:numPr>`
 * from its paragraph style (`<w:pStyle>` → `word/styles.xml` `<w:pPr><w:numPr>`,
 * possibly through a `<w:basedOn>` chain). Resolving that requires parsing
 * `styles.xml` and walking the style inheritance chain; it is intentionally NOT
 * implemented in this scaffold (it is orthogonal to the counter state machine that
 * is the next session's single task).
 *
 * Currently returns undefined — such paragraphs are treated as non-list. This is a
 * KNOWN GAP, not a silent drop: see the `.todo` test "resolves style-linked
 * numbering via pStyle → styles.xml". When
 * implemented, build a `styleId → {numId, ilvl}` map (honouring `basedOn`) and look
 * up the paragraph's `<w:pStyle w:val>`.
 */
function resolveStyleLinkedNumPr(_paraNode: PoNode): { numId: string; ilvl: number } | undefined {
  return undefined;
}

// ---------------------------------------------------------------------------
// Footnotes (footnotes.xml) — body paragraphs of real footnotes only
// ---------------------------------------------------------------------------

const SEPARATOR_FOOTNOTE_TYPES = new Set(['separator', 'continuationSeparator']);

function extractFootnoteParagraphs(footnotesXml: string): ParaNode[] {
  const tree = parsePreserveOrder(footnotesXml);
  const out: ParaNode[] = [];
  let index = 0;

  const collectParas = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:p') {
        out.push({
          id: `footnote-para-${index}`,
          index,
          text: extractRunText(node),
          styleName: extractStyleId(node) ?? 'Normal',
          container: 'footnote',
          getRange: noLiveRange,
        });
        index += 1;
        continue;
      }
      const value = node[tag];
      if (Array.isArray(value)) collectParas(value as PoNode[]);
    }
  };

  // Find <w:footnotes> → its <w:footnote> children; skip the separator footnotes.
  const walkFootnotes = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === 'w:footnote') {
        const type = attrOf(node, '@_w:type');
        if (type && SEPARATOR_FOOTNOTE_TYPES.has(type)) continue;
        collectParas(childrenOf(node));
        continue;
      }
      const value = tag ? node[tag] : undefined;
      if (Array.isArray(value)) walkFootnotes(value as PoNode[]);
    }
  };

  walkFootnotes(tree);
  return out;
}

// ---------------------------------------------------------------------------
// Shared node helpers (same shapes as compare/preprocess/pararecord.ts)
// ---------------------------------------------------------------------------

/** Concatenate all run text within a paragraph; tab/br/cr collapse to a space. */
function extractRunText(paraNode: PoNode): string {
  const parts: string[] = [];
  const collect = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:t') {
        parts.push(textOf(node));
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
  const value = paraNode['w:p'];
  if (Array.isArray(value)) collect(value as PoNode[]);
  return parts.join('');
}

/** The text content of a `<w:t>` node (its `#text` children), or "". */
function textOf(node: PoNode): string {
  const children = node['w:t'];
  if (!Array.isArray(children)) return '';
  let s = '';
  for (const child of children as PoNode[]) {
    const t = (child as Record<string, unknown>)['#text'];
    if (typeof t === 'string') s += t;
  }
  return s;
}

/** Read `<w:pPr>/<w:pStyle w:val="…">`; null when absent. */
function extractStyleId(paraNode: PoNode): string | null {
  for (const child of childrenOf(paraNode)) {
    if (tagOf(child) !== 'w:pPr') continue;
    for (const pPrChild of childrenOf(child)) {
      if (tagOf(pPrChild) !== 'w:pStyle') continue;
      const val = attrOf(pPrChild, '@_w:val');
      return val ?? null;
    }
    return null;
  }
  return null;
}
