/**
 * Paginator (feature 16 §4 hard-parts #1 & pagination) — host-free, pure.
 *
 * Stacks laid-out lines onto pages from one shared layout, honouring:
 *  - A4 geometry from sectPr;
 *  - widow/orphan control (≥2 lines either side of a split);
 *  - keepLines (don't split a paragraph) and keepNext (don't orphan a heading);
 *  - pageBreakBefore + paragraph space-before/after;
 *  - footnote bottom-of-page reservation via the Spike-B forward fixpoint:
 *    committing a footnote shrinks the body area, so a ref line that no longer
 *    fits ends the page before it; oversized footnotes split with a continuation;
 *  - real tables (tables.ts): rows placed whole, tall splittable rows split across
 *    pages, cantSplit rows kept together, borders emitted as rects — graceful
 *    degradation (content rendered, never crashed/dropped).
 */
import {
  type RenderModel, type LayoutPage, type LayoutResult, type LayoutLine,
  type FootnoteRegion, type Rect, contentWidthPt, contentHeightPt,
} from '../types.js';
import type { Shaper } from '../layout/shape.js';
import type { FontResolver } from '../fonts/registry.js';
import { layoutParagraph, type ParagraphLayout } from '../layout/paragraph.js';
import { layoutTable, splitRow, type TableRowLayout } from './tables.js';

const FOOTNOTE_SEP_PT = 12; // separator rule + gap above the footnote area
const WIDOW_MIN = 2;
const EPS = 0.01;

type Line = Omit<LayoutLine, 'topPt'>;

interface ParaUnit {
  kind: 'para';
  lines: Line[];
  spaceBeforePt: number;
  spaceAfterPt: number;
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
}
interface TableUnit {
  kind: 'table';
  rows: TableRowLayout[];
}
type Unit = ParaUnit | TableUnit;

function footnoteIdsOf(line: Line): string[] {
  const ids: string[] = [];
  for (const r of line.runs) if (r.footnoteId !== undefined) ids.push(r.footnoteId);
  return ids;
}
const sumH = (lines: Line[]): number => lines.reduce((a, l) => a + l.heightPt, 0);

/** Flatten the model's top-level blocks into placeable units (paragraphs + tables). */
function toUnits(model: RenderModel, cw: number, shaper: Shaper, fonts: FontResolver): Unit[] {
  const units: Unit[] = [];
  for (const b of model.blocks) {
    if (b.kind === 'paragraph') {
      const pl: ParagraphLayout = layoutParagraph(b.paragraph, cw, shaper, fonts);
      units.push({
        kind: 'para',
        lines: pl.lines, spaceBeforePt: pl.spaceBeforePt, spaceAfterPt: pl.spaceAfterPt,
        keepNext: pl.keepNext, keepLines: pl.keepLines, widowControl: pl.widowControl,
        pageBreakBefore: pl.pageBreakBefore,
      });
    } else {
      units.push({ kind: 'table', rows: layoutTable(b.table, cw, shaper, fonts).rows });
    }
  }
  return units;
}

/** Lay out each footnote body into lines (cached by id). */
function footnoteLines(model: RenderModel, cw: number, shaper: Shaper, fonts: FontResolver): Map<string, Line[]> {
  const out = new Map<string, Line[]>();
  for (const [id, fn] of model.footnotes) {
    const lines: Line[] = [];
    for (const b of fn.blocks) {
      if (b.kind === 'paragraph') lines.push(...layoutParagraph(b.paragraph, cw, shaper, fonts).lines);
    }
    out.set(id, lines);
  }
  return out;
}

interface PageState {
  body: Array<{ line: Line; topPt: number }>;
  y: number; // body cursor from content-area top
  fnLines: Line[]; // committed footnote lines (continuation first, then per-ref)
  continuedFromPrev: boolean;
  rects: Rect[]; // table borders/shading (highlight rects are added by decorate())
}

/** Height the next unit needs to even START on the current page (for keepNext). */
function firstUnitHeight(unit: Unit): number {
  return unit.kind === 'para' ? (unit.lines[0]?.heightPt ?? 0) : (unit.rows[0]?.heightPt ?? 0);
}

export function paginate(model: RenderModel, shaper: Shaper, fonts: FontResolver): LayoutResult {
  const g = model.geometry;
  const cw = contentWidthPt(g);
  const ch = contentHeightPt(g);
  const units = toUnits(model, cw, shaper, fonts);
  const fnMap = footnoteLines(model, cw, shaper, fonts);

  const pages: LayoutPage[] = [];
  let carry: Line[] = []; // oversized-footnote continuation lines for the next page

  const fnArea = (committed: Line[], carryIn: Line[]): number =>
    committed.length + carryIn.length === 0 ? 0 : FOOTNOTE_SEP_PT + sumH(committed) + sumH(carryIn);

  let page: PageState = { body: [], y: 0, fnLines: [], continuedFromPrev: false, rects: [] };

  const startPage = (): void => {
    page = { body: [], y: 0, fnLines: [], continuedFromPrev: carry.length > 0, rects: [] };
    // place continuation lines (they live in the footnote area; split again if needed)
    if (carry.length > 0) {
      const cap = ch - FOOTNOTE_SEP_PT;
      const fit: Line[] = [];
      let h = 0;
      let rest: Line[] = [];
      for (let k = 0; k < carry.length; k++) {
        const ln = carry[k]!;
        if (h + ln.heightPt <= cap || fit.length === 0) { fit.push(ln); h += ln.heightPt; }
        else { rest = carry.slice(k); break; }
      }
      page.fnLines = fit;
      carry = rest;
    }
  };

  const finishPage = (): void => {
    const region = buildRegion(page, ch);
    const lp: LayoutPage = { geometry: g, bodyLines: page.body.map((b) => ({ ...b.line, topPt: b.topPt })), rects: page.rects };
    if (region) lp.footnotes = region;
    pages.push(lp);
  };

  startPage();

  // ── Table row placement (whole rows; split splittable rows; keep cantSplit) ──
  const placeRow = (row: TableRowLayout): void => {
    for (const cl of row.cellLines) page.body.push({ line: cl.line, topPt: page.y + cl.relTopPt });
    for (const r of row.rects) page.rects.push({ ...r, yPt: page.y + r.yPt });
    page.y += row.heightPt;
  };
  const placeTable = (rows: TableRowLayout[]): void => {
    const work = [...rows];
    while (work.length) {
      const row = work.shift()!;
      const remaining = ch - fnArea(page.fnLines, carry) - page.y;
      if (row.heightPt <= remaining + EPS) { placeRow(row); continue; }
      if (page.body.length > 0) {
        if (!row.cantSplit) {
          const [head, tail] = splitRow(row, remaining);
          if (tail && head.heightPt <= remaining + EPS) {
            placeRow(head); finishPage(); startPage(); work.unshift(tail); continue;
          }
        }
        finishPage(); startPage(); work.unshift(row); continue;
      }
      // empty page and still too tall → split to avoid overflow (degrade even cantSplit)
      const avail = ch - fnArea(page.fnLines, carry);
      const [head, tail] = splitRow(row, avail);
      if (tail && head.heightPt <= avail + EPS) { placeRow(head); finishPage(); startPage(); work.unshift(tail); continue; }
      placeRow(row); // single row/line taller than a clean page → place anyway, never drop
    }
  };

  for (let u = 0; u < units.length; u++) {
    const unit = units[u]!;

    if (unit.kind === 'table') {
      placeTable(unit.rows);
      continue;
    }

    if (unit.pageBreakBefore && page.body.length > 0) { finishPage(); startPage(); }

    // space-before (suppressed at page top)
    if (page.body.length > 0) page.y += unit.spaceBeforePt;

    let placedAny = false;
    let lineIdx = 0;
    while (lineIdx < unit.lines.length) {
      const remaining = unit.lines.length - lineIdx;
      // Decide how many of the remaining lines fit on this page (with footnotes).
      let k = 0;
      let probeY = page.y;
      let probeFn = [...page.fnLines];
      for (let n = lineIdx; n < unit.lines.length; n++) {
        const ln = unit.lines[n]!;
        const ids = footnoteIdsOf(ln);
        const addFn = ids.flatMap((id) => fnMap.get(id) ?? []);
        const tentativeFn = [...probeFn, ...addFn];
        const avail = ch - fnArea(tentativeFn, carry);
        if (probeY + ln.heightPt <= avail + EPS) { k++; probeY += ln.heightPt; probeFn = tentativeFn; }
        else break;
      }

      // keepLines: do not split — need all remaining to fit, else move whole to next page.
      if (unit.keepLines && k < remaining) k = (page.body.length === 0 ? remaining : 0);

      // widow/orphan: avoid leaving <2 lines on either side of a split.
      if (unit.widowControl && k > 0 && k < remaining) {
        if (k < WIDOW_MIN) k = 0; // orphan: push start to next page
        else if (remaining - k < WIDOW_MIN) k = Math.max(0, remaining - WIDOW_MIN); // widow: pull a line down
        if (k < WIDOW_MIN && page.body.length === 0) k = Math.min(remaining, k || 1); // can't help on empty page
      }

      if (k === 0) {
        if (page.body.length > 0) { finishPage(); startPage(); continue; }
        // Empty page. If footnotes/continuation saturate it but the line would fit on
        // a page with less footnote load, flush (footnote-only page) and retry.
        const ln = unit.lines[lineIdx]!;
        const addFn = footnoteIdsOf(ln).flatMap((id) => fnMap.get(id) ?? []);
        const needed = page.y + ln.heightPt + fnArea([...page.fnLines, ...addFn], carry);
        const fitsCleanPage = ln.heightPt + fnArea(addFn, []) <= ch + EPS;
        if (needed > ch + EPS && (page.fnLines.length > 0 || carry.length > 0) && fitsCleanPage) {
          finishPage(); startPage(); continue;
        }
        k = 1; // line genuinely taller than a clean page → place anyway (degrade, don't drop)
      }

      // commit k lines (push body THEN reserve its footnotes, so the reservation
      // accounts for the ref line's own height — fixes a 1-line overflow).
      for (let n = 0; n < k; n++) {
        const ln = unit.lines[lineIdx]!;
        page.body.push({ line: ln, topPt: page.y });
        page.y += ln.heightPt;
        for (const id of footnoteIdsOf(ln)) commitFootnote(page, fnMap.get(id) ?? [], ch, (rest) => { carry = carry.concat(rest); });
        lineIdx++;
        placedAny = true;
      }
      if (lineIdx < unit.lines.length) { finishPage(); startPage(); }
    }

    // keepNext: don't strand a heading at the bottom — if the next unit can't even
    // start one line on this page, move this whole unit to the next page.
    if (unit.keepNext && placedAny && u + 1 < units.length && page.body.length > 0) {
      const firstNext = firstUnitHeight(units[u + 1]!);
      if (firstNext > 0) {
        const avail = ch - fnArea(page.fnLines, carry);
        if (page.y + firstNext > avail + EPS) {
          // remove this unit's just-placed lines from this page and re-place on a fresh page
          const moved = page.body.splice(page.body.length - unit.lines.length, unit.lines.length);
          if (moved.length === unit.lines.length) {
            finishPage(); startPage();
            for (const m of moved) { page.body.push({ line: m.line, topPt: page.y }); page.y += m.line.heightPt; }
          }
        }
      }
    }

    if (page.body.length > 0) page.y += unit.spaceAfterPt;
  }

  finishPage();
  // flush any remaining oversized-footnote continuation onto trailing pages
  while (carry.length > 0) { startPage(); finishPage(); }

  return { pages };
}

/** Commit a footnote's lines to the page; if they don't fit, split (carry the rest). */
function commitFootnote(page: PageState, lines: Line[], ch: number, onCarry: (rest: Line[]) => void): void {
  if (lines.length === 0) return;
  const sep = FOOTNOTE_SEP_PT;
  // available footnote height = whole content minus the body already placed and the separator
  const bodyUsed = page.y;
  const already = sumH(page.fnLines);
  const room = ch - bodyUsed - sep - already;
  let h = 0;
  const fit: Line[] = [];
  for (let k = 0; k < lines.length; k++) {
    const ln = lines[k]!;
    if (h + ln.heightPt <= room || (page.fnLines.length === 0 && fit.length === 0)) { fit.push(ln); h += ln.heightPt; }
    else { onCarry(lines.slice(k)); break; }
  }
  page.fnLines.push(...fit);
}

function buildRegion(page: PageState, ch: number): FootnoteRegion | undefined {
  if (page.fnLines.length === 0) return undefined;
  const areaH = sumH(page.fnLines);
  const separatorTopPt = ch - areaH - FOOTNOTE_SEP_PT + 2;
  let y = ch - areaH;
  const lines: LayoutLine[] = page.fnLines.map((l) => {
    const ln: LayoutLine = { ...l, topPt: y };
    y += l.heightPt;
    return ln;
  });
  return {
    separatorTopPt,
    lines,
    continuedFromPrev: page.continuedFromPrev,
    continuesToNext: false,
  };
}
