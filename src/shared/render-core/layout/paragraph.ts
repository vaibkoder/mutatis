/**
 * Paragraph layout — RcParagraph → positioned lines (host-free, pure, DOM-free).
 *
 * Word semantics (Spike A calibration, fixed policy):
 *  - greedy first-fit breaking at spaces; kerning/ligatures off (via the Shaper);
 *  - justification = inter-word space solved to the column, with up-to-20%
 *    inter-word shrink to pull one more cluster onto a line;
 *  - line height "single" = usWin formula; w:spacing line/lineRule honoured;
 *  - hanging indents + numbering label hung into the indent; first-line indent;
 *  - rounding policy: advances summed in font units, → pt once per run (the
 *    Shaper enforces this); positions accumulate in pt.
 *
 * Output lines have NO absolute y (the paginator assigns `topPt`); x is final.
 */
import type {
  RcParagraph, RcRun, PositionedRun, LayoutLine, EffectiveRPr,
} from '../types.js';
import type { Shaper } from './shape.js';
import type { FontResolver } from '../fonts/registry.js';
import { singleLineHeightPt, ascentPt as faceAscentPt } from '../fonts/sfnt-metrics.js';

const SHRINK_MAX = 0.20;
const TAB_STOP_PT = 36; // 0.5in default tab grid

export interface ParagraphLayout {
  lines: Array<Omit<LayoutLine, 'topPt'>>;
  spaceBeforePt: number;
  spaceAfterPt: number;
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
}

/** A measured atomic piece (word fragment / space / tab / hard-break / ref / image). */
interface Piece {
  text: string;
  rpr: EffectiveRPr;
  widthPt: number;
  kind: 'word' | 'space' | 'tab' | 'break';
  run: RcRun; // carries revision / changeClasses / footnoteId / imagePlaceholder
}

/** A cluster = consecutive non-space pieces that must not be split across a line. */
interface Cluster {
  pieces: Piece[];
  widthPt: number;
}

function measurePieces(p: RcParagraph, shaper: Shaper, fonts: FontResolver): Piece[] {
  const out: Piece[] = [];
  for (const run of p.runs) {
    const face = fonts.resolve(run.rpr);
    const size = run.rpr.sizePt;
    if (run.footnoteId !== undefined) {
      const text = run.footnoteId || '*';
      out.push({ text, rpr: run.rpr, widthPt: shaper.advanceWidthPt(face, text, size * 0.8), kind: 'word', run });
      continue;
    }
    if (run.imagePlaceholder) {
      out.push({ text: run.text || '￼', rpr: run.rpr, widthPt: size * 1.5, kind: 'word', run });
      continue;
    }
    // split into words / spaces / tabs / breaks, preserving order
    const parts = run.text.split(/(\s)/);
    for (const part of parts) {
      if (part === '') continue;
      if (part === '\n') { out.push({ text: part, rpr: run.rpr, widthPt: 0, kind: 'break', run }); continue; }
      if (part === '\t') { out.push({ text: part, rpr: run.rpr, widthPt: 0, kind: 'tab', run }); continue; }
      if (/^\s+$/.test(part)) {
        out.push({ text: part, rpr: run.rpr, widthPt: shaper.advanceWidthPt(face, ' ', size), kind: 'space', run });
        continue;
      }
      out.push({ text: part, rpr: run.rpr, widthPt: shaper.advanceWidthPt(face, part, size), kind: 'word', run });
    }
  }
  return out;
}

/** Single-line height + ascent for a set of pieces (max over runs present). */
function lineHeightFor(pieces: Piece[], p: RcParagraph, fonts: FontResolver): { heightPt: number; ascentPt: number } {
  const rprs = pieces.length ? pieces.map((pc) => pc.rpr) : [p.markRpr];
  let single = 0, asc = 0;
  for (const rpr of rprs) {
    const m = fonts.resolve(rpr).metrics;
    single = Math.max(single, singleLineHeightPt(m, rpr.sizePt));
    asc = Math.max(asc, faceAscentPt(m, rpr.sizePt));
  }
  // line spacing rule
  let height = single;
  if (p.ppr.lineRule === 'exact') height = p.ppr.lineValue / 20;
  else if (p.ppr.lineRule === 'atLeast') height = Math.max(single, p.ppr.lineValue / 20);
  else if (p.ppr.lineValue > 0) height = single * (p.ppr.lineValue / 240); // auto multiple
  return { heightPt: height, ascentPt: asc };
}

function positionedRun(piece: Piece, xPt: number): PositionedRun {
  const r: PositionedRun = { xPt, widthPt: piece.widthPt, text: piece.text, rpr: piece.rpr };
  if (piece.run.revision) r.revision = piece.run.revision;
  if (piece.run.changeClasses) r.changeClasses = piece.run.changeClasses;
  if (piece.run.footnoteId !== undefined) r.footnoteId = piece.run.footnoteId;
  if (piece.run.imagePlaceholder) r.imagePlaceholder = true;
  return r;
}

/**
 * Lay out one paragraph into lines within `contentWidthPt` (the content-area
 * width; indents are applied inside).
 */
export function layoutParagraph(
  p: RcParagraph,
  contentWidthPt: number,
  shaper: Shaper,
  fonts: FontResolver,
): ParagraphLayout {
  const pieces = measurePieces(p, shaper, fonts);

  // Indent geometry.
  const bodyX = p.ppr.indentLeftPt;
  const rightX = contentWidthPt - p.ppr.indentRightPt;
  const hanging = p.ppr.hangingPt > 0;
  // The first line's natural text x absent any numbering label.
  const naturalFirstTextX = hanging ? bodyX : bodyX + p.ppr.firstLinePt;
  const labelX = hanging ? Math.max(0, bodyX - p.ppr.hangingPt) : naturalFirstTextX;

  // Build the numbering label piece (first line only), if any.
  let labelPieces: Piece[] = [];
  if (p.numberLabel) {
    const face = fonts.resolve(p.markRpr);
    labelPieces = [{
      text: p.numberLabel, rpr: p.markRpr,
      widthPt: shaper.advanceWidthPt(face, p.numberLabel, p.markRpr.sizePt),
      kind: 'word', run: { text: p.numberLabel, rpr: p.markRpr },
    }];
  }

  // A hung label sits in the hanging-indent gutter and consumes no line space.
  // But when there is no gutter (numbering-level indents are a deferred bound,
  // so a numbered paragraph without its own w:ind has hangingPt = 0), or the
  // label is wider than the gutter, a free-hanging label would paint on top of
  // the first text cluster. In that case the label consumes leading space:
  // advance the first-line text to the next tab stop past the label, matching
  // Word's number → tab → text fallback. Hung labels that fit are unchanged.
  const labelWidthPt = labelPieces.reduce((a, lp) => a + lp.widthPt, 0);
  let firstTextX = naturalFirstTextX;
  if (labelPieces.length) {
    const labelRightPt = labelX + labelWidthPt;
    if (labelRightPt > naturalFirstTextX) {
      firstTextX = Math.ceil((labelRightPt - bodyX + 0.01) / TAB_STOP_PT) * TAB_STOP_PT + bodyX;
    }
  }

  // Break the pieces into lines of clusters + the inter-cluster gaps (space widths).
  const outLines: Array<Omit<LayoutLine, 'topPt'>> = [];
  let i = 0;
  let lineNo = 0;
  // pending state for the current line
  while (i < pieces.length || (lineNo === 0 && labelPieces.length && outLines.length === 0)) {
    const isFirst = lineNo === 0;
    const startX = isFirst ? firstTextX : bodyX;
    const avail = rightX - startX;

    const clusters: Cluster[] = [];
    const gaps: number[] = []; // natural space width before clusters[k], k≥1
    let pendingGap = 0;
    let used = 0;
    let endedByBreak = false;
    let tabAdjust = 0;

    // greedily accumulate clusters until the next won't fit
    while (i < pieces.length) {
      const piece = pieces[i]!;
      if (piece.kind === 'break') { i++; endedByBreak = true; break; }
      if (piece.kind === 'space') { pendingGap += piece.widthPt; i++; continue; }
      if (piece.kind === 'tab') {
        // advance to next tab stop from current natural x
        const curX = startX + used + pendingGap + tabAdjust;
        const next = Math.ceil((curX - startX + 0.01) / TAB_STOP_PT) * TAB_STOP_PT + startX;
        tabAdjust += Math.max(0, next - curX);
        i++;
        continue;
      }
      // gather a cluster: this word + following non-space/non-break pieces with no gap
      const cl: Cluster = { pieces: [piece], widthPt: piece.widthPt };
      let j = i + 1;
      while (j < pieces.length && pieces[j]!.kind === 'word') {
        cl.pieces.push(pieces[j]!);
        cl.widthPt += pieces[j]!.widthPt;
        j++;
      }
      const gap = clusters.length === 0 ? 0 : pendingGap;
      const tentative = used + gap + tabAdjust + cl.widthPt;
      if (tentative <= avail || clusters.length === 0) {
        clusters.push(cl);
        gaps.push(gap);
        used += gap + cl.widthPt;
        pendingGap = 0;
        i = j;
      } else {
        // Word shrink rule: fit one more if inter-cluster gaps can shrink ≤20%.
        const totalGapW = gaps.reduce((a, b) => a + b, 0) + gap;
        const overflow = tentative - avail;
        if (totalGapW > 0 && overflow <= totalGapW * SHRINK_MAX) {
          clusters.push(cl);
          gaps.push(gap);
          used += gap + cl.widthPt;
          pendingGap = 0;
          i = j;
        }
        break;
      }
    }

    const lastLine = i >= pieces.length;
    const justify = p.ppr.justification === 'both' && !lastLine && !endedByBreak && clusters.length > 1;

    // compute final spacing
    const gapCount = clusters.length - 1;
    const naturalGap = gaps.reduce((a, b) => a + b, 0);
    const clusterW = clusters.reduce((a, c) => a + c.widthPt, 0);
    let spacePer = (idx: number): number => gaps[idx] ?? 0;
    if (justify && gapCount > 0) {
      const target = avail - tabAdjust;
      const scale = naturalGap > 0 ? (target - clusterW) / naturalGap : 0;
      spacePer = (idx: number): number => (gaps[idx] ?? 0) * scale;
    }

    // alignment offset for left/right/center (non-justified)
    let alignOffset = 0;
    if (!justify) {
      const lineNatural = clusterW + naturalGap + tabAdjust;
      if (p.ppr.justification === 'right') alignOffset = avail - lineNatural;
      else if (p.ppr.justification === 'center') alignOffset = (avail - lineNatural) / 2;
      if (alignOffset < 0) alignOffset = 0;
    }

    // place runs
    const runs: PositionedRun[] = [];
    let x = startX + alignOffset + tabAdjust;
    if (isFirst && labelPieces.length) {
      // hung label sits at labelX (may be left of startX); doesn't consume line x
      let lx = labelX;
      for (const lp of labelPieces) { runs.push(positionedRun(lp, lx)); lx += lp.widthPt; }
    }
    clusters.forEach((cl, idx) => {
      x += spacePer(idx);
      for (const piece of cl.pieces) {
        runs.push(positionedRun(piece, x));
        x += piece.widthPt;
      }
    });

    const { heightPt, ascentPt } = lineHeightFor(clusters.flatMap((c) => c.pieces), p, fonts);
    const line: Omit<LayoutLine, 'topPt'> = { heightPt, ascentPt, runs };
    if (runs.some((r) => r.revision)) line.changeBar = true;
    outLines.push(line);
    lineNo++;

    if (lastLine && !endedByBreak) break;
    if (i >= pieces.length && endedByBreak) {
      // trailing hard break → emit a final empty line then stop
      const m = lineHeightFor([], p, fonts);
      outLines.push({ heightPt: m.heightPt, ascentPt: m.ascentPt, runs: [] });
      break;
    }
  }

  // Empty paragraph → one empty line at mark height.
  if (outLines.length === 0) {
    const m = lineHeightFor([], p, fonts);
    const runs: PositionedRun[] = [];
    if (labelPieces.length) { let lx = labelX; for (const lp of labelPieces) { runs.push(positionedRun(lp, lx)); lx += lp.widthPt; } }
    outLines.push({ heightPt: m.heightPt, ascentPt: m.ascentPt, runs });
  }

  return {
    lines: outLines,
    spaceBeforePt: p.ppr.spaceBeforePt,
    spaceAfterPt: p.ppr.spaceAfterPt,
    keepNext: p.ppr.keepNext,
    keepLines: p.ppr.keepLines,
    widowControl: p.ppr.widowControl,
    pageBreakBefore: p.ppr.pageBreakBefore,
  };
}
