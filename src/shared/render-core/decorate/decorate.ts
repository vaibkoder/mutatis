/**
 * Tracked-change decoration (feature 16 §2.1) — host-free, pure.
 *
 * Maps native revision spans (already parsed onto runs) to vector decorations
 * using ONE simple configurable theme (ADR-0024: no gallery). Styling NEVER
 * filters — every revision run is decorated (ADR-0008 never-suppress). The
 * 8-class annotations (run.changeClasses), when present, are used only to pick a
 * colour family; absence still yields a by-type decoration.
 */
import type { LayoutResult, LayoutLine, PositionedRun, Decoration, DecoSeg, RenderTheme, Rect } from '../types.js';

function decorationsForRun(run: PositionedRun, theme: RenderTheme): Decoration[] {
  const rev = run.revision;
  if (!rev) return [];
  const out: Decoration[] = [];
  switch (rev.type) {
    case 'ins':
      out.push({ kind: 'insUnderline', color: theme.insertColor });
      if (theme.highlightInserts) out.push({ kind: 'highlight', color: theme.insertColor });
      break;
    case 'del':
      out.push({ kind: 'delStrike', color: theme.deleteColor });
      break;
    case 'moveFrom':
      out.push({ kind: 'delStrike', color: theme.moveColor });
      if (theme.changeAnchor) out.push({ kind: 'moveAnchor', color: theme.moveColor, glyph: theme.changeAnchor });
      break;
    case 'moveTo':
      out.push({ kind: 'insUnderline', color: theme.moveColor });
      if (theme.changeAnchor) out.push({ kind: 'moveAnchor', color: theme.moveColor, glyph: theme.changeAnchor });
      break;
  }
  return out;
}

/**
 * The continuous-decoration pass (conventional redline look). Inter-word spaces are
 * folded into x-gaps by the layout (paragraph.ts), so each WORD is its own run.
 * Painting decorations per run leaves the gaps undecorated → a distracting
 * per-word "gappy" strike/underline/highlight. Here we merge maximal spans of
 * *adjacent* runs that share the same visual decoration (kind + colour) into one
 * continuous rule (DecoSeg) / one rect, bridging the gaps. This is geometry only:
 * `run.decorations` still carries every run's per-type decoration + colour, so
 * never-suppress (ADR-0008) and the text-colour path are unchanged.
 */
function decorateLine(line: LayoutLine, theme: RenderTheme, rects: Rect[], absTopPt: number, seenAnchors: Set<string>): boolean {
  let changed = false;
  for (const run of line.runs) {
    if (!run.revision) continue;
    changed = true;
    run.decorations = decorationsForRun(run, theme);
  }
  if (!changed) return false;

  // A move anchor marks the move ONCE, not once per word. Inter-word spaces are
  // folded into x-gaps so every word is its own run; emitting the glyph per run
  // litters the moved block with '*'s. Keep the anchor only on the first run of
  // each move — deduped by changeId+side across the whole document, falling back
  // to the first run of a contiguous move span on the line when no changeId is
  // present. This drops only the redundant glyph; the strike/underline + colour
  // still decorate every run, so never-suppress (ADR-0008) is unaffected.
  let prevMoveKey: string | null = null;
  for (const run of line.runs) {
    const decs = run.decorations;
    const ai = decs?.findIndex((d) => d.kind === 'moveAnchor' || d.kind === 'changeAnchor') ?? -1;
    if (!decs || ai < 0 || !run.revision) { prevMoveKey = null; continue; }
    const rev = run.revision;
    const key = `${rev.type}:${rev.changeId ?? ''}`;
    const keep = rev.changeId ? !seenAnchors.has(key) : prevMoveKey !== key;
    if (keep) seenAnchors.add(key);
    else decs.splice(ai, 1);
    prevMoveKey = key;
  }

  // Merge adjacent runs with the same underline/strike into continuous rules.
  const segs: DecoSeg[] = [];
  let cur: DecoSeg | null = null;
  for (const run of line.runs) {
    const dec = run.decorations?.find((d) => d.kind === 'insUnderline' || d.kind === 'delStrike');
    if (!dec) { cur = null; continue; }
    const kind = dec.kind as DecoSeg['kind'];
    if (cur && cur.kind === kind && cur.color === dec.color) {
      cur.x2Pt = run.xPt + run.widthPt;
      cur.sizePt = Math.max(cur.sizePt, run.rpr.sizePt);
    } else {
      cur = { kind, color: dec.color, x1Pt: run.xPt, x2Pt: run.xPt + run.widthPt, sizePt: run.rpr.sizePt };
      segs.push(cur);
    }
  }
  if (segs.length) line.decoSegs = segs;

  // Merge adjacent highlighted (insertion) runs into one continuous tint rect.
  let hi: Rect | null = null;
  for (const run of line.runs) {
    if (run.decorations?.some((d) => d.kind === 'highlight')) {
      if (hi) hi.widthPt = run.xPt + run.widthPt - hi.xPt;
      else { hi = { xPt: run.xPt, yPt: absTopPt, widthPt: run.widthPt, heightPt: line.heightPt, fill: theme.insertColor + '22' }; rects.push(hi); }
    } else {
      hi = null;
    }
  }

  if (theme.changeBar) line.changeBar = true;
  return true;
}

/** Apply the theme to a laid-out result in place. Returns the count of changed runs. */
export function decorateLayout(layout: LayoutResult, theme: RenderTheme): number {
  let changedRuns = 0;
  // Shared across pages/footnotes so a move's anchor glyph appears exactly once.
  const seenAnchors = new Set<string>();
  for (const page of layout.pages) {
    for (const line of page.bodyLines) {
      for (const r of line.runs) if (r.revision) changedRuns++;
      decorateLine(line, theme, page.rects, line.topPt, seenAnchors);
    }
    if (page.footnotes) {
      for (const line of page.footnotes.lines) {
        for (const r of line.runs) if (r.revision) changedRuns++;
        decorateLine(line, theme, page.rects, line.topPt, seenAnchors);
      }
    }
  }
  return changedRuns;
}
