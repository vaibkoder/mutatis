/**
 * Preview painter (feature 16 §2.1 Output A; ADR-0024) — main-thread DOM only.
 *
 * Paints the ONE shared `LayoutResult` (the same object the PDF backend consumes,
 * so the two paginate identically — acceptance criterion 7) into an absolutely-
 * positioned DOM tree: one page frame per `LayoutPage`, one `.rc-line` box per
 * line, one `<span>` per positioned run, with highlight/shading rects behind the
 * text. Coordinates are emitted in **points** (CSS `pt` units) so the painted
 * boxes equal the layout boxes 1:1.
 *
 * Framework-agnostic (plain DOM) so the host-free engine does not pull a UI
 * framework into `shared`; Phase 2 wraps this behind the Compare/preview Preact
 * surfaces. The tracked-change theme is one simple configurable CSS theme exposed
 * through CSS variables on the root — no gallery (ADR-0024). Styling never filters
 * (ADR-0008): every decorated run is painted with its decoration.
 *
 * Virtualisation: page *frames* (correct height) are always present so the scroll
 * height is right, but only pages within the viewport ± `overscanPages` have their
 * line content painted. When the container reports no height (e.g. jsdom, or before
 * first layout) every page is painted.
 */
import type {
  LayoutResult, LayoutPage, LayoutLine, PositionedRun, Rect, RenderTheme,
} from '../types.js';
import { DEFAULT_THEME, PAGE_INK } from '../types.js';

export interface PreviewOptions {
  theme?: RenderTheme;
  /** Pages beyond the viewport ± this many are not painted. Default 2. */
  overscanPages?: number;
  /** Font family used to DISPLAY text (metrics still come from the clones, so
   *  positions stay valid). Defaults to each run's requested family. */
  displayFontFamily?: string;
}

export interface PreviewHandle {
  /** The scroll container (append to the DOM yourself). */
  root: HTMLElement;
  /** Re-evaluate which pages are painted (call after scroll/resize). */
  refresh(): void;
  destroy(): void;
}

const STYLE_ID = 'rc-preview-style';
const SUP_SCALE = 0.8;
/** Text baseline as a fraction of run size above a bottom-anchored run (≈ typical font descent). */
const BASELINE_FRAC = 0.21;

/** Inject the preview stylesheet once per document (idempotent). */
function ensureStyle(doc: Document): void {
  if (doc.getElementById(STYLE_ID)) return;
  const style = doc.createElement('style');
  style.id = STYLE_ID;
  // `.rc-page` declares its own `color` so the host's foreground is not inherited
  // onto white paper. That declaration is NECESSARY and not sufficient: plain body
  // text is a `.rc-run` span with no inline colour (see `paintRun` — it colours
  // only `insUnderline`/`delStrike`), so it reaches this value by inheritance, and
  // any rule matching those spans outranks it. Both the declaration and the
  // variable it names have to be right. The failure is easy to miss either way,
  // because the tracked-change runs carry their own colour and stay readable
  // while the unchanged text goes near-white on white.
  style.textContent = `
.rc-preview { overflow-y: auto; background: var(--rc-canvas, #e7e8ea); padding: 16pt 0; }
.rc-page { position: relative; margin: 0 auto 16pt; background: var(--rc-page, #fff);
  color: var(--rc-ink, ${PAGE_INK});
  box-shadow: 0 1pt 4pt rgba(0,0,0,.18); box-sizing: border-box; overflow: hidden; }
.rc-line { position: absolute; white-space: pre; }
.rc-run { position: absolute; bottom: 0; line-height: 1; }
.rc-rect { position: absolute; }
.rc-deco { position: absolute; pointer-events: none; }
.rc-changebar { position: absolute; background: var(--rc-changebar, #808080); }
.rc-fn-sep { position: absolute; height: 0; border-top: 0.5pt solid var(--rc-fnsep, #666); }
.rc-anchor { font-size: 0.7em; vertical-align: super; }
`;
  (doc.head ?? doc.documentElement).appendChild(style);
}

function setThemeVars(root: HTMLElement, theme: RenderTheme): void {
  root.style.setProperty('--rc-ins', theme.insertColor);
  root.style.setProperty('--rc-del', theme.deleteColor);
  root.style.setProperty('--rc-move', theme.moveColor);
  root.style.setProperty('--rc-changebar', '#808080');
  // Paper ink, pinned the same way `--rc-changebar` above is. Custom properties
  // inherit, so anything in the host's cascade that defines `--rc-ink` would
  // otherwise reach the page; pinning it here is what stops that. Nothing else in
  // this repo defines it today, which makes the line look inert — that is a fact
  // about the current tree, not about the mechanism. Note this is NOT what keeps
  // the host's foreground out: that is the `color` declaration on `.rc-page`, and
  // specifically its naming `--rc-ink` rather than a host variable.
  root.style.setProperty('--rc-ink', PAGE_INK);
}

function pt(n: number): string { return `${n}pt`; }

function paintRun(doc: Document, run: PositionedRun, lineHeightPt: number, displayFont?: string): HTMLElement {
  const span = doc.createElement('span');
  span.className = 'rc-run';
  span.textContent = run.text;
  const s = span.style;
  s.left = pt(run.xPt);
  const sup = run.rpr.superscript || run.footnoteId !== undefined;
  const sub = run.rpr.subscript;
  s.fontSize = pt(sup || sub ? run.rpr.sizePt * SUP_SCALE : run.rpr.sizePt);
  s.fontFamily = displayFont ?? run.rpr.fontFamily;
  if (run.rpr.bold) s.fontWeight = 'bold';
  if (run.rpr.italic) s.fontStyle = 'italic';
  if (sup) s.bottom = pt(lineHeightPt * 0.33);
  else if (sub) s.bottom = pt(-lineHeightPt * 0.12);

  // Decorations: text colour + inline anchors here; underline/strike are drawn
  // once per merged segment (see paintLine → `.rc-deco`) so they stay continuous
  // across word gaps instead of breaking per word. Never filters (ADR-0008).
  const decs = run.decorations ?? [];
  const textDec: string[] = [];
  let colored = false;
  for (const d of decs) {
    if (d.kind === 'insUnderline' || d.kind === 'delStrike') { s.color = d.color; colored = true; }
    else if ((d.kind === 'moveAnchor' || d.kind === 'changeAnchor') && d.glyph) {
      const a = doc.createElement('sup'); a.className = 'rc-anchor'; a.textContent = d.glyph; a.style.color = d.color;
      span.appendChild(a);
    }
  }
  if (run.rpr.underline && !colored) textDec.push('underline');
  if (textDec.length) s.textDecoration = textDec.join(' ');
  return span;
}

function paintLine(doc: Document, line: LayoutLine, marginLeftPt: number, contentWidthPt: number, displayFont?: string): HTMLElement {
  const div = doc.createElement('div');
  div.className = 'rc-line';
  div.style.left = pt(marginLeftPt);
  div.style.top = pt(line.topPt);
  div.style.width = pt(contentWidthPt);
  div.style.height = pt(line.heightPt);
  for (const run of line.runs) {
    if (run.text.length === 0) continue;
    div.appendChild(paintRun(doc, run, line.heightPt, displayFont));
  }
  // Continuous underline/strike rules (merged spans — conventional redline look). Runs are
  // bottom-anchored (`bottom:0; line-height:1`), so the baseline sits ~BASELINE_FRAC
  // of the run size above the line bottom — leading-independent, so anchor the
  // rules to the bottom too (matches the glyphs; mirrors the PDF offsets ±0.12/0.28em).
  for (const seg of line.decoSegs ?? []) {
    const rule = doc.createElement('div');
    rule.className = 'rc-deco';
    rule.style.left = pt(seg.x1Pt);
    rule.style.width = pt(seg.x2Pt - seg.x1Pt);
    rule.style.height = pt(Math.max(0.5, seg.sizePt * 0.055));
    const bottom = seg.kind === 'insUnderline'
      ? (BASELINE_FRAC - 0.12) * seg.sizePt
      : (BASELINE_FRAC + 0.28) * seg.sizePt;
    rule.style.bottom = pt(bottom);
    rule.style.background = seg.color;
    div.appendChild(rule);
  }
  if (line.changeBar) {
    const bar = doc.createElement('div');
    bar.className = 'rc-changebar';
    bar.style.left = pt(-6);
    bar.style.top = '0';
    bar.style.width = '2pt';
    bar.style.height = pt(line.heightPt);
    div.appendChild(bar);
  }
  return div;
}

function paintRect(doc: Document, r: Rect, marginLeftPt: number, marginTopPt: number): HTMLElement {
  const div = doc.createElement('div');
  div.className = 'rc-rect';
  div.style.left = pt(marginLeftPt + r.xPt);
  div.style.top = pt(marginTopPt + r.yPt);
  div.style.width = pt(r.widthPt);
  div.style.height = pt(r.heightPt);
  if (r.fill) div.style.background = r.fill;
  if (r.stroke) div.style.border = `${r.strokeWidthPt ?? 0.5}pt solid ${r.stroke}`;
  return div;
}

/** Paint one page's content into its (already-created) frame. */
function paintPageContent(doc: Document, frame: HTMLElement, page: LayoutPage, displayFont?: string): void {
  const g = page.geometry;
  const cw = g.widthPt - g.marginLeftPt - g.marginRightPt;
  // rects first (behind text)
  for (const r of page.rects) frame.appendChild(paintRect(doc, r, g.marginLeftPt, g.marginTopPt));
  for (const line of page.bodyLines) {
    const div = paintLine(doc, line, g.marginLeftPt, cw, displayFont);
    div.style.top = pt(g.marginTopPt + line.topPt);
    frame.appendChild(div);
  }
  if (page.footnotes) {
    const sep = doc.createElement('div');
    sep.className = 'rc-fn-sep';
    sep.style.left = pt(g.marginLeftPt);
    sep.style.top = pt(g.marginTopPt + page.footnotes.separatorTopPt);
    sep.style.width = pt(cw * 0.32);
    frame.appendChild(sep);
    for (const line of page.footnotes.lines) {
      const div = paintLine(doc, line, g.marginLeftPt, cw, displayFont);
      div.style.top = pt(g.marginTopPt + line.topPt);
      frame.appendChild(div);
    }
  }
  frame.dataset.rcPainted = '1';
}

/**
 * Build the preview DOM. The returned `root` is a scroll container the caller
 * appends to the document; call `refresh()` on resize.
 */
export function renderPreview(layout: LayoutResult, opts: PreviewOptions = {}): PreviewHandle {
  const doc = (typeof document !== 'undefined' ? document : undefined);
  if (!doc) throw new Error('renderPreview requires a DOM (main thread only)');
  ensureStyle(doc);

  const theme = opts.theme ?? DEFAULT_THEME;
  const overscan = opts.overscanPages ?? 2;

  const root = doc.createElement('div');
  root.className = 'rc-preview';
  setThemeVars(root, theme);

  // One frame per page (always present → correct scroll height); content lazy.
  const frames: HTMLElement[] = layout.pages.map((page) => {
    const frame = doc.createElement('div');
    frame.className = 'rc-page';
    frame.style.width = pt(page.geometry.widthPt);
    frame.style.height = pt(page.geometry.heightPt);
    root.appendChild(frame);
    return frame;
  });

  const paint = (idx: number): void => {
    const frame = frames[idx];
    if (!frame || frame.dataset.rcPainted) return;
    paintPageContent(doc, frame, layout.pages[idx]!, opts.displayFontFamily);
  };
  const unpaint = (idx: number): void => {
    const frame = frames[idx];
    if (!frame || !frame.dataset.rcPainted) return;
    frame.replaceChildren();
    delete frame.dataset.rcPainted;
  };

  const refresh = (): void => {
    const view = root.clientHeight;
    // No measured height (jsdom / pre-layout) → paint everything.
    if (!view) { frames.forEach((_, i) => paint(i)); return; }
    const top = root.scrollTop;
    const bottom = top + view;
    for (let i = 0; i < frames.length; i++) {
      const f = frames[i]!;
      const fTop = f.offsetTop;
      const fBottom = fTop + f.offsetHeight;
      const overPx = overscan * f.offsetHeight;
      if (fBottom >= top - overPx && fTop <= bottom + overPx) paint(i);
      else unpaint(i);
    }
  };

  root.addEventListener('scroll', refresh, { passive: true });
  refresh();

  return {
    root,
    refresh,
    destroy(): void { root.removeEventListener('scroll', refresh); root.replaceChildren(); },
  };
}
