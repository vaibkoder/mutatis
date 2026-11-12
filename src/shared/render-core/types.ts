/**
 * RenderCore — public type model (feature 16, ADR-0024).
 *
 * Host-free, DOM-free. Three layers:
 *   1. RenderModel        — resolved input: blocks, footnotes, page geometry
 *                           (output of `input/` parsing + style cascade).
 *   2. LayoutResult       — ONE shared layout result: pages → lines → positioned
 *                           runs (+ decorations + rects). Painted two ways
 *                           (preview, PDF) so they are pixel-locked.
 *   3. Theme / annotations— tracked-change styling config + RenderCore-owned
 *                           change classes (Compare maps into these; RenderCore
 *                           never imports `@features/*` — keeps shared↛features).
 *
 * All measurements are in **points** (1pt = 1/72in) unless named `*Fu` (font units).
 */

// ─── Geometry ────────────────────────────────────────────────────────────────

/** A4 / sectPr page box. Margins are the Word content-area insets. */
export interface PageGeometry {
  widthPt: number;
  heightPt: number;
  marginTopPt: number;
  marginRightPt: number;
  marginBottomPt: number;
  marginLeftPt: number;
}

/** Content-area width available for body text. */
export function contentWidthPt(g: PageGeometry): number {
  return g.widthPt - g.marginLeftPt - g.marginRightPt;
}
/** Content-area height available for body text (before footnote reservation). */
export function contentHeightPt(g: PageGeometry): number {
  return g.heightPt - g.marginTopPt - g.marginBottomPt;
}

// ─── Effective run/paragraph formatting (post style-cascade) ───────────────────

export type Justification = 'left' | 'right' | 'center' | 'both';
export type LineRule = 'auto' | 'exact' | 'atLeast';

/** Resolved run properties (effective rPr after the cascade + toggle semantics). */
export interface EffectiveRPr {
  /** Requested font family (pre clone-mapping; the font subsystem maps it). */
  fontFamily: string;
  sizePt: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** vertAlign w:superscript / w:subscript. */
  superscript: boolean;
  subscript: boolean;
}

/** Resolved paragraph properties (effective pPr after the cascade). */
export interface EffectivePPr {
  justification: Justification;
  /** Left/right text indents (pt). */
  indentLeftPt: number;
  indentRightPt: number;
  /** First-line indent (pt, positive). Mutually exclusive with hangingPt in OOXML. */
  firstLinePt: number;
  /** Hanging indent (pt, positive): first line starts at indentLeft, wrapped lines at indentLeft+hanging. */
  hangingPt: number;
  spaceBeforePt: number;
  spaceAfterPt: number;
  lineRule: LineRule;
  /** For lineRule 'auto': multiple ×240 (Word). For exact/atLeast: twips. 0 ⇒ single. */
  lineValue: number;
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
}

// ─── Tracked-change (revision) model ───────────────────────────────────────────

export type RcRevisionType = 'ins' | 'del' | 'moveFrom' | 'moveTo';

/** RenderCore's own change taxonomy — mirrors Compare's 8 classes by name only. */
export type RcChangeClass =
  | 'structural'
  | 'obligationLanguage'
  | 'numericalValue'
  | 'definedTerm'
  | 'crossRefRenumber'
  | 'substantiveText'
  | 'whitespacePunct'
  | 'formatting';

export interface RcRevision {
  type: RcRevisionType;
  /** Stable change id (used to pair moveFrom/moveTo and to anchor the change list). */
  changeId?: string;
  author?: string;
  date?: string;
}

// ─── Parsed content (pre-layout) ───────────────────────────────────────────────

export interface RcRun {
  text: string;
  rpr: EffectiveRPr;
  /** Set when this run sits inside a native <w:ins>/<w:del>/<w:moveFrom*>/<w:moveTo*>. */
  revision?: RcRevision;
  /** Optional class annotations (styling only; never filters — ADR-0008). */
  changeClasses?: readonly RcChangeClass[];
  /** This run is a footnote reference mark (superscript), linking to footnoteId. */
  footnoteId?: string;
  /** This run stands in for an image/object Mutatis does not paint in v1. */
  imagePlaceholder?: boolean;
}

export interface RcParagraph {
  runs: RcRun[];
  ppr: EffectivePPr;
  /** Paragraph-mark run properties (default for empty paras + line-height fallback). */
  markRpr: EffectiveRPr;
  styleId: string | null;
  isHeading: boolean;
  /** Rendered numbering label ("6.1", "(a)") from the numbering resolver, if any. */
  numberLabel?: string;
  numberLevel?: number;
}

export interface RcCell {
  blocks: RcBlock[];
  /** Horizontal cell span (gridSpan). 1 if absent. */
  gridSpan: number;
}
export interface RcRow {
  cells: RcCell[];
  cantSplit: boolean;
}
export interface RcTable {
  rows: RcRow[];
  /** Column widths in pt from tblGrid (sum ≈ table width). */
  gridColsPt: number[];
}

export type RcBlock =
  | { kind: 'paragraph'; paragraph: RcParagraph }
  | { kind: 'table'; table: RcTable };

export interface RcFootnote {
  id: string;
  /** Footnote body content (paragraphs; tables in footnotes are rare/deferred). */
  blocks: RcBlock[];
}

/** The resolved render input — output of `input/` parsing. */
export interface RenderModel {
  blocks: RcBlock[];
  footnotes: Map<string, RcFootnote>;
  geometry: PageGeometry;
}

// ─── Layout result (one shared result → both painters) ─────────────────────────

export type DecorationKind =
  | 'insUnderline'
  | 'delStrike'
  | 'highlight'
  | 'changeBar' // left-margin vertical bar
  | 'moveAnchor'
  | 'changeAnchor';

/** A vector decoration to paint under/over a run or in the margin. */
export interface Decoration {
  kind: DecorationKind;
  /** Colour as a hex string (theme-derived), e.g. '#c00'. */
  color: string;
  /** Optional inline anchor glyph (e.g. '*') for change/move anchors. */
  glyph?: string;
}

/** A shaped, positioned run on a line. */
export interface PositionedRun {
  /** x offset from the content-area left edge (pt). */
  xPt: number;
  widthPt: number;
  text: string;
  rpr: EffectiveRPr;
  revision?: RcRevision;
  changeClasses?: readonly RcChangeClass[];
  decorations?: Decoration[];
  footnoteId?: string;
  imagePlaceholder?: boolean;
}

/**
 * A continuous underline/strike rule spanning one OR MORE adjacent same-revision
 * runs on a line. Painted as a single vector rule (PDF) / div (preview) so the
 * decoration bridges the inter-word gaps instead of breaking per word — the
 * conventional redline look (decoration geometry only; never affects membership, so
 * ADR-0008 never-suppress is untouched — every run is still painted).
 */
export interface DecoSeg {
  kind: 'insUnderline' | 'delStrike';
  /** Start/end x offset from the content-area left edge (pt). */
  x1Pt: number;
  x2Pt: number;
  color: string;
  /** Representative run size (pt; max over the merged group) for thickness + baseline offset. */
  sizePt: number;
}

export interface LayoutLine {
  /** Top of the line box, y offset from the content-area top edge (pt). */
  topPt: number;
  heightPt: number;
  /** Baseline offset from the line-box top (pt). */
  ascentPt: number;
  runs: PositionedRun[];
  /** True if a left-margin change bar should be drawn for this line. */
  changeBar?: boolean;
  /** Merged continuous underline/strike rules (set by the decorate pass). */
  decoSegs?: DecoSeg[];
}

export interface FootnoteRegion {
  /** y (from content-area top) of the separator rule. */
  separatorTopPt: number;
  lines: LayoutLine[];
  continuedFromPrev: boolean;
  continuesToNext: boolean;
}

/** A filled rectangle (highlight, table border/shading) in content-area coords. */
export interface Rect {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
  fill?: string;
  stroke?: string;
  strokeWidthPt?: number;
}

export interface LayoutPage {
  geometry: PageGeometry;
  bodyLines: LayoutLine[];
  footnotes?: FootnoteRegion;
  rects: Rect[];
}

export interface LayoutResult {
  pages: LayoutPage[];
}

// ─── Theme ─────────────────────────────────────────────────────────────────────

/** One simple, configurable tracked-change theme (no gallery — ADR-0024). */
export interface RenderTheme {
  insertColor: string;
  deleteColor: string;
  moveColor: string;
  /** Highlight insertions (background) in addition to the underline. */
  highlightInserts: boolean;
  /** Draw a left-margin change bar beside changed lines. */
  changeBar: boolean;
  /** Inline anchor glyph for moves/changes (e.g. '*'); empty disables. */
  changeAnchor: string;
}

export const DEFAULT_THEME: RenderTheme = {
  insertColor: '#0a7d00',
  deleteColor: '#c00000',
  moveColor: '#6a1b9a',
  highlightInserts: true,
  changeBar: true,
  changeAnchor: '*',
};

/** The page's own text colour — paper ink, never the host theme's foreground. */
export const PAGE_INK = '#1a1a1a';
