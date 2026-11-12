/**
 * Tunable thresholds for PDF geometry paragraph reconstruction. EVERY threshold lives here as a named constant — no magic literals
 * in reconstruct.ts — so the heuristic is auditable and adjustable in one place.
 *
 * Factors are multipliers on the per-page MEDIAN font size (robust to a stray
 * heading/footnote), or on the derived line height, except where noted.
 */
export const RECONSTRUCT_CONFIG = {
  /**
   * Two items are on the SAME line when their text-baseline y differs by less
   * than this × median font size. ~0.3 tolerates sub/superscript jitter without
   * merging adjacent lines. (hasEOL is only a secondary tie-breaker — it is
   * producer-dependent, so y-clustering is primary.)
   */
  LINE_Y_TOLERANCE_FACTOR: 0.3,

  /**
   * A vertical gap between consecutive lines greater than this × line height
   * starts a NEW paragraph. Above normal single-spacing (~1.15–1.2) but below a
   * blank-line separation, so clause-per-line bodies split correctly while a
   * genuinely wrapped paragraph stays whole.
   */
  PARAGRAPH_GAP_FACTOR: 1.4,

  /**
   * An interior x-gap (empty across the WHOLE page, found by merging every
   * item's horizontal span) wider than this × median font size is a column
   * gutter. A normal inter-word space is < 1 font-size; a column gutter is much
   * wider, and crucially must be empty on every line to qualify.
   */
  COLUMN_GUTTER_MIN_FACTOR: 2.0,

  /**
   * Within a line, an x-gap between successive items greater than this × font
   * size inserts a space; a smaller gap concatenates directly (so sub-pixel
   * kerned word fragments rejoin into one word — P11).
   */
  INTRA_LINE_SPACE_FACTOR: 0.22,

  /** Items rotated more than this many degrees from horizontal are skipped (watermarks/diagonal stamps). */
  ROTATION_MAX_DEGREES: 6,

  /** A line whose font size exceeds median × this is treated as a heading. */
  HEADING_FONT_RATIO: 1.2,

  /** An ALL-CAPS line with at most this many words is treated as a heading. */
  HEADING_MAX_WORDS: 12,

  /** Fallback font size when a page has no measurable items. */
  DEFAULT_FONT_SIZE: 11,
} as const;

export type ReconstructConfig = typeof RECONSTRUCT_CONFIG;
