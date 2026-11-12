/**
 * Feature 06 — Numbering.
 * Re-exports the canonical scheme types from shared/config (the single home for
 * preset data) and adds feature-local operational types.
 */

export type {
  NumberingPreset as NumberingScheme,
  NumberingLevel,
  NumberingConfig,
  ListNumberingStyle,
  LevelLabelStyle,
} from '@shared/config/index.js';

/** Scope of a scheme apply operation. */
export type ApplyScope = 'selection' | 'document';

/** Summary returned by the apply engine after applying a scheme. */
export interface ApplyReport {
  /** Number of lists affected. */
  listsAffected: number;
  /** Level indices that could not be applied (per-level error tolerance, §3.1). */
  degradedLevels: number[];
  /** Number of table-cell paragraphs excluded (§8.1). */
  tableParasSkipped: number;
  /** Number of footnote paragraphs skipped (§8.2). */
  footnoteParasSkipped: number;
  /** Number of statically-numbered paragraphs reported as skipped (§8.3 / ADR-0006). */
  staticParasSkipped: number;
  /** True when the document has multiple Word sections (§8.4 restart warning). */
  hasMultipleSections: boolean;
  /** True when a style-linked list was detected and skipped (§8.7). */
  styleLinkedSkipped: boolean;
}

/** State of the post-apply/promote/demote cross-reference review (§5.3). */
export interface ReviewState {
  /** Total candidate references highlighted. */
  count: number;
  /**
   * Highlighted text snippets for the review panel.
   * Each entry is the text of a matched reference range.
   */
  items: ReviewItem[];
  /** Whether a review has been run since the last apply/level-change. */
  active: boolean;
}

export interface ReviewItem {
  /** The matched reference text, e.g. "clause 4.2(a)". */
  text: string;
  /** Surrounding context (up to ~80 chars either side). */
  context: string;
  /** Paragraph ID from DocModel — used to navigate. */
  paraId: string;
  /** Character offset within the paragraph for range resolution. */
  startOffset: number;
  endOffset: number;
}

/** Boundary state for promote/demote buttons per selection. */
export interface BoundaryState {
  canPromote: boolean;
  canDemote: boolean;
}

/** Collision resolution choice when importing a scheme whose id already exists. */
export type CollisionResolution = 'replace' | 'keepBoth';
