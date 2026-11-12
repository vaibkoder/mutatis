/** The normalised lookup key for a section (e.g. "clause:4.2", "schedule:2/para:3", "recital:A"). */
export type SectionKey = string;

export type SectionKind =
  | 'clause'
  | 'schedule'
  | 'recital'
  | 'part'
  | 'annex'
  | 'appendix'
  | 'unknown';

export interface SectionNode {
  key: SectionKey;
  kind: SectionKind;
  /** Human-readable number/letter (e.g. "4.2", "(b)", "B", "II"). */
  rawNumber: string;
  /** Display label (e.g. "Clause 4.2", "Schedule 1", "Recital (A)"). Configured noun. */
  displayLabel: string;
  /** Parent section key, if any. */
  parentKey?: SectionKey;
  /** Depth in the document hierarchy (0 = top-level clause/schedule). */
  level: number;
  /** The DocModel paragraph ID. */
  paraId: string;
  /** Whether this section uses static (manually-typed) numbering — not list numbering. */
  isStaticNumbered: boolean;
  /**
   * True when the sub-clause token is a single character valid as both a roman numeral
   * (i, v, x, l, c, d, m) and a letter, and sibling context was insufficient to resolve.
   * The reference matcher should accept either interpretation.
   */
  ambiguousRomanLetter?: boolean;
  /**
   * Bookmarks anchored at or near this unit (populated by the feature-local
   * REF-field/bookmark pass — ADR-0014). Default empty array.
   */
  bookmarkNames: string[];
  /** Lazy range resolver. */
  getRange(ctx: Word.RequestContext): Word.Range;
}

/** Spec §4.2 SectionIndexEntry — carries the node plus any future per-entry metadata. */
export interface SectionIndexEntry {
  node: SectionNode;
  /** Disambiguation: when multiple nodes share the same key, all appear here → Ambiguous. */
}

export interface SectionIndex {
  /** All section nodes in document order. */
  nodes: ReadonlyArray<SectionNode>;
  /**
   * Flat lookup map: normalised key → array of nodes.
   * An array with >1 entry means the key is Ambiguous.
   */
  byKey: ReadonlyMap<SectionKey, SectionNode[]>;
  /** Lookup by paragraph ID. */
  byParaId: ReadonlyMap<string, SectionNode>;
  /** All nodes in document order (same as `nodes`; named alias for spec §4.2 interface). */
  ordered: ReadonlyArray<SectionNode>;
  /** Returns true when `key` has at least one entry. */
  has(key: SectionKey): boolean;
  /** Resolve a key: [] = not found, [x] = unique, [x,y…] = ambiguous. */
  resolve(key: SectionKey): SectionNode[];
}
