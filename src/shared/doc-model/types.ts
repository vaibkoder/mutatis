/** A single paragraph captured by the DocModel single-pass loader. */
export interface ParaNode {
  /** Stable paragraph ID derived from position (para-{index}). Re-resolved on reload. */
  id: string;
  /** Zero-based document order index. */
  index: number;
  /** Plain text content. */
  text: string;
  /** Named paragraph style (e.g. "Heading 1", "Normal", "ListParagraph"). */
  styleName: string;
  /** Rendered list string from listItem.listString (e.g. "4.2.1", "(b)"). Absent if not a list item. */
  listString?: string;
  /** List level (0-based). Absent if not a list item. */
  listLevel?: number;
  /**
   * Where this paragraph lives. Table-cell paragraphs remain in the `paragraphs`
   * array with their `body-para-{index}` id unchanged; only `container` and
   * `tableInfo` distinguish them.
   */
  container: 'body' | 'footnote' | 'tableCell';
  /** Table location when container === 'tableCell' (synthesised tableId, 0-based row/col). */
  tableInfo?: { tableId: string; row: number; col: number };
  /** Resolves a live Word.Range for this paragraph. Never hold across ctx.sync boundaries. */
  getRange(ctx: Word.RequestContext): Word.Range;
}

/** Snapshot of the whole document loaded by DocModel. */
export interface DocSnapshot {
  paragraphs: ReadonlyArray<ParaNode>;
  footnotes: ReadonlyArray<ParaNode>;
  loadedAt: number;
}
