/**
 * Data model for Feature 08 — Compare.
 * All types are pure data (no Office JS dependencies); consumed by preprocess,
 * diff, classify, output, and UI layers.
 * Spec: spec/features/08-compare.md §4.
 */
import type { RowModel } from './diff/tables-core.js';

// ---------------------------------------------------------------------------
// §4.1  Paragraph record — one per paragraph per side, produced by preprocess
// ---------------------------------------------------------------------------

export interface ParaRecord {
  /** Position in its own document (0-based, body order). */
  index: number;
  /**
   * Source part: 'document', a header/footer ROLE id (`header:s0:default` —
   * kind, section, reference type; ADR-0047), 'footnotes', 'tableCell:rXcY' …
   */
  partId: string;
  /** Original <w:p> OOXML (revision markup quarantined out, after flatten). */
  rawXml: string;
  /**
   * Normalised text (§5.2) of what the reader sees in the paragraph (§5.0,
   * ADR-0045: run text, equation text, field results, text-box content read
   * once): case preserved; whitespace collapsed; NBSP→space; curly
   * quotes/dashes normalised; numbering prefix stripped.
   */
  normText: string;
  /**
   * Aggressively normalised text (also lowercased + all punctuation stripped).
   * Used ONLY to detect the `whitespacePunct` class — never for hashing.
   */
  normTextAggressive: string;
  /** Leading auto/manual numbering token stripped from normText ("4.2(a) ", "(iii) "). */
  numberingPrefix: string;
  /** 64-bit FNV-1a hash (hex string) of normText. */
  hash: string;
  /** w:pStyle value — for style-change detection. null if no explicit style. */
  styleId: string | null;
  /**
   * The paragraph's EFFECTIVE numbering (ADR-0042) — its own `<w:numPr>` or, far
   * more often in a firm template, the one it inherits through its style chain.
   * null = not numbered. **Absent (`undefined`) on the Text and PDF paths**,
   * which have no OOXML numbering to resolve; absent and null compare alike, so
   * a source that cannot know never reports a change it cannot see.
   *
   * NOT part of `hash`, and never to be folded into it: `hash` anchors the text
   * alignment, and a paragraph whose number moved must still match its
   * counterpart.
   */
  numbering?: ParaNumbering | null;
  isTableCell: boolean;
  /** Heuristic: proximity to EXECUTED/SIGNED/DEED/witness/Director/Secretary keywords (§7.2). */
  isSignatureBlock: boolean;
  isHeaderFooter: boolean;
  isFootnote: boolean;
  /**
   * TEXT-VS-TEXT MARKDOWN PATH ONLY (ADR-0027). When set, this record represents
   * a whole Markdown table as a single alignment UNIT: `hash`/`normText` cover the
   * table's content (so the structural diff aligns identical tables as equal),
   * `rawXml` holds the clean `<w:tbl>`, and this grid drives the structured table
   * redline author (`output/author-table.ts`).
   *
   * Since ADR-0046 the Word-vs-Word extractor sets it too, on a body table unit
   * that ALSO carries {@link ParaRecord.wordTable}: the grid is what lets every
   * ADR-0027 guard (no Level-2 word diff on the unit, no move candidacy, no
   * merge/split, a positional gap stretch, the table classifier branch) apply
   * unchanged. Whether a unit is Markdown or Word is decided by `wordTable`,
   * never by `table`.
   */
  table?: RowModel[];
  /**
   * WORD-VS-WORD BODY TABLES ONLY (ADR-0046). When set, this record is a whole
   * `<w:tbl>` as ONE alignment unit that keeps its XML: `rawXml` is the table,
   * re-emitted verbatim when the pair is equal; `hash` covers the row, cell and
   * paragraph boundaries and every cell paragraph's own hash, so equal means
   * identical text AND identical shape; this structure drives the per-cell
   * analysis and `output/author-word-table.ts`. Headers, footers and footnotes
   * never carry it.
   */
  wordTable?: WordTableModel;
}

/**
 * A Word table unit's structure (ADR-0046). Property XML is kept verbatim so
 * the author can re-emit it; content is the cell's own records, in document
 * order, each `index` its position in that cell.
 */
export interface WordTableModel {
  /** `<w:tblPr>…</w:tblPr>` as written ('' when absent). */
  tblPrXml: string;
  /** `<w:tblGrid>…</w:tblGrid>` as written ('' when absent). */
  tblGridXml: string;
  rows: WordTableRow[];
  /**
   * False when the table holds a wrapper this model does not represent: a
   * content control or custom XML round a row, a cell, or a cell's paragraphs.
   * The wrapper is read THROUGH — the rows, cells and paragraphs inside it are
   * in the model and the hash like any others, so an edit inside one is
   * reported. Equal is still verbatim; a CHANGED table is authored as the whole
   * original struck and the whole revised inserted, never re-assembled from a
   * model that would drop the wrapper.
   */
  supported: boolean;
}

export interface WordTableRow {
  /** `<w:tblPrEx>…</w:tblPrEx>` as written ('' when absent). */
  tblPrExXml: string;
  /** `<w:trPr>…</w:trPr>` as written ('' when absent). */
  trPrXml: string;
  cells: WordTableCell[];
  /** Row hash: the cells' hashes and layout, in order. */
  hash: string;
}

export interface WordTableCell {
  /** `<w:tcPr>…</w:tcPr>` as written ('' when absent). */
  tcPrXml: string;
  /** `w:gridSpan` (1 when absent). */
  gridSpan: number;
  /** `w:vMerge`: 'restart', 'continue' (the element with no value), or null. */
  vMerge: 'restart' | 'continue' | null;
  /** The cell's paragraphs and nested table units, document order, `index` = position. */
  content: ParaRecord[];
  /** Cell hash: content boundaries + each content record's hash (layout NOT included). */
  hash: string;
}

// ---------------------------------------------------------------------------
// §4.2  Diff result model
// ---------------------------------------------------------------------------

/** Structural relationship between a paragraph pair or run in the level-1 diff. */
/**
 * A paragraph's resolved numbering (ADR-0042). Built in two steps because the
 * label is a property of the SEQUENCE, not of the paragraph: `numId`/`ilvl`/`key`
 * come from the style chain at extraction, `label`/`path` from the shared
 * counter machine once the whole part is in document order.
 */
export interface ParaNumbering {
  /**
   * `<w:numId>` — a DOCUMENT-LOCAL identifier. **Never compare it across two
   * documents.** Word reassigns these freely: a Save-As onto another template, a
   * list pasted from elsewhere, a counterparty re-creating the file, all
   * renumber the `w:num` instances while every clause keeps its number.
   * Comparing it directly reported six changes on a document with none.
   * It is kept because the counter machine needs it to group counters WITHIN a
   * document — and for nothing else.
   */
  numId: string;
  /** 0-based level. Comparable across documents: level 2 is level 2 anywhere. */
  ilvl: number;
  /**
   * The comparable identity: the level, plus the shape of the level definitions
   * this paragraph's number is RENDERED FROM. Two paragraphs have the same
   * numbering identity when they sit at the same level of a list that renders
   * that level the same way — a statement about the documents, not about Word's
   * internal ids.
   *
   * `numberingKey` in `preprocess/pararecord.ts` is the definition of record and
   * explains every inclusion and every exclusion, each of which was a measured
   * defect: it is NOT the whole level definition (an ancestor's `lvlText` and
   * `isLgl` are never read for a descendant), NOT only the paragraph's own level
   * (an ancestor's format and counter are), NOT only the levels the template
   * substitutes (deleting any level 0…`ilvl` moves the counters below it), and
   * NOT only what `renderListNumbers` reads — the paragraph's own level carries
   * its DECLARED `numFmt` as well as the effective one (when the template
   * substitutes that level; an own format nothing substitutes moves no label),
   * because that renderer folds `isLgl` over the own token where Word keeps the
   * declared format, and a key that mirrored the fold hid the edit that changed
   * it. Read that function before changing this field's meaning.
   *
   * It covers what the renderer reads and writes WHEN RENDERING THIS PARAGRAPH,
   * and so stops at `ilvl`. Two limits follow, both recorded in ADR-0042 §
   * Accepted limits and both pinned by tests rather than papered over: a clause
   * moved between two lists whose definitions are identical does not change this
   * key (cross-document that is indistinguishable from Word having reassigned the
   * instance id, and inventing a change on every Save-As is the worse error); and
   * a level definition deeper than `ilvl` can move this paragraph's label without
   * moving its key, which is reported — `numberLabel` is compared from the two
   * rendered labels — but reported as a bystander of a sibling that did not move.
   */
  key: string;
  /** Rendered label ("(vi)", "Clause 1.01"); null when the level does not resolve. */
  label: string | null;
  /**
   * The counter values behind `label`, levels 0…`ilvl` — `[2, 1]` for `"2.1"`.
   * A rendered label cannot be subtracted, and §6.1 groups a renumbering cascade
   * by its uniform SHIFT, so this is what makes a numbering cascade collapsible.
   *
   * The whole path rather than this paragraph's own counter, because under the
   * dominant SG/E&W `%1.%2` style a sub-clause's number moves when an ANCESTOR's
   * counter moves and its own does not: `1.1` → `2.1` is a shift of +1 at level
   * 0 and of 0 at level 1. Reading the own-level counter alone measured every
   * such cascade as a shift of zero, which §6.1 drops — one uncollapsed row per
   * sub-clause in the document.
   */
  path: readonly number[] | null;
  /**
   * The levels 0…`ilvl` whose counters this paragraph's label actually SHOWS.
   * Two filters, not one: the `lvlText` must SUBSTITUTE the level (`%N`), and
   * that level's effective format must RENDER something. `[0,1]` for `%1.%2`;
   * `[1]` for `(%2)`; `[1]` again for `%1.%2` whose level 0 is `numFmt="none"`,
   * which fills its slot with nothing. `renderedLevels` in
   * `preprocess/pararecord.ts` is the definition of record.
   *
   * §6.1's cascade header names a level and a shift, and the rows beneath it
   * show the two labels. Measuring the shift at the first differing level of
   * `path` makes those two disagree whenever the level that moved is one the
   * template never prints: splitting a clause list under `(%2)` moved three
   * sub-clauses `(c)`→`(a)`, `(d)`→`(b)`, `(e)`→`(c)` — minus two each — under a
   * header reading *"shifted (+1)"*, which is the parent counter that `(%2)`
   * does not render. The same contradiction came back one filter later, from a
   * SUBSTITUTED level at `numFmt="none"` whose counter moved and printed
   * nothing. The shift is therefore measured at the shallowest level the label
   * prints. Same rule as `key`: what this paragraph's number is rendered from,
   * and nothing else.
   */
  renders: readonly number[] | null;
}

export type OpType =
  | 'equal'    // identical paragraph — no output markup
  | 'insert'   // present only in Revised
  | 'delete'   // present only in Original
  | 'replace'  // matched paragraph pair with intra-paragraph word changes
  | 'move'     // verbatim-identical paragraph at a different index (Level 3)
  | 'property'; // identical TEXT, differing non-text property (ADR-0042) or zero-text object (ADR-0045 § Amendment)

/** What a zero-text object is, as the change list and the report name it. */
export type ZeroTextSort = 'footnote reference' | 'endnote reference' | 'picture' | 'text box' | 'embedded object';

/**
 * An object that is part of what the reader sees and contributes no text — so
 * a paragraph whose only change is one of them has the same text, and so the
 * same hash, on both sides, and pairs as `equal` (ADR-0045 § Amendment).
 *
 * `at` is its place: how many characters of the paragraph's `normText` come
 * before it — measured in the normalised, prefix-stripped text, so a whitespace
 * run written differently, or a manual number that moved, cannot move it. In
 * the manual number rather than after it, a code: `-2` before the number, `-1`
 * inside it — so moving across the number is a move and renumbering is not.
 *
 * Nothing that says WHICH footnote, image or object it is: a footnote's `w:id`
 * and a picture's relationship id are document-local and Word reassigns them
 * (one footnote inserted early renumbers every one after it), so comparing them
 * would report a change at every footnote of a document that gained one.
 */
export interface ZeroTextObject {
  sort: ZeroTextSort;
  at: number;
  /**
   * Inside an equation or text box: the ordinal (document order, within this
   * paragraph) of the innermost equation, text-box content or text-box
   * paragraph holding it. Absent at the paragraph's own level. A collapsed
   * space can give "just before the box" and "just inside its start" the same
   * `at`; they are different places.
   */
  within?: number;
}

/** One side of a {@link PropertyChange}: the values the surfaces print. */
export interface PropertySide {
  styleId: string | null;
  numberLabel: string | null;
  ilvl: number | null;
  /** The paragraph's zero-text objects, in order. Set when `objects` is. */
  objects?: readonly ZeroTextObject[];
}

/**
 * Which non-text properties differ across a matched, text-identical pair
 * (ADR-0042). Set on `property` ops only; it is what the taxonomy reads to tell
 * a re-levelled clause (`formatting`, §6 row 8 / §9.9) from a clause whose
 * number merely shifted because a sibling moved (`crossRefRenumber`, §9.4).
 */
export interface PropertyChange {
  /** `w:pStyle` differs. */
  styleId: boolean;
  /**
   * `ParaNumbering.key` differs — the paragraph itself was re-levelled, or the
   * definition of its level was changed. **Not** `(numId, ilvl)`: a `w:num` id is
   * document-local and comparing it across two packages reports changes that did
   * not happen. See `ParaNumbering.key`.
   */
  numberingIdentity: boolean;
  /**
   * The rendered label differs (may be true with `numberingIdentity` false — a
   * cascade). Compared only when BOTH sides resolved a label: one side having
   * none is evidence that its level definition did not resolve, not that a
   * number moved. `numberingIdentity` still covers a paragraph that gained or
   * lost numbering. Spec §5.1.
   */
  numberLabel: boolean;
  /**
   * The VALUES that moved, so the row preview, the navigation anchor and the
   * exported report can say what changed. A property op has no `wordDiffs` —
   * every one of those surfaces reads `wordDiffs` — so without this the change
   * is counted and then shown as a blank row that does nothing when clicked.
   * Facts, not a sentence: the formatting belongs to each surface.
   */
  from: PropertySide;
  to: PropertySide;
  /**
   * The ordered zero-text objects differ (ADR-0045 § Amendment): a footnote or
   * endnote reference, a picture, a text box or an embedded object was added,
   * removed or moved, and nothing else in the text was. Absent means false.
   * Unlike every other property here this is a change to the CONTENT, so it is
   * classified `substantiveText`, is never a cascade, and is the one property
   * the redline marks (`output/author.ts`).
   */
  objects?: boolean;
  /**
   * Set ONLY when this is a pure cascade — the label moved and nothing about the
   * paragraph itself did. Carries what §6.1's collapse needs (`cascade.ts` reads
   * `wordDiffs`, and a property op has none): the two labels for the example,
   * the shift so a +1 cascade is not grouped with a +2 one, and the level the
   * shift happened AT so two levels' cascades do not merge.
   *
   * `shiftedIlvl` is the level whose counter moved, which is **not** necessarily
   * this paragraph's own level: `1.1` → `2.1` is a level-1 paragraph carried by
   * a level-0 shift. Grouping on the paragraph's own level would split one
   * insertion's cascade into one group per level it touched, and — worse — the
   * own-level shift is zero there, so there would be no group at all.
   */
  cascade?: { from: string; to: string; shiftedIlvl: number; shift: number };
}

/** Atomic-token kind for a WordDiff token. */
export type AtomicKind =
  | 'definedTerm'
  | 'clauseRef'
  | 'amount'
  | 'date'
  | 'percent'
  | 'duration';

/** One unit from the Level-2 word diff (diff-match-patch word mode). */
export interface WordDiff {
  op: 'eq' | 'ins' | 'del';
  text: string;
  /** Set when this token must not be split mid-token (§5.3). */
  atomic?: AtomicKind;
}

/** §6 change-classification taxonomy. Priority: structural > obligation > numerical > … > formatting. */
export type ChangeClass =
  | 'structural'
  | 'obligationLanguage'
  | 'numericalValue'
  | 'definedTerm'
  | 'crossRefRenumber'
  | 'substantiveText'
  | 'whitespacePunct'
  | 'formatting';

/** One operation in the ordered diff result, enriched by classify (B1/B2/B3) and author (C1). */
export interface DiffOp {
  /** Stable identifier for UI navigation. */
  id: string;
  type: OpType;
  /** ParaRecord indices in Original (null for pure inserts). */
  origRange: { start: number; end: number } | null;
  /** ParaRecord indices in Revised (null for pure deletes). */
  revRange: { start: number; end: number } | null;
  partId: string;
  /** Present when type === 'replace'. */
  wordDiffs?: WordDiff[];
  /** Present when type === 'property' — which properties differ (ADR-0042). */
  propertyChange?: PropertyChange;
  /** Primary (priority-resolved) classification — set after B1/B2. */
  classification: ChangeClass;
  /** All applicable classifications (multi-class — a change can match several). */
  classes: ChangeClass[];
  /** Injected w:author (from CompareConfig.author). */
  author: string;
  /** ISO-8601 timestamp injected by the driver — not sourced from Date.now() in pure functions. */
  date: string;
  /** Set by B3 if this op is collapsed into a renumbering-cascade annotation. */
  collapsedInto?: string;
  /** For 'move': the id of the paired DiffOp (the other end of the move). */
  moveTargetId?: string;
  /** Set by B2 when the change is a known en-US→en-GB spelling variant (§7.1). */
  spellingVariant?: { usSpelling: string; gbSpelling: string };
  /**
   * Set by the PDF fuzzy-anchor pass (feature 12, B3) when this op was paired by
   * high-similarity fallback rather than exact-hash anchoring — a low-confidence
   * match the UI badges. Never used to suppress (ADR-0008).
   */
  fuzzyAnchored?: boolean;
  /**
   * Generic low-confidence marker (feature 12) — set by the PDF fuzzy-anchor
   * pass and the PDF approximate table-region cell diff (B4). The UI badges it;
   * it is NEVER a suppression filter (ADR-0008).
   */
  lowConfidence?: boolean;
  /**
   * Set on an op that pairs two Word table units (ADR-0046): how the rows and
   * cells were aligned, and each matched cell's own ops from the body's stages.
   * The author re-assembles the table from this; nothing else may re-derive it.
   */
  tableDiff?: TableDiff;
}

/** The row and cell alignment of a Word table pair (ADR-0046). */
export interface TableDiff {
  rows: TableRowSlot[];
  /**
   * Set when the grid's width changed and at least one paired row is full width
   * (one unmerged cell per grid column) on both sides: the column map, read from
   * those rows, and the authored grid, column by column — a kept or added column
   * takes the revised width, a removed one the original's. Rows that cannot take
   * it read as out and in. Absent ⇒ the revised `tblGrid` as is.
   */
  columns?: TableColumnSlot[];
}

/** One column of the authored grid when columns changed (ADR-0046). */
export type TableColumnSlot =
  | { kind: 'match'; o: number; r: number }
  | { kind: 'insert'; r: number }
  | { kind: 'delete'; o: number };

/**
 * One row of the authored table. `match` rows have the same cell layout on both
 * sides, or take the column map (`TableDiff.columns`), and are revised cell by
 * cell; a row whose layout changed otherwise (a merge, a split, a vertical
 * merge, a merged row a column passes through) is a `delete` of the original
 * row plus an `insert` of the revised one.
 */
export type TableRowSlot =
  | { kind: 'match'; o: number; r: number; cells: TableCellSlot[] }
  | { kind: 'insert'; r: number }
  | { kind: 'delete'; o: number };

/**
 * One cell of a matched row. `match` carries the cell content's ops (indices
 * into the two cells' `content`). `insert`/`delete` are a column added or
 * removed (`w:cellIns`/`w:cellDel`).
 */
export type TableCellSlot =
  | { kind: 'match'; o: number; r: number; ops: DiffOp[] }
  | { kind: 'insert'; r: number }
  | { kind: 'delete'; o: number };

// ---------------------------------------------------------------------------
// §4.3  Cascade annotation (from §6.1 renumbering-cascade collapse)
// ---------------------------------------------------------------------------

export interface CascadeAnnotation {
  id: string;
  /** Which reference token kind is cascading (config referenceTokens). */
  tokenKind: string;
  /** Uniform numeric shift across all members (+1, -1, etc.). */
  shift: number;
  /** Number of raw crossRefRenumber changes collapsed. */
  affectedCount: number;
  /** Up to three human-readable examples ("Clause 4.2 → 4.3"). */
  exampleRefs: string[];
  /**
   * What cascaded: cross-REFERENCES to a clause whose number moved, or the
   * clause NUMBERS themselves (ADR-0042). Both are §6.1 cascades with the same
   * cause and the same collapse; they are different sentences to a reader, and
   * grouping them together would claim a document referenced a clause it only
   * renumbered. Absent ⇒ `crossReference` (every cascade before ADR-0042).
   */
  source?: 'crossReference' | 'numbering';
}

// ---------------------------------------------------------------------------
// §4.3  Change-summary buckets (feeds the Findings/summary panel)
// ---------------------------------------------------------------------------

export interface ChangeSummary {
  total: number;
  byCategory: Record<ChangeClass, number>;
  renumberingCascades: CascadeAnnotation[];
  movedParagraphs: number;
  spellingVariantCount: number;
  quarantinedCommentsOriginal: number;
  quarantinedCommentsRevised: number;
  /**
   * §9.7/§9.8 — changes found in the parts diffed OUTSIDE the body: headers,
   * footers, footnotes. Headers and footers are zero at the default config,
   * where they are not processed AT ALL — not processed and then hidden.
   * **Footnotes are not**: `compare.footnotes` defaults to `'compare'` (§7.5),
   * because §9.8 makes a footnote content change substantive, so this count can
   * be non-zero with nothing switched on.
   *
   * These counts exist because the driver computed them and threw them away:
   * `compare.excludeHeaders: false` used to do nothing whatsoever. The
   * individual part ops are deliberately NOT merged into the navigable change
   * list — `goToChange` resolves a body paragraph index, and a header op's
   * range is an index into a different part, so merging them would mis-navigate.
   * They ARE counted into `total` and `byCategory` (`mergePartsIntoSummary`),
   * because §8.2 makes `total` the full count and §9.8 makes a footnote content
   * change substantive; only the navigation half is outstanding.
   */
  partChanges: { headers: number; footers: number; footnotes: number };
  /**
   * §9.9 — a style DEFINITION change (styles.xml itself differs), as a note for
   * the summary panel; null when unchanged. Distinct from a per-paragraph
   * styleId change, which is a `property` op in the body (ADR-0042).
   */
  styleDefinitionNote: string | null;
  filtersActive: {
    hideFormattingOnly: boolean;
    collapseRenumbering: boolean;
  };
}
