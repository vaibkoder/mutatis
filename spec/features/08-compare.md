# 08 — Compare (Compare tab)

> **Status:** ✅ v1 (Word-vs-Word) · 🔜 v2 (Word-vs-PDF, fuzzy moves). Primary consumer of the [OOXML Service](../ARCHITECTURE.md#the-ooxml-service).
>
> Heading scheme: `#` doc title, `##` numbered top sections, `###` sub-sections.

---

## 1. Purpose & user value

Lawyers spend a large fraction of every transaction reconciling versions: the counterparty returns a marked-up draft, a colleague sends a "clean" copy that is not actually clean, or a 300-page credit agreement comes back with changes buried across forty schedules. Microsoft Word ships a native **Compare** feature, but it is **not exposed in the Office JS API surface** ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 17) — an add-in cannot invoke it. Firms that need richer, classified, jurisdiction-aware comparison today reach for server-side tools that require uploading the document off-machine — a confidentiality and (in SG/UK) often a regulatory problem.

The **Compare tab** reproduces a Word-quality legal redline **entirely inside the task pane** (see [Client-Side Model](../ARCHITECTURE.md#client-side-model)), with three pieces of value the native tool does not provide:

1. **A true tracked-changes redline** between two `.docx` versions, written as native `<w:ins>`/`<w:del>` markup so the output opens in any Word install and the lawyer can **accept/reject each change in the standard Word UI** — no proprietary viewer, no lock-in.
2. **Rule-based change classification** ("Compare Summarization") so the user can instantly separate formatting noise from substantive edits, and **collapse a renumbering cascade** (hundreds of "Clause 4.2 → Clause 4.3" shifts) into a single annotation.
3. **Moved-paragraph detection** so a clause relocated wholesale is shown as *moved*, not as a delete-here-plus-insert-there pair that hides whether the text changed.

The defining constraint, set by the legal-malpractice risk model: the classifier **categorizes but never suppresses**. A change can be *labelled* "formatting-only" and *hidden behind a filter the user toggles*, but no change is ever silently dropped. (See §2 CUT, §6.)

---

## 2. Scope

### 2.1 In scope ✅ v1

| # | Capability |
|---|---|
| 1 | **Word-vs-Word legal redline** — two `.docx` versions → new document with `<w:ins>`/`<w:del>` tracked-change markup carrying `w:author`/`w:date`. Renders identically to Word's native tracked changes; accept/reject works in standard Word. |
| 2 | **Rule-based change classification** (Compare Summarization) — every change bucketed into the §6 taxonomy so the user can filter formatting-only noise vs substantive changes. **No ML.** |
| 3 | **Hash-match moved-paragraph detection** — a paragraph word-for-word identical in both documents but at a different position is marked **moved**. Covers ~70% of real move cases; ~1-day implementation. |

### 2.2 Deferred 🔜 v2 (not in this spec)

- **Word-vs-PDF comparison** 🔜 — the counterparty's *execution version* very often arrives as a PDF (signed/scanned). Implemented client-side via **PDF.js** in the task-pane WebView — still zero-setup and zero-upload, just deferred. **Not** an edge case: in SG/UK practice, execution copies as PDF are routine (§7).
- **Fuzzy "moved-and-edited" paragraph detection** 🔜 — a paragraph both relocated *and* lightly edited. v1's hash match requires verbatim identity and misses these (shows them as delete+insert). v2 adds similarity-threshold matching (token Jaccard / normalised Levenshtein) before declaring a move.

See [README → Deferred to v2](../README.md#deferred-to-v2-).

### 2.3 Cut ❌ (will not be built)

- **ML semantic-equivalence detection** ❌ (e.g. "these two clauses *mean* the same, hide the change"). **Categorically unacceptable.** A false positive that hides a substantive change is a malpractice risk. Rule-based classification **CATEGORIZES** a change so the user may filter it; it **never SUPPRESSES**. This is a hard product invariant, not a resourcing decision, and applies to v2 and beyond. (See [README → What was cut and why](../README.md#what-was-cut-and-why-).)

---

## 3. Dependencies

- [The OOXML Service](../ARCHITECTURE.md#the-ooxml-service) — Compare is one of its two primary consumers (with Repair): it adds **no new packaging code**, only a *consumer* that reads two parsed packages and emits one. The slice loop, unzip, base64, and `insertFileFromBase64` / `createDocument().open()` fallback all live there. Compare's specific consumer pattern: **read both packages → diff → author a third package** whose body carries `<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>` revision markup.
- [Findings Model](../ARCHITECTURE.md#findings-model) — the `ChangeSummary` buckets feed the shared Findings/summary panel.
- [Technology Stack](../ARCHITECTURE.md#technology-stack) — the diff/zip/xml libraries (diff-match-patch, fast-diff, PizZip, JSZip, fast-xml-parser, PDF.js for v2) are specified there; §3.1 keeps only the per-library **roles/licences** relevant to Compare.
- [Performance Guidelines](../ARCHITECTURE.md#performance-guidelines) and [Cross-Platform Matrix](../ARCHITECTURE.md#cross-platform-matrix) — large-document chunking, two-docs-in-memory limit, web workers.
- [Localisation / Config Layer](../ARCHITECTURE.md#localisation--config-layer) — the classifier reads SG/UK rules from here; values in [en-GB Spelling](../CONVENTIONS.md#en-gb-spelling) and [Terminology](../CONVENTIONS.md#terminology). `compare.*` config keys (§7.5) are feature-local.
- Office JS facts used and **not** restated — see [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference): native Compare not in Office JS (row 17), cross-document reading impossible / second doc via File-API upload (rows 10, 12), `insertFileFromBase64` fidelity artifacts (row 11), batch `ctx.sync()` (row 13).

### 3.1 Office JS surface + libraries (Compare-specific)

| API | Use |
|-----|-----|
| `getFileAsync(Office.FileType.Compressed, …)` | Read the **active** document as `.docx` bytes (canonical OOXML Service entry; sliced via `File.getSliceAsync`, reassembled to a `Uint8Array`). |
| `body.getOoxml()` / `range.getOoxml()` | Content-level OOXML read where a full-package read is unnecessary (previewing a single paragraph). |
| `body.insertOoxml(xml, location)` | Content-level OOXML write — the *summary annotations* injected into output, and the lightweight path when the output is the active document. |
| `Word.run` / `context.sync()` | Orchestration, range navigation, "go to next change". |
| `insertFileFromBase64` | Primary mechanism to materialise the assembled redline `.docx`. **Fidelity caveat — §9.13.** |
| `Office.context.ui` / dialog API | Optional: open the generated redline in a new document window. |

**Libraries (roles; full table in [Technology Stack](../ARCHITECTURE.md#technology-stack)):** `diff-match-patch` (Apache-2.0) — Level-2 word-mode diff; `fast-diff` (Apache-2.0) — short-field char/word diff (cells, headings); `pizzip` (MIT) — sync (un)zip; `jszip` (MIT/GPLv3) — async streaming (un)zip for large packages (§9.10); `fast-xml-parser` (MIT) — parse/serialize `document.xml`; `PDF.js` (Apache-2.0) — **v2 only**, client-side PDF text extraction. **Reference architecture (not code): LibreOffice Writer document comparison (MPL-2.0)** — see §11.

### 3.2 Loading the second document

No Office JS API reads a *second arbitrary file* from disk into an add-in ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) rows 10, 12). Two supported sources:

1. **User upload (primary).** A hidden `<input type="file" accept=".docx">`. On `change`, read via the File API (`file.arrayBuffer()`), feed the `ArrayBuffer` into the OOXML Service unzip step. No server, no temp file.
2. **Active document as one side.** `getFileAsync` supplies the *active* document. Typical flow: active doc = "their markup / latest", uploaded file = "our last sent version" (or vice-versa).

The user explicitly designates **Original (base)** and **Revised (compared)**; redline direction (`<w:ins>` = present in Revised but not Original; `<w:del>` = present in Original but not Revised) follows from that choice.

---

## 4. Data model

All types live in `compare/model.ts` and feed the shared [Findings Model](../ARCHITECTURE.md#findings-model) for the summary panel.

### 4.1 Paragraph hash record

Produced once per side during Level-1 pre-processing.

```ts
interface ParaRecord {
  index: number;            // position in its own document (0-based, body order)
  partId: string;           // 'document' | 'header1' | 'footnotes' ... (a cell's paragraphs keep their part's id)
  rawXml: string;           // original <w:p> OOXML (revision markup quarantined out)
  normText: string;         // normalised text (§5.2) of what the reader sees (§5.0, ADR-0045:
                            //   runs, equation text, field results, text-box content once):
                            //   case kept; whitespace collapsed,
                            //   numbering prefix stripped, smart quotes/dashes normalised, NBSP→space
  normTextAggressive: string; // also lowercased, all punctuation stripped — used ONLY to
                            //   detect the whitespacePunct class (§6 priority 7), never to hash
  numberingPrefix: string;  // stripped auto/manual numbering ("4.2(a)", "(iii)") to restore
  hash: string;             // 64-bit hash (e.g. xxhash / FNV-1a) of normText
  styleId: string | null;   // w:pStyle val — for style-change detection
  // ADR-0042 — the numbering a paragraph does not state. In a firm template the
  // <w:numPr> lives on the STYLE, so `numberingPrefix` (scraped from run text) is
  // '' for every automatically numbered clause. Resolved through
  // pStyle → basedOn → numPr and counted in document order. `numId` and `ilvl`
  // resolve INDEPENDENTLY — each from its own nearest declaration (ECMA-376
  // §17.7.2 merges the style hierarchy leaf by leaf), so a style that overrides
  // only the level re-levels the instance it inherits.
  // NOT part of `hash`, and never to be folded into it (§5.2).
  numbering?: ParaNumbering | null;   // null = unnumbered; absent on the Text/PDF paths
  isTableCell: boolean;
  isSignatureBlock: boolean;// heuristics §7.2 — forces "substantive"
  isHeaderFooter: boolean;
  isFootnote: boolean;
  // A whole table as ONE alignment unit (ADR-0027): the cell-text grid every
  // shared stage reads to keep its hands off a table (no word diff, no move,
  // no merge/split, positional gap stretch, table classifier branch). Set by the
  // Text-vs-text Markdown path AND, since ADR-0046, on every Word table unit
  // (alongside `wordTable`). A record without it is treated exactly as before.
  table?: RowModel[];
  // WORD-VS-WORD BODY TABLES ONLY (ADR-0046): the table's structure. `rawXml` is
  // then the whole <w:tbl>, `hash` is structural (row, cell and paragraph
  // boundaries plus each cell paragraph's hash — a merge holding the same words
  // is a different hash), and `normText` is the cells joined " | " and the rows
  // "  ¶  ". Built by a post-pass AFTER numbering labels are assigned, so every
  // counter sees the cell paragraphs exactly as before. Body part only: header,
  // footer and footnote tables stay flat records.
  wordTable?: WordTableModel;
}

// ADR-0046. XML is kept verbatim, so an unchanged table is re-emitted byte for
// byte and a changed one is re-assembled from the REVISED properties.
interface WordTableModel {
  tblPrXml: string;          // '' when absent
  tblGridXml: string;
  rows: WordTableRow[];
  // false when the table holds a wrapper the model does not represent (a
  // content control or custom XML round a row, a cell or a cell's paragraphs).
  // The wrapper is read THROUGH — what is inside it is in the model and the
  // hash, so an edit there is reported — but a changed table is authored whole
  // (§5.5), wrapper and all.
  supported: boolean;
}
interface WordTableRow {
  tblPrExXml: string;
  trPrXml: string;
  cells: WordTableCell[];
  hash: string;              // over each cell's gridSpan, vMerge and hash
}
interface WordTableCell {
  tcPrXml: string;
  gridSpan: number;          // 1 = not merged across
  vMerge: 'restart' | 'continue' | null;
  content: ParaRecord[];     // paragraphs, and nested tables grouped the same way
  hash: string;              // over the content records' hashes, paragraph vs table marked
}

// ADR-0042. Built in two steps: numId/ilvl/key from the style chain at
// extraction, label/path from the shared counter machine once the whole part
// is in document order (a label is a property of the SEQUENCE, not the paragraph).
interface ParaNumbering {
  // DOCUMENT-LOCAL. **Never compared across two documents.** Word reassigns
  // `w:num` instance ids freely — a Save-As onto another template renumbers
  // every one while each clause keeps its number — and comparing it directly
  // reported six changes on a document with none. Kept only so the counter
  // machine can group counters WITHIN a document.
  numId: string;
  ilvl: number;        // 0-based level. Comparable: level 2 is level 2 anywhere.
  // THE COMPARABLE IDENTITY: the level, plus the shape of the level definitions
  // this paragraph's number is RENDERED FROM — a statement about the two
  // documents, not about Word's bookkeeping. Compare this, never numId. That is:
  //   - the own level's lvlText (the template);
  //   - for each level the template SUBSTITUTES (%1…%9), that level's effective
  //     format (its numFmt, or decimal when the own level's isLgl overrides every
  //     token), its start (startOverride already folded in), and its lvlRestart
  //     reduced to its effective behaviour;
  //   - and, for the OWN level WHEN THE TEMPLATE SUBSTITUTES IT, its DECLARED
  //     numFmt as well as that effective one (an own format the lvlText never
  //     substitutes moves no label in this build or in Word, so it is not in the
  //     key). The two differ only under isLgl, and they must both be carried:
  //     renderListNumbers folds isLgl over the own token and Word does not
  //     (measured — decimalZero + isLgl over an upperRoman ancestor renders
  //     1.01, the ancestor converted and the own format kept), so a key that
  //     carried only the folded value went blind exactly where the renderer is
  //     blind and an edit to the own level's numFmt produced no op at all.
  //   - and, for every level 0…ilvl, whether it has a definition at all — the
  //     counter machine skips a paragraph whose level does not resolve, so
  //     deleting a level definition moves the counters of the paragraphs below it.
  // Deliberately NOT an ancestor's lvlText or isLgl (never read for a descendant),
  // nor any field of a level the template does not substitute.
  // The key covers what the renderer reads and writes WHEN RENDERING THIS
  // PARAGRAPH. It stops at ilvl: a definition DEEPER than this level is read by
  // other paragraphs' render calls, and a per-paragraph key cannot hold a
  // sequence's state. Deleting one moves this paragraph's label — reported, since
  // `numberLabel` is compared from the rendered labels — and classes it as a
  // bystander of a sibling that did not move. Recorded limit: ADR-0042 §
  // Accepted limits, pinned by a test.
  key: string;
  label: string | null;    // rendered label ("(vi)", "Clause 1.01")
  // The counters behind the label, levels 0…ilvl — [2,1] for "2.1". §6.1 needs a
  // SHIFT and a rendered label cannot be subtracted. The whole path, not this
  // paragraph's own counter: under the dominant "%1.%2" style a sub-clause
  // renumbers when its PARENT does and its own counter never moves, so an
  // own-level reading measures every such cascade as a shift of zero.
  path: readonly number[] | null;
  // The levels 0…ilvl whose counters this label actually SHOWS: substituted by
  // its lvlText (%1…%9) AND rendered non-empty by their effective numFmt. [0,1]
  // for "%1.%2"; [1] for "(%2)"; [1] again for "%1.%2" whose level 0 is
  // numFmt="none". §6.1's cascade header names a level and a shift and prints
  // example labels under it, so the shift is measured over THESE levels — read
  // over every level of `path` instead, splitting a clause list under "(%2)"
  // headed three rows that each moved -2 with "shifted (+1)", the parent counter
  // that "(%2)" renders nowhere. Same rule as `key`: what this paragraph's
  // number is rendered from.
  renders: readonly number[] | null;
}
```

### 4.2 Diff result model

```ts
type OpType =
  | 'equal'      // identical paragraph (no output markup)
  | 'insert'     // present only in Revised
  | 'delete'     // present only in Original
  | 'replace'    // matched paragraph pair with intra-paragraph word changes
  | 'move'       // verbatim-identical paragraph at a different index (Level 3)
  | 'property';  // ADR-0042 — matched pair, IDENTICAL text, a non-text property differs
                 //   (styleId, or the effective numbering). Authored exactly as
                 //   'equal': the redline XML is byte-identical, only `ops` grows.
                 //   EXCEPT `objects` (ADR-0045 § Amendment): a footnote/endnote
                 //   reference, picture, text box or embedded object added, removed
                 //   or moved — a change to the content, and marked (§5.5).

interface DiffOp {
  id: string;                      // stable id for UI navigation
  type: OpType;
  origRange: { start: number; end: number } | null;  // ParaRecord indices in Original
  revRange:  { start: number; end: number } | null;  // ParaRecord indices in Revised
  partId: string;
  wordDiffs?: WordDiff[];          // present when type === 'replace'
  propertyChange?: PropertyChange;  // present when type === 'property' (ADR-0042)
  classification: ChangeClass;     // §6 — the dominant/primary category
  classes: ChangeClass[];          // ALL applicable categories (a change can be several)
  author: string;                  // injected w:author (Config; default "Mutatis Compare")
  date: string;                    // ISO-8601 → w:date
  collapsedInto?: string;          // id of a renumbering-cascade annotation, if collapsed
  moveTargetId?: string;           // for 'move': paired DiffOp id
  spellingVariant?: { usSpelling: string; gbSpelling: string };  // §7.1 — badged, never filtered
  fuzzyAnchored?: boolean;         // PDF path (§12): matched by fuzzy anchor — badged, never hidden
  lowConfidence?: boolean;         // PDF path: approximate table-region cell — badged, never hidden
  // ADR-0046 — present on an op pairing two Word table units ('replace', or a
  // 'property' op for a text-equal table whose cells changed invisibly). One op
  // per table: its changes live in the cells' own ops.
  tableDiff?: TableDiff;
}

// ADR-0046. Rows, then cells within each paired row, each cell's content
// analysed by the body's own stages (§5, stage 1b) — so `ops` is an ordinary op
// list over that cell's `content`, nested tables included.
interface TableDiff {
  rows: TableRowSlot[];          // document order; a deletion before the insertion beside it
  columns?: TableColumnSlot[];   // set when a column was added or removed (§9.3)
}
type TableRowSlot =
  | { kind: 'match'; o: number; r: number; cells: TableCellSlot[] }
  | { kind: 'insert'; r: number }
  | { kind: 'delete'; o: number };
type TableCellSlot =
  | { kind: 'match'; o: number; r: number; ops: DiffOp[] }   // [] = content identical
  | { kind: 'insert'; r: number }                            // a cell of an added column
  | { kind: 'delete'; o: number };
type TableColumnSlot =
  | { kind: 'match'; o: number; r: number }
  | { kind: 'insert'; r: number }
  | { kind: 'delete'; o: number };

interface WordDiff {
  op: 'eq' | 'ins' | 'del';
  text: string;
  atomic?: 'definedTerm' | 'clauseRef' | 'amount' | 'date' | 'percent' | 'duration';
}

// ADR-0042 — which non-text properties differ across a text-identical pair. It
// is what the taxonomy reads to tell a clause that was RE-LEVELLED (formatting,
// §6 row 8 / §9.9) from one whose number merely MOVED because a sibling did
// (crossRefRenumber, §9.4 — collapsible under §6.1).
interface PropertyChange {
  styleId: boolean;             // w:pStyle differs
  numberingIdentity: boolean;   // ParaNumbering.key differs — NOT (numId, ilvl)
  numberLabel: boolean;         // the rendered label differs (may be true alone — a cascade)
  // The VALUES that moved. A property op has no `wordDiffs`, and the row preview,
  // the navigation anchor and the exported report all read `wordDiffs` — without
  // these the change is counted and then shown as a blank row that does nothing
  // when clicked. Facts, not a sentence: each surface formats them itself.
  from: PropertySide;
  to:   PropertySide;
  // ADR-0045 § Amendment — the ordered zero-text objects differ (absent = false):
  // a footnote or endnote reference, picture, text box or embedded object was
  // added, removed or MOVED, and not a word changed. Classified substantiveText
  // (§6), never a cascade, and the one property the redline marks (§5.5).
  objects?: boolean;
  // Set only for a pure cascade — never when `objects` is set. ... §6.1 groups by a uniform shift and "(d)" → "(e)"
  // cannot be subtracted, so the counter paths supply it: the shift is the
  // shallowest level at which the two paths disagree, AMONG THE LEVELS THE LABEL
  // SHOWS (ParaNumbering.renders) — a header naming a level the template never
  // prints contradicts the labels printed beneath it. `shiftedIlvl` is THAT
  // level — the one whose counter moved, not the paragraph's own ("1.1" → "2.1"
  // is a level-1 paragraph carried by a level-0 shift) — and it keeps two
  // levels' cascades from merging.
  cascade?: { from: string; to: string; shiftedIlvl: number; shift: number };
}

interface PropertySide {
  styleId: string | null;
  numberLabel: string | null;
  ilvl: number | null;
  objects?: readonly ZeroTextObject[];  // set when `objects` is: this side's, in order
}

// A zero-text object and its place: `at` = characters of the normalised,
// prefix-stripped text before it, so a whitespace run written differently or a
// manual number that moved cannot move it. In the manual number rather than
// after it, a code — -2 before the number, -1 inside it — so renumbering does
// not move it and moving it across the number does. Nothing says WHICH footnote or image
// (`w:id`, `r:embed` are document-local and Word reassigns them — one footnote
// inserted early renumbers every one after it).
interface ZeroTextObject {
  sort: 'footnote reference' | 'endnote reference' | 'picture' | 'text box' | 'embedded object';
  at: number;
  within?: number;  // inside an equation or text box: the innermost one holding it (a
                    // collapsed space can make "just before" and "just inside" one count)
}
```

### 4.3 Change-summary buckets (Findings model)

```ts
interface ChangeSummary {
  total: number;
  byCategory: Record<ChangeClass, number>;     // counts per §6 category
  renumberingCascades: CascadeAnnotation[];    // each collapses N raw shifts → 1
  movedParagraphs: number;
  spellingVariantCount: number;                // §7.4 — its own line in §8.2
  // §5.0 — comments quarantined out of each side before the diff. Reported,
  // never dropped silently ("Original contained 12 comments — excluded from
  // compare") rather than letting a reviewer assume one survived the redline.
  quarantinedCommentsOriginal: number;
  quarantinedCommentsRevised: number;
  // §9.7–§9.9 — the parts diffed OUTSIDE the body. Headers and footers are zero
  // at the default config, where they are not processed at all (not
  // processed-and-hidden). FOOTNOTES ARE NOT: compare.footnotes defaults to
  // 'compare' (§7.5) because §9.8 makes a footnote content change substantive,
  // so this count can be non-zero with nothing switched on. All three are
  // counted into `total` and `byCategory`; only their NAVIGATION is outstanding
  // (a part op's range indexes its own part's paragraphs, not the body's).
  partChanges: { headers: number; footers: number; footnotes: number };
  styleDefinitionNote: string | null;          // §9.9 — styles.xml itself changed
  filtersActive: {
    hideFormattingOnly: boolean;
    collapseRenumbering: boolean;              // default ON
  };                                           // two, not three — ADR-0036
}

interface CascadeAnnotation {
  id: string;
  // A reference token (§7.4) for a crossReference cascade; for a numbering
  // cascade, the clause numbering and its level ("Clause numbering (level 2)").
  tokenKind: string;
  shift: number;                  // +1 / -1
  affectedCount: number;          // how many raw reference changes it represents
  exampleRefs: string[];          // e.g. ["Clause 4.2 → 4.3", "Clause 4.3 → 4.4"]
  // ADR-0042 — what cascaded: the document POINTING at a clause whose number
  // moved, or the clause NUMBERS themselves. Same cause, same collapse, two
  // different sentences to a reader. Absent ⇒ 'crossReference'.
  source?: 'crossReference' | 'numbering';
}

type ChangeClass =
  | 'formatting'
  | 'whitespacePunct'
  | 'crossRefRenumber'
  | 'obligationLanguage'
  | 'definedTerm'
  | 'numericalValue'
  | 'structural'
  | 'substantiveText';
```

---

## 5. Algorithm — the 3-level diff pipeline

```
┌───────────────────────────────────────────────────────────────────┐
│ 0. PRE-PROCESS each side                                          │
│    flatten existing tracked changes (accept-all | reject-all)     │
│    quarantine remaining revision markup + comments                │
│    extract parts (document, headers/footers, footnotes)           │
│    normalise text + strip numbering prefixes → ParaRecord[]       │
│    body <w:tbl> → ONE unit each, after numbering (ADR-0046)      │
├───────────────────────────────────────────────────────────────────┤
│ 1. STRUCTURAL DIFF  (paragraph hashes → LCS → add/del/match)      │
│    anchor matching first, then LCS on hash sequences             │
│    each gap paired by similarity, positional fallback (ADR-0039) │
├───────────────────────────────────────────────────────────────────┤
│ 1a. NON-TEXT PROPERTIES of each text-equal pair (ADR-0042)        │
│    style id + resolved numbering (identity, rendered label)      │
│    → 'property' op; authored byte-identically to 'equal'         │
│    + zero-text objects added/removed/moved (ADR-0045 § Amend.)   │
│      → 'property' op with `objects`; the objects are marked      │
├───────────────────────────────────────────────────────────────────┤
│ 1b. WORD TABLES (ADR-0046): each pair of table units →           │
│    rows (LCS, gaps by similarity) → cells, each cell's content   │
│    through stages 1–4a recursively → `tableDiff` on the op       │
├───────────────────────────────────────────────────────────────────┤
│ 2. WORD DIFF inside each matched-but-changed pair (DMP word mode) │
│    atomic tokens: defined terms, clause refs, amounts, dates     │
├───────────────────────────────────────────────────────────────────┤
│ 3. MOVE DETECTION (light): hash-match deletes↔inserts → 'move'    │
├───────────────────────────────────────────────────────────────────┤
│ 4. CLASSIFY every DiffOp (rule-based taxonomy §6)                 │
├───────────────────────────────────────────────────────────────────┤
│ 4a. MARK ORDER — the insertion reads before the deletion it       │
│    replaces; seam-bearing windows exempt (ADR-0043)              │
├───────────────────────────────────────────────────────────────────┤
│ 5. AUTHOR OUTPUT OOXML  (<w:ins>/<w:del>, numbering restored)     │
│    + summary annotations → ChangeSummary                         │
└───────────────────────────────────────────────────────────────────┘
```

### 5.0 Pre-processing

**Existing tracked changes.** A document may already contain `<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`. Diffing raw markup-laden XML produces garbage. Before diffing, **flatten**:

- **Accept-all** (default): apply all revisions — `<w:ins>` content kept, `<w:del>` content removed — yielding the "as-if-accepted" text.
- **Reject-all**: inverse — `<w:del>` content kept, `<w:ins>` removed.
- **A revision on a paragraph MARK** (`<w:pPr><w:rPr><w:del/>`, the direct `rPr`, never one inside `w:pPrChange`) is a paragraph join (ADR-0045): the policy that removes the mark — accept-all a deleted or moved-away mark, reject-all an inserted or moved-here one — **joins** that paragraph to the next, its runs put before the next paragraph's and the next paragraph's properties and attributes kept, because its mark is the one that survives (the rule applied here; how Word desktop picks the joined paragraph's properties is **unmeasured** until the live-host smoke test — it decides the paragraph style, and so the numbering label, only where an input document already carries a tracked join); the other policy keeps both and drops the mark. A chain of removed marks joins into one paragraph. Only a `w:p` that is the next block in the same container is joined (bookmarks and comment ranges between are skipped and stay where they are): the last paragraph of a body or cell, and one followed by a table, stays. Before this, a document that already carried a tracked paragraph join was compared as two paragraphs, and the redline's own merges (§5.5, ADR-0032) and whole-paragraph deletions (ADR-0035) read through the flattener as one paragraph too many.

The user chooses per side (UI toggle, §8). Flattening is done on the parsed tree, not via Word, so it works on the *uploaded* side too where we have no live document. After flattening, **quarantine** any residual revision elements and `<w:commentRangeStart/End>` / `w:comment` references so they cannot leak into output and so comments don't perturb text extraction. Quarantined comments are *reported* ("Original contained 12 comments — excluded from compare") but never silently merged.

**Paragraph text.** A paragraph's compared text is what the reader sees in it, read by one extractor (`preprocess/inline-units.ts → collectText`, ADR-0045): run text (`<w:t>`), **equation text (OMML `<m:t>`)**, field *results* (never `<w:instrText>`), and text-box content — the `mc:Choice` copy only, never the `mc:Fallback` duplicate. An equation, a text box and each paragraph inside a box read as their own words; one that holds no text reads as nothing. Until ADR-0045 `<m:t>` was not read, so a change to an equation and nothing else reported no change at all. An object with no text of its own — a footnote reference, an inline image — contributes nothing, so a paragraph whose *only* change is losing one is still reported equal (ADR-0045 § Accepted limits 1).

**Tables (ADR-0046).** In the body part, each top-level `<w:tbl>` becomes **one** record (`ParaRecord.wordTable`, §4.1) carrying the whole table's XML and its rows and cells, each cell's paragraphs and nested tables as records of their own. The grouping runs **after** numbering labels are assigned over the flat paragraph sequence, so every counter and label is what it was; it walks the XML exactly as the paragraph extractor does and, if the two ever disagree about the paragraph count, returns the flat records untouched (nothing is renumbered until the whole walk is known to be consistent). A content control or custom XML round a row or cell is read through (§4.1 `supported`); a table holding a paragraph outside every cell, or a cell outside every row, — schema-invalid, but a generator can write it — is left as flat records alone, every change to that paragraph reported, its structure not kept. Flattening resolves tracked **rows and cells** the way Word does: accept-all drops a row whose `trPr` carries `<w:del>` and a cell carrying `<w:cellDel>`, reject-all drops the inserted ones, the cell marks are always stripped, and a table every row of which the policy removed goes whole — no empty `<w:tbl>` shell is left where a tracked table insertion was rejected.

**Numbering normalisation.** Auto-numbering (`<w:numPr>`) renders a visible prefix ("4.2(a)") *not* in the run text, while some documents hard-code the prefix into run text. Normalise both: detect leading numbering tokens, strip them into `ParaRecord.numberingPrefix`, hash/diff the **prefix-free** text, and **restore the prefix in output**. This prevents a single inserted clause (which shifts every subsequent number) from making every following paragraph look "changed" at Level 1 — the cascade is instead caught and collapsed by the classifier (§6).

### 5.2 Text normalisation (for hashing & whitespace classification)

Applied to produce `normText`:
- Collapse runs of whitespace to a single space; trim; NBSP → space.
- Normalise curly quotes (`" " ' '`) and dashes (`– —`) to canonical forms.
- Strip the numbering prefix (above).
- **Case is preserved** (defined terms and en-GB/US spelling are case- and letter-sensitive).
- A *second*, more aggressive normalisation (also lowercase, strip all punctuation) is computed **only** to test the `whitespacePunct` classification — never used for hashing the redline.

### 5.1 Level 1 — structural / paragraph diff

1. Build `hash` for every `ParaRecord` on both sides.
2. **Anchor matching before LCS** (§9.1): find paragraphs whose hash is **unique on both sides** and identical across sides — unambiguous anchors. Pin them, then run LCS only *between* consecutive anchor pairs. Keeps reordering and large inserts from producing pathological alignments and bounds LCS cost on big documents.
3. **LCS on the hash sequence** within each anchor-bounded segment → longest common subsequence of matching paragraphs. Matching pairs with equal hash → `equal`. What lies between two matches is a **gap**: originals and reviseds with no hash match.
4. **Gap pairing (ADR-0039).** Each gap's leftovers are paired into `replace` candidates; anything unpaired becomes `insert` (Revised-only) or `delete` (Original-only). The pairing has three rules, in order:
   1. **Lock** — a *positional* pair (the k-th original with the k-th revised) that renders inline today keeps its partner. "Inline today" is decided by running the gap's positional ops through the merge/split re-pairing (below) and the §5.3 demotion, never re-derived.
   2. **Align** — in each stretch between locked pairs, choose the **order-preserving** set of pairs whose Dice is ≥ `replaceSimilarityThreshold`, maximising total Dice. Pairs never cross, and a gap is bounded by matched paragraphs, so nothing is paired across unchanged text (that would be moved-and-edited detection, v2 — SCOPE). A stretch holding a table unit stays positional; a paragraph with no content tokens is never a candidate; a paragraph whose text appears verbatim on the other side is **held** for §5.4's move detection: never chosen by the alignment, and never paired by rule 3 beside a re-aligned pair (where positional pairing already paired it — a locked pair, or a stretch the alignment leaves alone — it stays as it was), and never absorbed by the merge/split pass below unless positional pairing's own merge pass absorbed it (`GapPairing.mayAbsorb`).
   3. **Fall back** — what remains is paired by position (held records excepted). A stretch the alignment does not touch is exactly what positional pairing made it.

   Candidacy shares the demotion's Dice arithmetic (`bagDice` / `diceFromOverlap`), so a pair chosen here is never demoted at §5.3. The alignment is bounded per stretch by a work budget (`MAX_ALIGN_WORK`: cells × (a fixed per-cell cost + mean distinct tokens), fitted to measured time) and a cell cap (`MAX_ALIGN_CELLS`, 10⁶), and the whole comparison by a total (`MAX_ALIGN_WORK_TOTAL`, three stretch budgets); above any of them the stretch keeps positional pairing. Rule 3 walks both sides in document order, so every op list stays in order on both sides (§5.5 authors by walking it). Applied by `runComparePipeline` (Word-vs-Word, Text-vs-text, batch). The PDF pipeline and the parts diff (headers, footers, footnotes) stay positional. The guarantee, what it does not promise (op order inside a re-aligned stretch, so a merge presented orphan-first is not detected; new adjacency among previously churned records can form a merge or split) and its measurements are in ADR-0039 §1.

Output: an ordered list of `equal` / `insert` / `delete` / `replace` spans. One inserted paragraph costs exactly one `insert` op — it does **not** cascade misalignment onto every following paragraph, **including inside a gap**: before ADR-0039 a paragraph inserted or deleted in a run of changed paragraphs shifted every positional pair after it, and the §5.3 demotion then showed each shifted pair as a whole paragraph struck and re-typed.

**Paragraph merge / split re-pairing (ADR-0032).** Level-1's matching is one-to-one, so a drafter joining two paragraphs leaves the pipeline with a large `replace` beside an orphaned `delete` — the same words struck in one paragraph and inserted into the next, which is the P2 failure at paragraph scale. `diff/merge-split.ts → detectMergesAndSplits` re-pairs those two ops into **one `replace` spanning the whole range** on the many-side (`origRange` covering two paragraphs for a merge, `revRange` for a split); a split is the mirror. No new op type is needed — the ranges already exist and `recordsInRange` already walks them. It runs **before** the ADR-0026 threshold below, so that gate scores the merge against the *joined* original text, which is what it actually replaces.

Guarded on structure alone (ADR-0008: no semantic-equivalence scoring): the absorbed paragraph must be at least 4 content tokens; at least **85%** of it must reappear token-for-token **and in order** at the matching end of the survivor (positional, so a neighbour that merely shares vocabulary is not "absorbed"); what remains of the host must still resemble the paragraph the `replace` matched, scored at `replaceSimilarityThreshold` so the two gates cannot disagree; the orphan must be a single paragraph immediately **after** its pair (`replace`+`delete` for a merge, `replace`+`insert` for a split — ADR-0032's correction 1 removed the mirrored orderings; since ADR-0039 an orphan-first shape is reachable among previously churned records and is deliberately not matched, and a record held for a move is refused as the orphan unless positional pairing's merge pass absorbed it too); and never a table unit, which is one alignment unit (ADR-0027). A host that is *entirely* the absorbed paragraph is a deletion of the other paragraph, not a merge. The pass only changes **which op owns which record** — it reconstructs both sides' full text before and after and returns its input unchanged on any mismatch.

**Non-text property changes (ADR-0042).** `hash` is prefix-stripped visible text, so two paragraphs that differ only in `w:pStyle` or in resolved numbering pair as `equal` and produce no op at all — a change found and then not reported, which ADR-0008 forbids. `diff/property.ts → detectPropertyChanges` walks the one-to-one `equal` pairs and emits a `property` op (§4.2 `PropertyChange`) where the style id, the numbering IDENTITY (`ParaNumbering.key`) or the rendered `numberLabel` differs. The label is compared **only when both sides resolved one**: a label present on one side and absent on the other is not evidence of a renumbering but of one side's level definition failing to resolve, and reporting it as a moved number would invent a change. The real case is still covered, because a paragraph that gained or lost numbering also changed identity. It runs **after** the merge/split re-pairing and the §5.3 threshold demotion — `pipeline.ts` calls it stage 2.75 — so it sees the final set of `equal` pairs and never re-examines one a later pass would have re-owned. **`equal` pairs are also its whole reach, which is a stated limit and not an oversight:** a paragraph whose words were edited *or* which moved pairs as a text op, so a property change riding on it is **not** reported — the reader sees the text change and nothing about the renumbering (ADR-0042 § Accepted limits, Four). Making `replace` and `move` carry a `propertyChange` too is a real design question — one op, two kinds of change, one `classification` — and is deferred to its own decision rather than settled here.

A `property` op is authored by exactly the `equal` path, so `authoredBodyXml` is **byte-identical** with and without the pass; only `ops` and the summary grow — **with one exception, below**. The `w:numPr` a firm template puts on the STYLE rather than the paragraph is resolved through `@shared/numbering` (ECMA-376 §17.3.1.19), which is why these changes were invisible before: on such a document `document.xml` contains no `<w:numPr>` anywhere.

**Zero-text objects (ADR-0045 § Amendment).** A footnote or endnote reference, a picture, a text box or an embedded object contributes no text, so a paragraph whose ONLY change is one of them added, removed or moved paired as `equal` and produced no op (`summary.total` 0) — ADR-0008 suppression. The same pass now also compares each one-to-one pair's ordered list of such objects (`preprocess/inline-units.ts → zeroTextObjects`): a sort and a place, the place counted in the normalised, prefix-stripped text. Not which footnote or image it is — `w:id` and `r:embed` are document-local, and one footnote inserted early renumbers every reference after it. A difference sets `PropertyChange.objects` (§4.2). Unlike every other property it is a change to the content: classified `substantiveText` (§6), never a cascade member (§6.1), and **marked** in the redline (§5.5). A table record is never read here (ADR-0027). Comments, bookmarks and a link or field added around unchanged words are not zero-text objects and are out of this rule (ADR-0045 § Accepted limits 1).

### 5.3 Level 2 — word diff inside changed paragraphs

**Change-complexity threshold (pre-word-diff, ADR-0026).** Before word-diffing, a `replace` pair whose paragraph similarity (Sørensen–Dice over content tokens) is **below `compare.replaceSimilarityThreshold`** (default `0.5`) is **demoted to a whole-paragraph `delete` + `insert`** — the conventional "too dissimilar → show as insert/delete" readability rule of commercial comparison tools, so a wholesale-rewritten clause reads as a clean swap instead of token-by-token "diff salad". This is **never suppression** (ADR-0008): the demoted pair emits the *entire* original paragraph as `<w:del>` and the *entire* revised paragraph as `<w:ins>` — strictly **more** markup, never less. Word-vs-Word only; the PDF path keeps its opposite-direction fuzzy-anchor re-pairing (§12, ADR-0017). `replaceSimilarityThreshold = 0` disables it.

A `replace` op's two sides are joined from their record ranges by `rangeDiffText`, which puts a Private Use Area sentinel (`PARA_SEAM`, U+E000) where each paragraph mark sat. A range of **one** record is its `normText` **unchanged**, so every ordinary one-to-one `replace` produces a byte-identical word diff. The sentinel is neither whitespace nor a word character, so the tokenizer emits it as its own token and the diff strikes or inserts the paragraph mark **as a unit** — a first-class edit, never glued away invisibly (§5.5, ADR-0032).

For each surviving `replace` pair, run **diff-match-patch in word mode** (tokenize on word boundaries, map words → sentinel chars, `diff_main`, `diff_cleanupSemantic`, map back). **Atomic tokens** — tokenized as single indivisible units so they are never split mid-token:

- Defined terms in quotes — `"Permitted Encumbrance"`.
- Clause/Schedule reference **values** — the `4.2(a)` of `Clause 4.2(a)`, the `3` of `Schedule 3`, both the `2` and the `4` of `paragraph 2 of Schedule 4` (ADR-0029; see below).
- Amounts — `S$1,250,000`, `£40,000`, `$2.5m`.
- Dates — `1 January 2026`, `1 Jan 2026`.
- Percentages / durations — `7.5%`, `30 days`, `three (3) Business Days`.

Atomicity matters legally: `£40,000` → `£50,000` must surface as one numerical-value change, not "`4`→`5`". Each `WordDiff` carries its `atomic` kind to feed classification.

**How the worked examples below are written.** Every `[-…-]` / `{+…+}` rendering
in this section is shown in the **authored** order — the insertion first, then the
deletion it replaces (§5.5, ADR-0043). The passes described here are all
length-preserving re-partitionings; none of them decides that order, and the
transposition example at the end of the section was already written this way
before the order was made uniform, which is why it reads as the precedent rather
than as an exception.

**Cross-reference granularity (ADR-0029).** A cross-reference is *matched* as a whole span, **case-insensitively** — in-text references are routinely lower-case (`the limitations in clause 9.13`) and must classify as `crossRefRenumber` and join a cascade just as `Clause 9.13` does — but it is *emitted* as several tokens: the reference noun, the whitespace and the `of` connective as ordinary words, and only each reference **value** tagged `clauseRef`. Atomicity makes the number indivisible; it does not make the noun in front of it borrowable, and marking both sides of a renumbering makes the reader re-read a word that did not change:

```
bad   The limitations in {+Clause 9.14+}[-Clause 9.13-] apply …
good  The limitations in Clause {+9.14+}[-9.13-] apply …
```

Two consequences for callers. The word diff keys its sentinels on **kind + text**, so a reference value never aligns with an identical-looking ordinary word. And a bare value (`9.13`) is not recognisable out of context — the value pattern alone matches almost any word — so anything needing a token's kind reads the `atomic` tag the tokenizer produced (`tokenizeTagged`) rather than re-deriving it; `classify/cascade.ts` recovers the reference *noun* by reading back the unchanged text immediately before the mark, normalised to the profile's own spelling so `Clause` and `clause` group into one cascade. The rule is one-directional and binds every downstream pass: a pass may test `atomicKindOf` to decide **not** to split a token, but may never use it to **give** a token a kind the tokenizer withheld. `Subclause 4.2` is the case that makes this bite — `Subclause` is not one of the profile's `referenceTokens`, so its `4.2` is deliberately untagged, and a pass that re-derived a kind from the bare value would invent a cross-reference the tokenizer had already declined to find.

**Relocation-preserving cleanup (pre-decode, ADR-0034).** The same
sentinel-space mistake bites one step earlier, inside `diff_cleanupSemantic`
itself. Its rule is "eliminate an equality no longer than the edits either side
of it", measured with `String.length` — so on a **slot relocation**, where one
span is struck where it sat and re-inserted a few words later with the text
between untouched, the one-token equality in the middle loses to the two-token
edits flanking it. `diff_main` had rendered the move correctly; the cleanup fuses
it into a block that strikes and re-inserts the untouched middle for nothing (P1,
P2), and nothing downstream can undo that — `refineReplaceBlocks` refuses to
strike a word in one mark and re-insert it in another, and `detectTransposition`
needs a shape the fusion has already erased. `diff/relocate.ts →
cleanupPreservingRelocations` therefore applies `diff_cleanupSemantic` to every
stretch of the diff **except** across a relocation window, so every other
over-fusion is still cleaned up exactly as before:

- the window is a `del` run, **one** `eq` run and an `ins` run (either order) — nothing else changed between the two ends;
- the two runs must be **verbatim identical**: the same tokens in the same order, ignoring only pure-whitespace tokens, compared on the **sentinels** so the comparison carries each token's atomic kind and a reference value can never match an identical-looking ordinary word (ADR-0029). Exact sequence equality, never a similarity score;
- the relocated span must carry **readable content** — a letter or a digit.

Both sides are reconstructed and compared before and after; on any mismatch the
ordinary whole-diff cleanup is returned instead (ADR-0008). The relocation stays
an ordinary insert/delete pair — as §5.4 records, `w:moveFrom`/`w:moveTo` is
reserved for paragraph-scale verbatim moves.

**Replace-block refinement (post-decode, before anchoring — ADR-0030).** `diff_cleanupSemantic` measures an equality with `String.length`, and the word diff runs it in **sentinel** space where one character is one **token**. Whitespace is a token of its own, so a lone space out-weighs real shared prose: the equality is destroyed, the neighbouring edits grow by what they swallowed, the next equality becomes eligible in turn, and three small deletions avalanche into one whole-paragraph swap. `diff/refine.ts → refineReplaceBlocks` re-diffs each struck-beside-inserted block on **real tokens** with no cleanup, and keeps the finer rendering only when it earns its extra marks:

```
bad   … indemnify {+the Purchaser against all Losses.+}[-and hold harmless the Purchaser against any and all Losses whatsoever.-]
good  … indemnify [-and hold harmless -]the Purchaser against [-any and -]all Losses[- whatsoever-].
```

Two acceptance rules. It must recover a shared **word** — recovering only whitespace turns one rewritten phrase into two marks for nothing (P3), so `{+from each+}[-out of any-]` stands. And it must not strike a word in one mark and re-insert it in another (P2), so a transposition stays coarse and is left to move detection. When the re-diff is rejected, one fallback is tried: pull out the run the deletion **ends** on and the insertion **begins** with, so that shared run belongs to neither mark — DMP's `diff_commonOverlap` move, which its character-count ratio never reaches in token space. Punctuation tails are split off locally at the block only, never for an atomic value, exactly as in anchoring below. The pass is **length-preserving** and self-checks: it re-partitions the same two texts and returns its input unchanged if reconstruction ever differs (ADR-0008).

**Edit-boundary anchoring (post-decode, ADR-0028).** `diff_main` returns *a* minimal edit script, not the most readable one, and the word diff runs in **sentinel** space — every token mapped to an arbitrary code point — so DMP's own boundary-shifting pass (`diff_cleanupSemanticLossless`) is scoring characters that carry no punctuation or whitespace information. After decoding back to real tokens, `diff/anchor.ts → slideEditBoundaries` slides each `eq → (all-ins | all-del) → eq` edit right through identical neighbouring tokens and keeps the best-reading placement, scored on structural signals only (no vocabulary lists): a span that **opens on the same word that reappears immediately after it closes** is penalised heavily; opening on (or straight after) punctuation, and closing on punctuation, are rewarded; whole-word edges break ties, and ties prefer the *later* placement. This fixes the appended-proviso case, where the new material ends on the word the original sentence already ended on:

```
bad   … ending on or before {+Completion, provided that … paid before +}Completion.
good  … ending on or before Completion{+, provided that … paid before Completion+}.
```

To let a boundary reach a word carrying a punctuation tail, the pass splits that tail into its own token **locally, at the boundary only** (`Completion.` → `Completion` + `.`), and never for a token `atomicKindOf` recognises — amounts, dates, periods, clause references and quoted defined terms stay indivisible. The global tokenizer is unchanged.

The pass is **length-preserving**: a shift is legal only when the text leaving the front of the span is the same text arriving at its back, so it moves *where a mark sits*, never *what is marked* — no change can leave the redline (ADR-0008). It verifies that on every call, reconstructing both sides before and after and returning its input unchanged if they ever differ. A two-sided `replace` has no free boundary and is left alone.

**Adjacent transposition (last, post-coalesce — ADR-0031).** Reordering two words is one editing act, and Myers/LCS diffing renders it as two: it keeps the separator between them as an equality and reports two independent substitutions, leaving the reader to notice unaided that the struck words are the inserted words in the other order. `diff/transpose.ts → detectTransposition` recognises the `del(A) ins(B) eq(mid) del(B) ins(A)` shape and re-expresses it as one relocation:

```
bad   The Vendor shall {+defend+}[-indemnify-], {+indemnify+}[-defend-] and hold harmless …
good  The Vendor shall {+defend, +}indemnify[-, defend-] and hold harmless …
```

One span moves across the connective; the other is left unmarked. It stays an ordinary insert/delete pair — P7 "at phrase scale" does not need `w:moveFrom`/`w:moveTo` inside one sentence. Guarded on **structure alone** (ADR-0008: no semantic-equivalence scoring):

- both spans ≤4 non-whitespace tokens, verbatim-identical at each occurrence and different from each other;
- the connective carries **no word** — whitespace and punctuation only. The rewrite duplicates it (inserted with B, struck with B's old copy), and a word struck in one mark and re-inserted in another is the same P2 failure the refinement above refuses to commit; the two rules have to agree. `A and the B` → `B and the A` therefore stays two clean one-span replacements;
- the connective is ≤8 **characters** — it carries no word, so its cost to the reader is the distance it puts between the two spans, not a token count;
- the five-entry window must cover **whole** del/ins runs (the entry before it is not a `del`, the entry after is not an `ins`), so the pass cannot fire on a slice through a longer span and report a partial move.

Each span keeps its atomic kind, so a swap of two amounts still classifies as a numerical-value change; where a span is atomic `coalesce` leaves it as its own entry, and the renderer draws adjacent same-op entries as one continuous mark. Length-preserving and self-verifying like the passes above. This is the phrase-scale half of P7; §5.4 is the paragraph-scale half.

### 5.4 Level 3 — verbatim moved-paragraph detection

After Level 1, take the set of `delete` paragraph hashes and the set of `insert` paragraph hashes. Any hash present in **both** is a verbatim move: removed from position X (Original) and appearing at position Y (Revised) unchanged. Pair them, retype both ops as `move` (`moveTargetId` cross-links the pair), and emit Word native `<w:moveFrom>` / `<w:moveTo>` markup (shared `w:name`) so Word shows it as a move, not delete+insert.

Limitations: requires **verbatim** identity (hash equality). A relocated-and-edited paragraph fails the hash test and is reported as delete+insert — the v2 fuzzy-move gap (§2, §12). If a hash appears in both sets more than once (duplicated boilerplate), match greedily by nearest index and leave residue as plain insert/delete (do **not** guess).

### 5.5 Output OOXML construction

Author a fresh `document.xml` body by walking the ordered op list:

- `equal` → emit the paragraph unchanged (numbering prefix restored). So does a `property` op (ADR-0042) — except one with `objects` set (§5.1, ADR-0045 § Amendment), which is authored in three tries, most faithful first: (1) the **revised paragraph as it is**, every run with its own formatting, and only the objects that differ marked — a revised one with no counterpart inserted where it sits, an original one struck at its own place (a run of text split there if need be). The objects are aligned on sort and place (longest common subsequence), so one footnote removed of three marks that one alone, and the result passes the replace author's check plus a check that each view's zero-text objects sit where they sit in its document; (2) failing that — an object that differs inside a link or text box, an object holding a space — the object-aware replace below over an all-unchanged word diff, which marks the container whole and rebuilds the words — from their source runs where that reconstructs (below), otherwise from text, when their run formatting goes (ADR-0045 § Accepted limits 5, as narrowed later) — held to the same places check, since its census cannot see an object that moved inside a container; (3) failing that, **one** paragraph — the revised paragraph's properties, its runs inserted and the original's struck — never a deleted paragraph plus an inserted one, which authored a section break twice (a mark carrying `w:sectPr` is never revised, above). A self-closed `<w:p/>` is opened before any of the three. Never the revised paragraph verbatim (the redline silent) and never the text rebuild (the objects gone).
- `insert` → wrap inserted runs in `<w:ins w:id="…" w:author="…" w:date="…">…</w:ins>`, **and mark the paragraph mark itself** `<w:pPr><w:rPr><w:ins …/></w:rPr></w:pPr>` (ADR-0035).
- `delete` → wrap deleted runs in `<w:del …>` with `<w:delText>` replacing `<w:t>`, **and strike the paragraph mark itself** `<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>` (ADR-0035).

  A whole-paragraph change must revise the paragraph **mark** as well as the runs inside it, or the paragraph outlives the change: accepting a deletion whose mark is unmarked leaves an **empty paragraph** where the clause was, and rejecting an insertion whose mark is unmarked leaves one where the added clause was. Word removes the paragraph only when its mark carries the revision. The same two shapes are exempt in **both** directions, because revising their mark destroys something the change was never asked to touch — one rule for both sides, so accept-all and reject-all cannot disagree about which paragraphs survive:
  - **a mark carrying `w:sectPr`** — the last paragraph of a section holds that section's page setup, so removing the mark merges this section into the next and silently re-paginates the document;
  - **the last paragraph of its part** — a body must end with a paragraph and a table cell must contain at least one, so there is nothing to merge into and removing it yields a package Word may refuse.

  For those two the empty paragraph stands: a smaller defect than a re-paginated document or an invalid file. Marks are placed by the same `CT_PPr` rules the merge/split path below uses (one `<w:rPr>`, ordered `rPr → sectPr → pPrChange`, revision first inside `CT_ParaRPr`, leading `w:pPr` matched by depth-counting).
- `replace` → emit the paragraph with run-level `<w:ins>`/`<w:del>` from the Level-2 `WordDiff` list interleaved (unchanged runs plain, inserted in `<w:ins>`, deleted in `<w:del>`), **in list order — the author never re-normalises a changed region**. Within a changed region that order is **insertion first, then the deletion it replaces** — `{+non-exclusive+}[-exclusive-]`, not `[-exclusive-]{+non-exclusive+}` (ADR-0043). The text that SURVIVES is what the reader is being asked to approve, so it reads first, and it is what `detectTransposition` (§5.3) already emitted before the order was made uniform. **This is a deliberate choice, and differs from the conventional redline style, which shows the struck text first.** Whole-paragraph rewrites, which §6 demotes to a separate delete and insert, still read struck-first. The order is set once, on the `WordDiff` stream after classification (stage 4a of §5's diagram), so the redline, the paginated preview, the side-by-side panel and every export cannot disagree about it. **Two** consumers are order-**sensitive** and are named rather than left to be found (ADR-0043 § Consequences). §8.1's *Go to change* anchor takes the first changed line that is *distinctive* enough to search for, so it is now usually drawn from the revised text rather than the original — not always, because a short insertion beside a long deletion still leaves the anchor on the deletion. It was never drawn from both, and the redline itself carries both sides, so navigating the default output is unaffected. §8.3's change-list row preview walks both sides in stream order and truncates from the head, so it now leads with the insertion, and for a replace whose insertion alone exceeds the 80-character budget the struck text no longer appears in the row — the old order truncated the insertion instead. The row, its category and its drill-down are unchanged, so nothing is suppressed; which side a long replace's preview favours is a UI question this spec does not settle. The pass swaps **whole `del`/`ins` runs** and never re-orders within a side, so each side's text is preserved by construction; like the four §5.3 passes it reconstructs both sides and returns its input unchanged on any mismatch (ADR-0008). There is deliberately **no config key** for it: a second authored order doubles the surface of the one artefact that has to round-trip through Word, and presentation config belongs in the style profile, not beside the behavioural keys in §7.5.

  **Two authoring sites are NOT reached by this, and both are stated rather than left to be discovered.** The order is decided on the `WordDiff` stream, and a `replace` pair whose records carry `.table` never gets one — Stage 3 skips the word diff for a table unit (§9.3, ADR-0027), so the op's `wordDiffs` is `undefined` and stage 4a passes over it. Consequently:
  - a **changed table cell** is authored old-struck-then-new-inserted inside its one `<w:p>` (`output/author-table.ts → changedCellParagraph`) — the del-first order, in the one place it survives;
  - a **table paired with a paragraph** (a kind mismatch) emits the whole deletion then the whole insertion (`output/author.ts`), which is the *block*-level shape this rule never governed — the same shape as an ordinary `delete` op followed by an `insert` op, and as ADR-0026's demotion.

  On the Markdown-table path (`io/text/extract-text-records.ts`) that is how it stays. A Word table unit carries `.table` too (ADR-0046), but a Word pair is routed to `output/author-word-table.ts` before either site above, and its cells are analysed through every stage including 4a, so **a changed Word table cell reads insertion first**: the cell exception is confined to Text-vs-text Markdown tables. The block-level shapes read struck-first on BOTH paths, exactly as a `delete` op followed by an `insert` op does: a table paired with a paragraph, a Word row whose layout changed (struck row, then inserted row), and a Word table authored whole by the self-check's fallback (struck table, then inserted table). Whether the cell case should follow §5.5's order is an open question, not a settled exception: it was found by review, and answering it means editing code held behind a live-host gate. Recorded in ADR-0043 § Accepted limits. Because this path **rebuilds** the paragraph — from source runs or, failing that, from `WordDiff`s, which derive from prefix-free `normText` — it must re-emit the manual number (`ParaRecord.numberingPrefix`) itself (§5.0 "restore the prefix in output"; ADR-0028 as amended). Without it an edited clause loses its number while untouched ones keep theirs, i.e. numbering disappears exactly on the clauses under discussion. The same number on both sides is one leading **unmarked** run (written from its own source runs where the paragraph is rebuilt from them); a number that **changed, appeared or went** with the edit is the revised number inserted and the original struck, insertion first — written unmarked, reject-all would keep the new number on the old clause. No new op kind, and the classification is unchanged. `equal` must **not** re-inject: it emits `rawXml`, which already carries the prefix.
  **What is not text survives the rebuild (ADR-0045).** A one-to-one `replace` whose paragraphs carry a footnote reference, hyperlink, field, content control, equation, drawing, text box or bookmark is authored by `output/inline-objects.ts`, which maps each word-diff character back to the run or object it came from. An object whose text the diff left untouched, that lines up with an object of the same sort on the other side, that holds the same objects inside it (a link that gained a footnote is not unchanged), and that owns every space in its XML (none collapsed into a neighbour's or trimmed at the paragraph's edge), is re-emitted **verbatim** from the revised paragraph. A link, field or content control whose text changed keeps the revised **shell** (a field's `begin…separate` and `end`) with the word diff authored inside it, in its first run's properties — provided the rebuilt part holds nothing but ordinary runs on both sides and its characters are one unbroken stretch of the diff (checked until nothing more breaks, since replacing one pair whole re-homes a character it shared with the next). A space that is only the boundary read around an equation or text box belongs to no run or object — it is nobody's edit — and an unchanged object's own edge space that lines up with such a boundary on the other side is struck beside the object. An `eq` space no object can author — lined up with a character of another object that does not correspond, or across a deletion or insertion — is split into a deletion and an insertion, and the objects paired again. Anything else — an equation, a text box, a container with a footnote, field or equation inside it, an object with no counterpart — is emitted **whole**: the revised object inserted, the original struck, insertion first. Zero-width objects (a footnote reference, an image) are kept from the revised side when both sides carry the same sequence of them, none holds a space (a link of only a space), and each pair has one place common to both sides — a footnote that moved, or whose neighbouring words were replaced on both sides, is marked, inserted where it is and struck where it was — and otherwise all marked; the walk's own revisions are put back in §5.5's order — in every run of inserted and deleted pieces (a bookmark inside it keeping its place among the insertions), insertions first, ids re-dealt — except across a kept footnote reference or image, which would be misplaced in the rejected view (ADR-0045 § Accepted limits 13); bookmarks are kept from the revised side and never revised. The result is checked before use — accept-all must read as the revised paragraph and reject-all as the original, both in `normText` (the diff's own model — a space added beside an equation is absorbed there, ADR-0045 § Accepted limits 12) and in a census of the objects text cannot see (footnote and endnote references, fields, links, equations, drawings, embedded objects, content controls) — and any paragraph that fails, or holds a shape the unit model does not cover (a field crossing a paragraph mark, an object holding part of a stripped manual number), is authored by the text rebuild as before.

  **An edited paragraph keeps what did not change (ADR-0045).** Every one-to-one `replace` — with objects or without — is first authored **from source runs**, through the same character map (`textLeaves` in `preprocess/inline-units.ts` names the run, tab, break or silent hyphen each character came from): each kept and inserted character is written from the **revised** paragraph's own run, each struck one from the **original**'s, every run's `rPr` and attributes kept — a run split at a diff boundary keeps its `rPr` on both pieces; a formatting revision the input carried (`w:rPrChange`) is not copied, so ids stay unique and no other party's revision stays live — and every character as the document has it (a curly quote, a non-breaking space, an en dash, a tab, a break, a non-breaking or soft hyphen), not §5.2's fold of it. A collapsed whitespace run writes every source character; the revised paragraph's trailing whitespace goes out with an inserted last character and otherwise ends the paragraph unmarked. Plain text keeps the word diff's order as stage 4a set it; only a paragraph with an object has the walk's own revisions put back insertion first. The rebuild is kept only when its dry run **reconstructs both documents exactly** — accept-all and reject-all against the two source paragraphs, character for character and run-formatting for run-formatting, by `output/reconstruction.ts`, the comparison the test yardstick makes — with two tolerances, each proved, never assumed: an unrevised character accept-all showed right may read in the revised spelling in reject-all where the documents spell an `eq` span differently under §5.2's fold (one authored character cannot satisfy both views), and its formatting may differ in reject-all where the documents differ only in formatting, which Compare does not detect (§12). Failing that the paragraph takes the object-aware walk above (a paragraph with objects) and then the text rebuild: a fallback, never a loss (ADR-0008).

  For `insert` and `delete`, a paragraph carrying a hyperlink, simple field, equation or text box is revised **inside** them — `<w:hyperlink><w:ins>…`, `<m:oMath><w:del><m:r>…`, the box's own paragraphs marked too — because `CT_RunTrackChange` may not contain a `w:hyperlink` or `w:fldSimple`; every other paragraph keeps its single wrapper. A deleted field's instruction becomes `<w:delInstrText>` (ECMA-376 §17.16.13).
- **A Word table (ADR-0046).** An `equal` table unit is its `rawXml`, byte for byte. A changed pair (an op with `tableDiff`) is re-assembled from the **revised** `tblPr`, `tblGrid`, `trPr` and `tcPr`, every cell's content authored from its own ops by this same author with one shared `AuthorContext` (unique `w:id`s across every nesting level), so a changed cell reads as any changed paragraph does — insertion first included (ADR-0043). An inserted row carries `<w:ins>` in its `trPr` and its content inserted; a deleted row `<w:del>` and its content struck, keeping the original `trPr`/`tcPr`. A column added or removed (§9.3) marks each of its cells `<w:cellIns>`/`<w:cellDel>` and the grid is built from the column map. A whole table inserted or deleted, or one paired with a paragraph, keeps its structure with every row and paragraph marked. Every authored cell ends with a paragraph. **Self-check (ADR-0008):** before use, accept-all of the authored table must equal the revised table and reject-all the original in tables, rows, cells, each cell's merges (`gridSpan`, `vMerge`), its text and its zero-text objects; a table that fails, or that holds a wrapper the model does not represent (`WordTableModel.supported`), is authored as the whole original table struck followed by the whole revised table inserted — structure kept, nothing lost. A blank paragraph written self-closed (`<w:p/>`, which is how Word writes the one a cell must end with) is opened before any revision is placed in or on it — read as an open tag it put the revision round the whole paragraph, which `CT_RunTrackChange` does not allow.
- `move` → `<w:moveFromRangeStart/End>` + `<w:moveFrom>` at the source, `<w:moveToRangeStart/End>` + `<w:moveTo>` at the destination, sharing a `w:name`.
- `replace` spanning **more than one paragraph** on either side (a merge or split — §5.1, ADR-0032) → the same run-level interleaving — **except the mark-order swap above, which skips any window containing a `PARA_SEAM`**: the seam is a paragraph MARK, so moving a run across it moves that run into a different `<w:p>` (ADR-0043) — split back into one `<w:p>` per record at each `PARA_SEAM`, with the paragraph **mark** itself marked where the seam was struck or inserted: `<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>` for a lost paragraph break, `…<w:ins …/>…` for a new one. This is the revision Word records for a paragraph join, and it is what makes the merge **rejectable**: accepting it joins the paragraphs, rejecting it restores the break. Authoring a merge without it would make the break simply vanish. `CT_PPr` permits **one** `<w:rPr>` and orders it `rPr → sectPr → pPrChange`, so the mark is **merged into** an `rPr` the paragraph already has (a paragraph mark carrying direct formatting, or one emptied by §9.2 quarantine) and otherwise **inserted before** any `sectPr`/`pPrChange`; inside `CT_ParaRPr`, `w:ins`/`w:del` come first. `w:pPr` **nests** — `w:pPrChange` contains a complete `<w:pPr>` of its own, a `CT_PPrBase` with no `rPr` at all — so the leading `<w:pPr>` must be matched by depth-counting, not by a lazy close-tag match: landing the revision inside `w:pPrChange` leaves the paragraph mark unmarked and the merge **unrejectable** (ADR-0032, correction 3). A range whose records hold no object is first rebuilt **from source runs** as a replace is (above), cut at each seam, under the same dry-run check over the whole range; a range with an object, or one that fails the check, is rebuilt from text, its first paragraph's manual number marked as a replace's is.

Every revision element carries `w:author` (Config-driven, default `"Mutatis Compare"`) and `w:date` (run timestamp, ISO-8601). IDs (`w:id`) are unique and monotonically assigned. The resulting package **renders identically to Word's native tracked changes** and supports accept/reject in the standard Review UI. Summary annotations (§8) are injected as a leading content block via `insertOoxml`, *outside* the revision stream, so accepting/rejecting changes doesn't disturb them (and the user can delete the block in one action).

### 5.6 Classifier
Runs over the final op list; see §6.

---

## 6. Change classification taxonomy

Rule-based, derived purely from OOXML inspection + regex on extracted text. **A change may match several categories** (`DiffOp.classes`); the **primary** category (`classification`) is chosen by the priority order below (top = highest), so "substantive" always wins over "formatting" when both apply.

| Priority | Category | OOXML / text signal | Notes |
|---|---|---|---|
| 1 | **Structural** | Level-1 op is `insert` / `delete` / `move` of a whole paragraph; table row/cell add/delete (§9.6). | Always surfaced. |
| 2 | **Obligation language** | Word diff adds/removes/swaps a modal keyword: `\b(shall|must|may|will|should|shall not|may not)\b` (case-insensitive). | **High** legal significance; `shall→may` is a flagship case. |
| 3 | **Numerical value** | A `WordDiff` whose `atomic` ∈ {amount, date, percent, duration} changed; **any** value change inside a **table cell** (§7.3). | Amounts `S$`/`£`/`$`, dates, `%`, durations. |
| 4 | **Defined-term change** | Change touches a quoted term or a Title-Case multi-word term, or alters the definition paragraph of one. | Quoted/Title-Case detection. **The test is the CHANGE, not its neighbourhood** (ADR-0033) — a term merely *mentioned* elsewhere in the paragraph does not qualify. Title-Case matching catches a sentence-initial capital plus a capitalised word (`The Vendor shall …`), so a paragraph-wide test fires on almost every clause in an SG/E&W contract and, because this row outranks row 5, stops a pure renumbering ever classifying as `crossRefRenumber` — and therefore ever joining a cascade. The one paragraph-level limb is the *definition* paragraph itself (a quoted term followed by a definitional verb), where any edit is an edit to the definition. |
| 5 | **Cross-reference renumbering** | `WordDiff` is *only* a change to reference tokens `\b(Clause|Schedule|paragraph)\s+[\dA-Za-z().]+\b`, and the numeric delta is consistent with a document-wide shift. | **Collapsed** — see below. |
| 6 | **Substantive text** | Any remaining run-text insert/delete/replace not caught above. | Default for real wording edits. |
| 7 | **Whitespace / punctuation** | `normText` identical after aggressive normalisation (case/punct/space) — only spacing/punctuation differs. | Filterable noise. |
| 8 | **Formatting-only** | Change is **entirely** within `<w:rPr>` / `<w:pPr>` (font, bold, italic, colour, spacing, indent, style id) with **zero text delta**. | Filterable noise (but see §7.1 — en-GB/US spelling is NOT formatting). |

> **Row 6 includes a zero-text object** (ADR-0045 § Amendment): a footnote or endnote reference, picture, text box or embedded object added, removed or moved in a paragraph whose words did not change is a run inserted or deleted, so it is `substantiveText` — never `formatting`, which *Hide formatting-only* hides by default (§8.2).

> **Implementation note — row 8 has four limbs and two are built.** A change to a paragraph's **style id** or to its **resolved numbering** produces a `property` op and is classified `formatting` (§4.2, ADR-0042). A change confined to **direct `<w:rPr>`** (bold, font, colour) or to **direct `<w:pPr>`** (spacing, indent) produces no op at all: measured, a bold-only and an indent-only pair each give `summary.total` 0 and `ops` `[equal]`, at this commit and at the base commit alike. So **§10 T2's bolded heading is not yet met** — the change is not hidden by the filter, it is not detected. Building it is a fifth property class and needs its own noise answer first, because Word rewrites `<w:rPr>` on almost every paragraph of a real redline and ADR-0008 forbids the answer being "hide it". Owned in ADR-0042 § Open with the measurement. *(Recorded here because Constitution §3 makes this file the source of truth, and §8.2 already carries the same kind of note for §9.9's half-built filter. Found at the twentieth review pass on 08b; that it was recorded everywhere except the spec was found at the twenty-seventh.)*

### 6.1 Renumbering-cascade collapse

The malpractice-grade pain point: inserting one clause shifts every downstream cross-reference, producing hundreds of "Clause 4.2 → Clause 4.3" changes that bury the *one* substantive edit.

**Collapse logic:**
1. Gather every op classified `crossRefRenumber` for a given token kind. For a text op that is its `WordDiff` pair and the kind is a reference noun (`Clause` / `Schedule` / `paragraph`, from `compare.referenceTokens`). A `property` op has **no** `wordDiffs` — its two paragraphs are word-for-word identical — so its kind and shift come from `PropertyChange.cascade` instead, and the kind is `"<primaryUnit> numbering (level N)"`, which keeps a clause whose own NUMBER moved in a different group from a reference TO it (ADR-0042). Without that branch, inserting one sub-clause into a list of thirty puts thirty uncollapsed rows in the change list and produces no cascade at all.
2. Compute the numeric `shift` (old→new). Group those sharing the same `tokenKind` **and** the same `shift` (e.g. all `+1`).
3. If the group size exceeds a threshold (Config `compare.renumberThreshold`, default ≥ 3) and the shift is monotonic/consistent with a structural insert/delete detected at Level 1, emit **one** `CascadeAnnotation` (`affectedCount = N`, two or three `exampleRefs`) and set each member's `collapsedInto`.
4. In the **redline output the individual `<w:ins>`/`<w:del>` markup is still written** (so the document is correct and accept/reject works), but in the **summary panel** the cascade shows as a single line: *"Clause references renumbered (+1): 47 — example: Clause 4.2 → 4.3"* (the reference noun leads, because it comes from `compare.referenceTokens` and must read as the jurisdiction's own). A numbering cascade reads *"Clause numbering (level 1) shifted (+1): 3 — example: 1.1 → 2.1"*; the level named is the one whose COUNTER moved, not the paragraph's own — under `%1.%2` a sub-clause renumbers because its parent did, which is why §4.2 reads this same example as *"a level-1 paragraph carried by a level-0 shift"* (`shiftedIlvl` 0, printed 1-based); `CascadeAnnotation.source` distinguishes the two (`'crossReference'` / `'numbering'`) so one is never described as the other. The "collapse renumbering" filter (default ON) hides the individual entries from the navigation list.

This is collapse-in-presentation, never suppression: the changes remain in the document and are one toggle away in the list.

---

## 7. SG/UK conventions & config keys used

Redline mechanics (`<w:ins>`/`<w:del>`) are jurisdiction-neutral. The **classifier** is tuned for SG and UK (England & Wales) via the [Localisation / Config Layer](../ARCHITECTURE.md#localisation--config-layer); the variant map and terminology values live in [en-GB Spelling](../CONVENTIONS.md#en-gb-spelling) and [Terminology](../CONVENTIONS.md#terminology).

### 7.1 en-GB ↔ en-US spelling is surfaced, not filed as formatting

A change of `organize → organise`, `-ize → -ise`, `color → colour`, `-or → -our`, `program → programme` etc. is a **real text delta**, so it can **never** be classified `formatting-only`. It is surfaced as `substantiveText` and additionally tagged with a `spellingVariant` sub-flag so the summary shows *"3 en-US→en-GB spelling changes"* as its own filterable line. Detection: the bidirectional en-GB/en-US variant map in [en-GB Spelling](../CONVENTIONS.md#en-gb-spelling). **Surfaced regardless** — the rule errs toward visibility.

### 7.2 Signature / execution-block changes are ALWAYS substantive

UK deed execution blocks ("EXECUTED as a DEED", "in the presence of", witness attestation) and SG execution blocks (common-seal / director-secretary signing) carry legal weight wildly disproportionate to their text size. Any change inside a paragraph flagged `isSignatureBlock` (heuristics: proximity to *EXECUTED, SIGNED, as a DEED, in the presence of, Director, Secretary, Witness, Common Seal, Authorised Signatory*; position in last region of the body; signature-line glyphs) is forced to `classification = 'substantiveText'` (with `structural` if the block was added/removed) and is **never** filterable as formatting/whitespace.

### 7.3 Table-cell value changes are ALWAYS substantive

Term sheets and schedules carry the commercial deal in tables. Any text change inside a `<w:tc>` (table cell), and especially any `numericalValue` change in a cell, is forced to `numericalValue` / `substantiveText` and is **never** filterable as formatting/whitespace — even if the *only* visible difference looks minor.

### 7.4 Reference tokens

The renumbering-collapse and cross-reference logic uses **SG/UK** tokens: **`Clause`**, **`Schedule`**, and **`paragraph` (of a Schedule)** — **not** the US `Section` / `Exhibit` (per [Terminology](../CONVENTIONS.md#terminology)). The token list is in Config so a firm can extend it (e.g. `Recital`, `Annex`, `Appendix`); the SG/UK defaults ship as above. (`Section` may be added by firms drafting statutes, but is off by default to avoid mis-collapsing statutory section references.)

### 7.5 Config keys (feature-local)

```
compare.author            (string, default "Mutatis Compare")
compare.referenceTokens   (["Clause","Schedule","paragraph"] — a DETECTION set: which
                           reference nouns to recognise in prose. Unordered.)
compare.primaryUnit       (string — the jurisdiction's noun for a top-level body
                           division, used where Compare must NAME the unit rather than
                           recognise it: the §6.1 numbering-cascade heading. DEFAULTS TO
                           the profile's `xref.units.primaryBody`, which CONVENTIONS
                           § Terminology makes the single source of truth for that noun;
                           set it here only to make Compare differ from the
                           cross-reference layer. Kept separate from referenceTokens
                           because that is an unordered DETECTION set, and naming the
                           primary unit by taking its first element is a coincidence.)
compare.renumberThreshold (int, default 3)
compare.replaceSimilarityThreshold (float 0–1, default 0.5 — §5.3 change-complexity threshold; 0 disables)
compare.spellingVariants  (en-GB/en-US map; surfacing ON)
compare.excludeHeaders    (bool, default FALSE — §9.7, ADR-0047)
compare.excludeFooters    (bool, default TRUE)
compare.footnotes         ('compare' | 'ignore', default 'compare')
compare.flattenOriginal   ('acceptAll' | 'rejectAll', default 'acceptAll')
compare.flattenRevised    ('acceptAll' | 'rejectAll', default 'acceptAll')
```

> **v2 relevance note:** because SG/UK counterparties routinely send **execution versions as PDF**, the v2 Word-vs-PDF compare (§2) is a mainstream workflow, not an edge case — and PDF.js keeps it zero-setup/zero-upload, consistent with the product's confidentiality guarantee.

---

## 8. UI / task-pane behaviour

### 8.1 Run-compare flow

1. **Compare** tab. Top: two slots — **Original (base)** and **Revised (compared)**.
2. Each slot is **"Use current document"** (→ `getFileAsync`) or **"Upload .docx"** (→ File API). User must pick which is which (direction control, §3.2).
3. Per-side **"Flatten existing tracked changes"** control: *Accept all* / *Reject all* (default Accept all), shown only when that side actually contains revisions (detected on load and warned: *"Revised contains 18 tracked changes — they will be accepted before comparison"*).
4. **Run Compare** → progress indicator (chunked progress for large docs, §9.12).
5. Output: a **new redlined document**. Default = open in a new Word window (dialog / `createDocument`); option = insert into the active document (when the active doc is one of the inputs and the user wants the redline in place).

### 8.2 Summary panel

After a run, the right side of the pane shows the **ChangeSummary**:

- **Total changes**, then **counts by category** (Structural, Obligation language, Numerical, Defined-term, Cross-reference renumbering, Substantive text, Whitespace/punct, Formatting-only), plus dedicated lines for **Moved paragraphs** and **en-GB/US spelling**.
- **Filter toggles:** *Hide formatting-only* (default ON), *Collapse renumbering cascades* (default ON). Toggling a filter only changes what's shown in the navigation list — **the redline document is never altered** by a filter.

  There is deliberately **no** *Hide whitespace/punctuation* toggle (ADR-0036). `whitespacePunct` (§6 priority 7) requires a real text delta, and `substantiveText` (priority 6) *is* "has a real text delta", so the class can never be an op's primary classification — a filter reading `op.classification` had nothing to hide. Rewiring it to read `op.classes` was rejected rather than adopted: an op carries both classes, so hiding on the secondary would hide a change that is also substantive, and in SG/E&W drafting a comma governs the scope of a proviso. Whitespace and curly-quote differences never reach a filter at all — §5.0 normalisation absorbs them before the diff, so they produce no op. The class is still recorded in `op.classes` and reported; nothing filters on it. **Consequence:** the §8.2 `byCategory` counts are keyed on `classification` so they partition `total`, which leaves the Whitespace/punct bucket permanently `0` in the summary panel.
- **Cascade rows** render collapsed: one line per cascade with `affectedCount` and example refs; expandable to show members. A numbering cascade (`source: 'numbering'`) and a cross-reference cascade render as different sentences — see §6.1 step 4.
- **Other parts** — a separate block below the categories, present only when there is something in it. It holds two kinds of line and they behave differently, so the panel says which is which rather than letting the reader assume:
  - the header and footer change counts (§9.7) and the footnote count (§9.8) **are** counted into `total` and `byCategory` like any other op (`mergePartsIntoSummary`), so they are a breakdown of changes already in the totals and never something to add to them. The block carries that sentence. What is outstanding for them is **navigation**, not counting: a part op's range indexes its own part's paragraphs, so clicking such a row would scroll to an unrelated body clause.
  - the **style-definition note** (§9.9) is **not** counted. `diffStyles` produces a synthetic formatting-only op, the driver keeps its note and discards the op, so a `styles.xml` change is always reported here and is absent from `total`, from `byCategory` and from the *Hide formatting-only* filter — §9.9 asks for both halves and only the reporting half is built (ADR-0042 § Open, on its recorded cost: a style definition is not at a paragraph, so counting it adds a number with no row to click). Nothing is suppressed; the line is always visible. The row says *"not counted above"* on itself, because it can be the only line in the block.

  Headers and footers are excluded by default (§7.5); footnotes are **not** — `compare.footnotes` defaults to `compare` because §9.8 makes a footnote content change substantive — so this block can be non-empty with nothing switched on.

### 8.3 Navigate changes

- A scrollable **change list** (one row per non-hidden `DiffOp`) with category icon, a short text preview, and (for `move`) a "moved from/to" indicator. A zero-text object change (§5.1) has no changed words, so its preview names the objects, struck and inserted — `-[footnote reference]`, `+[picture]`, and a moved one as both; the exported report prints each side's paragraph text with its objects written in at their places. *Go to change* anchors on the paragraph's own words, as for any `property` op; a paragraph with no words (a picture alone, emptied) has nothing to search for and its row does not navigate.
- A **Word table** op (ADR-0046) is one row for the whole table. Its preview is the first changed cell's own preview, followed by `(+N more in this table)` when the table carries other changes — rows, cells, or further cell edits — so one cell never reads as the whole of it; a table changed only by rows or columns previews its shape (`Table: 1 row inserted, 1 column added`). *Go to change* aims inside one cell, since Word's search cannot cross a cell boundary: the first changed cell's own word diff, else an inserted row's first non-empty cell, else — for a deletion, whose words are gone — the first surviving row's. The exported report prints the whole table's text on each side.
- Clicking a row uses `Word.run` + range navigation to **scroll the redlined document to that change** and select it. Next/Previous step through the *filtered* list.
- Signature-block, table-cell, and spelling-variant items carry a small "always substantive" / "house-style" badge so the user understands why they were not filtered.

---

## 9. Edge cases

### 9.1 Paragraph reordering — anchor matching before LCS
Plain LCS on reordered content yields misleading alignments. We pin **unique-hash anchors** shared by both sides first, then LCS only between anchors (§5.1). Stabilises reordering and bounds cost on long documents. What the LCS leaves unmatched between two matches is paired by similarity, never across a match (§5.1 step 4, ADR-0039), so one paragraph inserted or deleted among changed paragraphs no longer shifts the pairing of the rest.

### 9.2 Documents with pre-existing tracked changes
Flatten (accept-all/reject-all) per side **before** diffing; quarantine residual revision/comment markup; report (never silently merge) what was flattened/excluded (§5.0).

### 9.3 Table structural changes
Adding a column shifts every cell's column index, so naive cell-by-cell position diffing reports the whole table as changed. **Diff tables by position in the *new* structure**: align rows first (row-hash LCS), then within aligned rows align cells; a genuinely added column surfaces as **structural** (new cells = `insert`), not a wholesale rewrite. Row insert/delete is `structural`. Cell *value* changes are always substantive (§7.3).

**Word-vs-Word (ADR-0046, stage 1b).** For each op pairing two Word table units — a `replace`, or an `equal` whose XML differs:

1. **Rows.** LCS on the row hashes; each gap between two matches is paired by position when both sides have the same number of rows, otherwise order-preserving by similarity (the higher of the rows' text Dice and the share of identical cells, at `replaceSimilarityThreshold`), leftovers whole-row inserts and deletes. **When the grid's width changed** (its `tblGrid` column count), every row's hash changes with it and the LCS would find nothing, so the whole table would pair by position — and a row deleted plus a row inserted elsewhere would then pair every row between them with its neighbour. Rows that are **full width** (one unmerged cell per grid column) are therefore keyed by their cells with the wider side's surplus columns left out, the subset left out being the one under which the most rows agree (up to 64 subsets tried); any other row keeps its own hash.
2. **Cells.** A paired row whose layout is identical (same cell count, same `gridSpan` and `vMerge` per cell) is compared cell by cell. A cell whose content is identical paragraph for paragraph gets no ops; any other cell's content goes through stages 1–4a — the same structural diff, gap pairing, merge/split, demotion, property detection, word diff, classification and mark order as the body — recursively for a nested table.
3. **Columns.** When the width changed, a column map is read from the paired rows that are full width on both sides: columns pair order-preserving, keeping the most of those rows' cells identical, and the wider side's surplus columns are the ones added or removed. A paired row that is full width on both sides, and whose mapped cells are merged down (`vMerge`) the same way, is compared through the map — its added or removed cells `insert`/`delete` slots. A row merged across the table (a heading, a total) cannot take the map.
4. **Layout change.** A paired row that fits neither — cells merged or split, a vertical merge changed, a merged row whose merge widened with a new column — is a whole-row delete followed by a whole-row insert. Always valid, never suppressing, and the reader sees both rows.
5. **Classification.** The union of the nested ops' classes, plus `structural` when a row or cell was inserted or deleted (§6, §7.3 unchanged). One change item per table (ADR-0027's granularity).
6. **An equal table still looks inside.** A text-equal table whose cells changed what text identity does not anchor on — a cell paragraph's style or numbering, a footnote added in a cell — becomes a `property` op carrying the `tableDiff`, its `propertyChange` the first such change with no cascade marker, so §6.1 never collapses a table into a renumbering run.

ADR-0026's demotion is never applied to a table unit — it would strike and re-insert a whole schedule for one changed amount — and runs inside each cell, except on a pair including the cell's **last** paragraph: that paragraph's mark can never carry a revision (§5.5), so demoting it would leave a blank paragraph in the cell whichever way the reader resolves it.

### 9.4 Cross-document numbering-scheme differences
The two documents may use different auto-numbering definitions (different `numbering.xml`). Because Level 1 hashes the **prefix-stripped** text (§5.0), a clause numbered `4.2` in Original and `4.2.` (different scheme) in Revised still hash-matches on body text; the *numbering* delta is handled as a `crossRefRenumber`/formatting signal, not a content change.

**Two `<w:num>` over one `<w:abstractNum>` are ONE sequence.** Word continues the count across numbering instances that share a definition and restarts only where the instance carries `<w:startOverride>` — which is what that element is for. While the renderer keyed counters per instance, a clause inserted into the first group moved the second group's numbers in Word and Compare reported **nothing**, and a clause moved between instances produced a renumbering cascade over numbers Word never renders (ADR-0042 § Accepted limits One; exec-plan 08b, pass 25).

**A list attached to a Word LIST STYLE resolves like any other.** Word puts the level definitions on a second `<w:abstractNum>` carrying `<w:styleLink>` and leaves the one the `<w:num>` points at holding only `<w:numStyleLink>` (ECMA-376 §17.9.21/§17.9.28); `resolveLevel` follows that delegation. It is the shape a firm template usually uses for house numbering, and while the resolver did not follow it a redefinition of such a list produced **no op at all** — ADR-0008 suppression, not a filter (ADR-0042; exec-plan 08b, pass 24).

### 9.5 (folded into 9.3.)

### 9.6 Table cell/row/column ops
See 9.3. Implementation reuses the Level-1 LCS machinery on row hashes and the whole body pipeline on each cell's content (`diff/word-table.ts`); the table is ONE `ParaRecord` (`wordTable`, §4.1) whose cells hold their paragraphs as records of their own, each with `isTableCell=true`. Authoring is `output/author-word-table.ts` (§5.5).

### 9.7 Headers / footers
Headers/footers are **separate OOXML parts** (`header1.xml` …) and are frequently metadata (page numbers, doc IDs, draft stamps) that change every save and would generate pure noise. **Configurable, compared by default** (`compare.excludeHeaders/Footers = FALSE`, ADR-0047). They are diffed as their own parts, paired by role, authored into their own parts and listed with a location badge so a *substantive* footer change (e.g. a changed governing-law footer stamp) can still be caught on demand.

### 9.8 Footnotes
Footnotes live in a **separate part** (`footnotes.xml`). **Content** changes to a footnote are **substantive** and compared by default. Footnote *renumbering* (insert one footnote → all later numbers shift) is **noise**, handled like the cross-reference cascade — collapsed, not surfaced per-number. `compare.footnotes = 'ignore'` disables footnote comparison entirely for noisy docs.

A footnote **reference** added, removed or moved in an otherwise unchanged body paragraph is a `property` op with `objects` (§5.1, ADR-0045 § Amendment): reported, classified substantive, and marked in the redline. The reference's place is compared, never its `w:id`, so the renumbering above produces no body op. Which footnote a kept reference points at is not compared from the body (ADR-0045 § Accepted limits 9a).

### 9.9 Style-definition changes
Edits to `styles.xml` (a style's font/spacing redefined) can change appearance with **no body text delta**. Detected by diffing `styles.xml`; reported as **formatting-only** with a distinct "style definition changed" note (filterable), and per-paragraph `styleId` changes are likewise formatting-only — except where they alter a signature/structural paragraph.

### 9.10 Very large documents (200–400pp credit agreements)
Two whole packages in memory plus parsed trees is heavy, and `getFileAsync`'s slice reassembly has practical size ceilings ([Performance Guidelines](../ARCHITECTURE.md#performance-guidelines)). Mitigations:
- **Stream/chunk** unzip+parse (JSZip async) rather than loading everything synchronously.
- **Hash-first, lazy-detail**: compute all paragraph hashes and Level-1 alignment before any Level-2 word diff; run Level-2 only on `replace` pairs, lazily as the user scrolls the change list.
- **Lazy-load the diff UI** (virtualised list).
- **Mind the two-docs-in-memory `getFileAsync` limit** ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 12) — if both sides are huge, prefer the File-API path for the second doc (it streams from disk) and free each `ArrayBuffer` after parsing. Surface a soft warning past a size threshold.

### 9.11 (folded into 9.10.)

### 9.12 Progress & cancellation
Long runs report chunked progress (parse %, align %, word-diff %) and are cancellable; partial state is discarded cleanly.

### 9.13 `insertFileFromBase64` fidelity
`insertFileFromBase64` can drop or alter some package-level constructs (certain styles, numbering links, section properties) when *inserting* a built package into an existing document ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 11; [The OOXML Service](../ARCHITECTURE.md#the-ooxml-service) fidelity caveats). For Compare output we therefore **prefer materialising a standalone new document** via the "create new document" path (open the assembled `.docx` as its own document) and use `insertFileFromBase64` only for the *insert-into-active-doc* option. Where in-place insert is requested, validate the round-trip (re-read the inserted content's OOXML and diff against intended) and fall back to opening a new document on detected fidelity loss. Document this caveat in-product (§12).

---

## 10. Test cases / acceptance criteria

| # | Scenario | Expected result |
|---|----------|-----------------|
| T1 | **Inserted clause** mid-document (Original 40 clauses → Revised 41, one new Clause 4.3). | Exactly **one** `insert` (the new clause) as `<w:ins>`; **no cascade misalignment** — every following clause is `equal` at body level (only its number shifts). Renumbering handled by T3. |
| T2 | **Formatting-only** change — a heading bolded, no text delta. | Classified `formatting`; with default filters **hidden** from the change list; redline document still contains the `<w:rPr>` revision so it's accept/rejectable. |
| T3 | **Renumbering cascade** — the T1 insert shifts 47 downstream "Clause N" references. | Summary shows **one** `CascadeAnnotation` ("+1, 47 Clause references", with examples); the 47 individual entries collapsed (default ON) but still present in the output document. |
| T4 | **Moved paragraph** — an indemnity clause relocated verbatim from §12 to §3. | One `move` op pair; Word shows `<w:moveFrom>`/`<w:moveTo>`; **not** reported as delete+insert; summary "Moved paragraphs: 1". |
| T5 | **`shall` → `may`** in an obligation. | Word diff atomic-flags the modal swap; classified `obligationLanguage` (substantive); surfaced with the high-significance badge; never filtered. |
| T6 | **en-GB ↔ en-US spelling** — `organize → organise`. | A real text delta; classified `substantiveText` + `spellingVariant`; **surfaced**, never `formatting`; summary shows it on the en-GB/US line. |
| T7 | **Pre-existing tracked changes** in Revised. | User warned; flatten (accept-all default) applied before diff; comments quarantined and reported; output redline contains only *our* new revisions. |
| T8 | **Table cell amount** `S$1,250,000 → S$1,500,000`. | `numericalValue`, forced substantive (§7.3); never filtered even though "tiny" diff; atomic amount token (not digit-level). |
| T9 | **Signature block** witness line edited. | Forced `substantiveText`, "always substantive" badge; never filterable (§7.2). |
| T10 | **Header page-number metadata** differs every save. | Headers are compared by default (ADR-0047): the change surfaces in its header part, classified by content (fields compare by cached result); with `excludeHeaders=TRUE` it produces **zero** changes. |
| T11 | **Round-trip render** of the output `.docx`. | Opens in stock Word; every change is individually accept/rejectable via the Review ribbon; document is byte-valid OOXML. |
| T12 | **Large doc (300pp)** with ~500 changes. | Completes without OOM; Level-2 word diff runs lazily; change list virtualised; progress + cancel work. |

Acceptance gate: T1–T11 pass deterministically; T12 passes within an agreed time/size budget on the reference machine; **no test** demonstrates a change being *suppressed* (only filtered).

---

## 11. Reference architecture

**Reference (not code): LibreOffice Writer document comparison (MPL-2.0).** Its comparator solves exactly our hard edge cases — paragraph-level alignment, move detection, table and numbering handling, redline emission. We treat its **algorithm design as transferable reference** (a proven legal-grade comparator) but **do not port its C++/UNO code** (MPL file-level copyleft + different runtime). Our pipeline (§5) mirrors its layered approach in TypeScript over OOXML. The supporting library roles/licences are in §3.1 and [Technology Stack](../ARCHITECTURE.md#technology-stack).

---

## 12. Known limitations

1. **v1 is Word-vs-Word only.** No PDF comparison until v2 (PDF.js). Given SG/UK PDF execution copies, a real workflow gap, consciously deferred (§2, §7.5).
2. **Hash-match move detection misses edited moves.** A relocated *and* edited paragraph is reported as delete+insert (no verbatim hash match). Fuzzy moves are v2.
3. **No ML / no semantic equivalence — by design.** The classifier categorizes; it never decides two differently-worded clauses "mean the same" and hides a change. **Permanent invariant** (§2.3 CUT).
4. **`insertFileFromBase64` fidelity caveats** (§9.13) for the insert-into-active-doc path; mitigated by preferring a standalone new-document output and validating round-trips.
5. **Classification is heuristic.** Defined-term and signature-block detection are rule/regex based and can mis-tag at the margins — but because nothing is suppressed, a mis-tag at worst mis-files a change in the summary; it can never hide it. Filters default to *show more*.
6. **Two whole packages in memory** stress the `getFileAsync` size ceiling on very large pairs; mitigated but not eliminated (§9.10).
7. **A zero-text object is compared by sort and place only, and only five sorts are compared** (ADR-0045 § Amendment 1, Accepted limits). A picture swapped for another in the same place, a symbol character (`w:sym`) or non-breaking hyphen added alone, and a link or field put around words that did not change, produce no op. The change-set and AI-readable exports omit `property` ops, so they omit a footnote or picture added or removed too; the redline, the change list, the summary and the report carry it.
8. **Word tables (ADR-0046 § Accepted limits).** Table and cell formatting (`tblPr`, `trPr`, `tcPr`: widths, borders, shading) is not compared — the same class as direct run formatting. A cell merge, split or vertical-merge change reads as the whole row deleted and re-inserted, as does a row merged across the table when a column is added through it. A moved table reads as a delete and an insert. Tables in headers, footers and footnotes stay flat records (their parts are counted, not authored). One change item per table. RenderCore's preview does not decorate row or cell marks. How Word desktop accepts and rejects `<w:cellIns>`/`<w:cellDel>` is unverified until the live-host smoke test.
9. **An edited paragraph's reject-all, at its edges (measured; ADR-0045 § Accepted limits 14).** A stretch the diff calls unchanged but the two documents spell differently under §5.2's fold — a straight against a curly quote or apostrophe, a hyphen against an en or em dash, a space against a non-breaking space, a tab or break against a space — is written in the **revised** spelling, because one authored character cannot carry two spellings; reject-all shows the revised typography there. Only those fold-set characters are excused, and only in unchanged stretches — each one the view shows unrevised and proved by accept-all (a substitution is excused on that alone), and each one it lacks outright with every neighbour so, never beside a revision — and the yardstick counts them, each side of a substitution once: in a private test pair, 3 paragraphs / 15 characters on reject-all (2 paragraphs beyond whitespace), 0 on accept-all; v1.1 1 / 1 (whitespace), 0 on accept-all. Marking them would report what §5.2 folds away on purpose (corpus C30) — a §5.2/§6 design decision. A difference of formatting alone between the documents is not detected, so the redline carries the revised formatting both ways. A paragraph whose source rebuild fails its check, and a merge/split holding an object, are rebuilt from text: their kept words lose run formatting and typography, as before the source-run rebuild.
10. **A renumbered clause whose words did not change rejects to its new number (§6.1).** A manually typed number that moved because a clause above it was inserted or deleted leaves the clause an `equal` op, written from the revised side, so reject-all reads the revised number there (an edited clause's number is marked, §5.5). Auto-numbering is unaffected — Word renumbers. Measured on the readability corpus: exactly the renumbered clauses (2 each in C18, C19, C20, C28); none in the private test pairs.
11. **An input document's formatting revision (`w:rPrChange`) survives into an unchanged paragraph.** The flattener resolves inserted, deleted and moved content and paragraph marks, not formatting revisions, so an `equal` paragraph's `rawXml` keeps another party's `w:rPrChange` — live in the redline, its `w:id` possibly repeating one Compare mints (found in review: ids `1, 1, 2`). An edited paragraph rebuilt from source runs drops it (§5.5). Predates the source-run rebuild; the fix belongs in the flattener (accept-all keeps the run's `rPr`, reject-all restores the old one).
12. **A changed character outside the Basic Multilingual Plane (an emoji, a rare CJK character) is split in the redline.** The word diff's tokeniser emits non-ASCII text one UTF-16 code unit at a time (`diff/atomic-tokens.ts`), so `😀` → `😃` is authored as an unmarked high surrogate, an inserted low one and a struck low one; the package writer turns each lone half into U+FFFD. Both views rejoin the pair, so the reconstruction check (which compares UTF-16 strings) cannot see it. Found in review; predates the source-run rebuild. The fix belongs in the tokeniser: one token per code point.

---

## 13. Build milestones

| M | Deliverable |
|---|-------------|
| **M1 — OOXML in/out plumbing** | Load Original (`getFileAsync`) + Revised (File API); unzip/parse both via the OOXML Service; extract parts; emit a trivial pass-through redline (single forced change) that round-trips in Word (T11). |
| **M2 — Pre-processing** | Flatten existing tracked changes (accept/reject), quarantine comments, normalise text + strip/restore numbering; `ParaRecord` build. (T7) |
| **M3 — Level 1 structural diff** | Paragraph hashing, anchor matching, segmented LCS → insert/delete/equal/replace. Inserted-clause no-cascade behaviour. (T1) |
| **M4 — Level 2 word diff** | diff-match-patch word mode with atomic tokens; replace-pair run-level `<w:ins>`/`<w:del>`. |
| **M5 — Level 3 move detection** | Verbatim hash-match moves → `<w:moveFrom>`/`<w:moveTo>`. (T4) |
| **M6 — Classifier + taxonomy** | Full §6 categories, priority resolution, multi-class tagging. (T2, T5, T8, T9) |
| **M7 — Renumbering-cascade collapse** | Cascade detection + `CascadeAnnotation`; collapse in summary, keep in doc. (T3) |
| **M8 — SG/UK rules** | en-GB/US surfacing, signature-block + table-cell always-substantive, Clause/Schedule tokens, Config wiring. (T6, T8, T9) |
| **M9 — Summary panel + navigation UI** | Counts by category, filter toggles, virtualised change list, go-to-change. (T2, T3) |
| **M10 — Edge cases & scale** | Tables (§9.3), headers/footers (§9.7), footnotes (§9.8), styles (§9.9), chunking/lazy UI + new-document output + fidelity fallback (§9.10, §9.13). (T10, T12) |
| **M11 — Hardening** | Full acceptance suite T1–T12; in-product caveat copy; Config defaults review. |
| **v2 (post-M11)** 🔜 | PDF.js Word-vs-PDF; fuzzy moved-and-edited detection. |
