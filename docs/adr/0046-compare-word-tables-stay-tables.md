# ADR-0046: Compare — a Word table stays a table in the redline (Word-vs-Word structured tables)

- **Status:** Accepted — direction approved by the maintainer 2026-10-18; the specific choices below are listed for the maintainer to redirect.
- **Date:** 2026-10-18
- **Relates to:** ADR-0008 (never suppress), ADR-0026 (change-complexity demotion), ADR-0027 (Markdown tables — the text path this mirrors and does NOT touch), ADR-0039 (gap pairing), ADR-0042 (non-text identity), ADR-0043 (mark order), ADR-0045 (inline objects survive a changed paragraph). Spec 08 §9.3, §9.6, §7.3.

## Context

Spec 08 §9.3/§9.6 has always said Word tables are diffed row-then-cell. The Word-vs-Word path never did it. Measured 2026-10-17 on a minimal synthetic pair through `runPipeline`:

- An **unchanged** table loses its `<w:tbl>` wrapper: `extractParaRecords` walks into `<w:tc>` and records only the cell paragraphs, and `authorBody` concatenates their `rawXml`. Every table in every Word-vs-Word redline comes out as loose paragraphs — grid, widths, borders, merges all gone.
- A **changed** cell (`5 days` → `10 days`) is authored as one paragraph struck and another inserted, outside any cell.

ADR-0027 knew this (its Context, fact 2) and deferred "structured tables for Word-vs-Word" as a follow-up because it changes shipped behaviour. The building blocks exist: `computeLCS`, ADR-0039's Dice arithmetic, the ADR-0027 table-unit guards in every shared stage (word diff, move, merge/split, gap pairing, classifier), and the paragraph author with ADR-0045's object handling. What is missing is a Word table record that keeps the real XML, and an author that preserves it.

The Markdown author (`output/author-table.ts`) is **not** reusable for Word: it synthesises a clean table from cell strings, so it would discard `tblPr`, `tcPr`, widths, merges, shading and every run's formatting. It is also behind a live-host gate (ADR-0043 § Accepted limits), so it is left untouched.

## Decision

### 1. A body-level `<w:tbl>` is one alignment unit that keeps its XML

After `extractParaRecords` and `assignNumberLabels` have run over the flat record list (so numbering counters see cell paragraphs exactly as today), a post-pass groups each top-level `<w:tbl>` of the **body** part into one `ParaRecord`:

- `rawXml` — the whole `<w:tbl>`, so an equal table is re-emitted verbatim.
- `hash` — over a structural serialisation: row, cell and paragraph boundaries plus each cell paragraph's own hash. Equal hash therefore means identical text **and** identical row/cell/paragraph shape.
- `normText` — cell texts joined row by row, for display and the change list.
- `table` — the `RowModel[]` grid, so every existing ADR-0027 guard applies unchanged (no Level-2 word diff on the unit, no move candidacy, no merge/split, positional gap stretch, table classifier branch).
- `wordTable` (new) — the structure: `tblPr`/`tblGrid` XML and `supported`; per row its `tblPrEx` and `trPr` XML, its cells and a row hash (each cell's `gridSpan`, `vMerge` and hash); per cell its `tcPr` XML, `gridSpan`, `vMerge`, its content as `ParaRecord[]` (paragraphs, and nested tables grouped the same way, recursively) and a cell hash. A content control or custom XML round a row, a cell or a cell's paragraphs is read THROUGH — what is inside is in the model and the hash — and sets `supported: false` (§3). A table with a paragraph outside every cell or a cell outside every row (schema-invalid) is not grouped: a unit has no place for that paragraph's style, objects or position, so it stays flat records, compared as before this ADR.
- `isTableCell: true`; `isSignatureBlock` true if any cell paragraph is one (§7.2 survives the grouping).

Headers, footers and footnotes keep flat records — the parts diff only counts, it does not author.

### 2. A table pair is analysed cell by cell with the body's own stages

`runComparePipeline` stages 2–6.5 are extracted, byte-identically, into one function (`analyseRecords`). A new stage after the ADR-0026 demotion analyses every op that pairs two Word table units and attaches `op.tableDiff`:

- **Rows:** `computeLCS` on row hashes; inside a gap, rows are paired by position when the two sides have as many rows, otherwise order-preserving by similarity — the higher of the rows' text Dice (ADR-0039's arithmetic and threshold) and the share of identical cells — so one inserted row does not shift the pairing of the rows after it. Unpaired rows are whole-row inserts/deletes. **When the grid's width changed**, every row hash changes with it, so full-width rows are keyed past the wider side's surplus columns (the subset under which the most rows agree) — see § Measured, defect 2.
- **Cells:** an aligned row pair whose cell layout is identical (same cell count, same `gridSpan`/`vMerge` per cell) is compared cell by cell, each cell's content through `analyseRecords` — the same structural diff, gap pairing, merge/split, demotion, property detection, word diff, classification and **mark order** as the body. Nested tables recurse.
- **Columns:** when the grid's width changed, a column map is read from the paired rows that are *full width* on both sides (one unmerged cell per grid column) and applied to every paired row that is full width and whose mapped cells are merged down the same way; the added or removed column's cells are authored `<w:cellIns>`/`<w:cellDel>`. A row merged across the table (a heading, a total) cannot take the map — see § Measured, the real pair.
- **Layout change** (a merge, a split, a vertical merge changed, or a row that cannot take the column map): the row is a whole-row delete plus a whole-row insert.
- **Equal tables still look inside:** a hash-equal pair is analysed cell by cell too, so a style or numbering change on a cell paragraph is reported as today (ADR-0042). Today those paragraphs are top-level records and get `property` ops; grouping them must not lose that.
- **ADR-0026's demotion does not apply to a Word table unit.** Demotion still runs inside each cell, except on a pair that includes the cell's LAST paragraph: that paragraph's mark can never carry a revision (ADR-0035), so demoting it left a blank paragraph in the cell whichever way the reader resolved it. A table-level threshold remains the deferred follow-up.
- **Classification:** the union of the nested ops' classes, plus `structural` when a row or cell was inserted, deleted or re-laid. §7.3 (a table-cell change is always substantive) is unchanged.
- **Summary granularity:** one change item per changed table, as ADR-0027. Per-cell change-list items are a follow-up.

### 3. The author keeps the table and revises inside it

A new `output/author-word-table.ts`, routed in `author.ts` **before** the Markdown branch:

- Equal, or property-only → `rawXml` verbatim (as ADR-0042 authors a `property` op).
- Changed → the revised `tblPr`/`tblGrid`; per aligned row the revised `trPr`; per cell the revised `tcPr` and the cell content authored by `authorBody` over the nested ops, with **one shared `AuthorContext`** (unique `w:id`s, one `classByRevId`).
- Inserted/deleted row → `<w:ins>`/`<w:del>` in `trPr` (schema order: after the `CT_TrPrBase` elements, before `trPrChange`), its content wholly inserted/deleted; a deleted row keeps the original's `trPr`/`tcPr`.
- A whole table inserted or deleted, or a table paired with a paragraph → every row marked, structure kept.
- Every authored cell contains at least one `<w:p>`.
- **Self-check (ADR-0008):** accept-all of the authored table must equal the revised table, and reject-all the original, in tables, rows, cells, each cell's merges (`gridSpan`, `vMerge`), its whitespace-free text and its zero-text objects (`signature`). On failure — or when either table holds a wrapper the model does not represent (`supported: false`) — the table is authored as the whole original table struck followed by the whole revised table inserted, structure kept. There is no flat fallback: the whole pair is always valid and always reconstructs. The check runs on a scratch id counter, so a result that fails mints no ids.

### 4. Consumers that read an op

- *Go to change* must search for text inside one cell (Word's search cannot cross a cell boundary): a table op's anchor comes from its first changed cell's word diff, else an inserted row's first non-empty cell, else the first surviving row's (`diff/table-summary.ts`).
- The change-list preview of a table op is its first changed cell's preview plus `(+N more in this table)`, so one cell never reads as the whole change; a table changed only by rows or columns previews its shape. Its row carries §8.3's "always substantive" badge, as a Markdown table cell does.
- The exported report prints the whole table's text on each side (it reads `normText`, not `wordDiffs`).
- Batch (17a/17b) runs the same pipeline; its pairing sketch now sees one hash per table.
- A document with no tables authors byte-identically to before, except where a blank paragraph written self-closed (`<w:p/>`) is inserted, deleted or replaced: that used to be authored invalidly (`<w:ins><w:p/></w:ins>`) and is now revised by its mark (§ Consequences). Measured by review: 150 seeded no-table pairs plus every readability case — byte-identical wherever no such paragraph is revised.

## Rationale

- It is the spec's own design, built on modules that already pass their tests: the ADR-0027 guards, `computeLCS`, ADR-0039's alignment and the paragraph pipeline. The only genuinely new code is grouping and a structure-preserving author.
- Running the body's stages inside each cell means every readability and correctness rule the body earned (ADR-0028…0045) applies to cell text for free, including insertion-first mark order — which closes ADR-0043's table-cell gap for the Word path without touching the gated Markdown author.
- Grouping after numbering is assigned keeps every counter and label byte-identical.
- Row-level delete + insert for layout changes is always valid OOXML and never suppresses; the finer `cellIns`/`cellDel` form is attempted only where the grid can be made consistent.

## Rejected alternatives

- **Reuse the Markdown author (`author-table.ts`).** It rebuilds cells from strings and would destroy table and run formatting; it is also behind 13b's live-host gate.
- **Keep cell paragraphs as top-level records and re-wrap tables at authoring time.** Ops can span a table boundary (a gap stretching from a body paragraph into a cell), so the wrapper cannot be rebuilt reliably, and a cell would still be paired with body paragraphs.
- **Hash a table by its flattened text only.** A merged cell holding the same words as the two cells it replaced would compare equal and the merge would never be reported.
- **Let ADR-0026 demote a table pair to whole delete + insert.** It would strike and re-insert an entire schedule for one changed amount; the per-cell demotion already handles a rewritten cell.
- **Per-cell change-list items now.** Useful, but it touches the summary model, the UI and the exports; the redline shows every cell change either way (categorisation, not suppression). Deferred.
- **Diff `tblPr`/`tcPr` formatting (widths, borders, shading).** The same class as direct run formatting, which ADR-0042 leaves Open for want of a noise answer. Recorded as a limit.

## Consequences

- New: `preprocess/table-units.ts`, `diff/word-table.ts` (stage 2.8 in code, 1b in spec §5), `diff/table-summary.ts`, `output/author-word-table.ts`, `analyseRecords` in `pipeline.ts`, `ParaRecord.wordTable`, `DiffOp.tableDiff`.
- Shared modules changed for everyone, each change confined to an input that was previously mishandled: `flattenTrackedChanges` resolves tracked rows and cells and drops a table every row of which the policy removed; `wrapParaRuns`, `markParagraphMark` and `splitParagraph` open a self-closed `<w:p/>` before revising it (it used to be wrapped whole — `<w:ins><w:p/></w:ins>`, which `CT_RunTrackChange` does not allow — or left unmarked).
- Shipped behaviour changes for every Word-vs-Word redline that contains a table: the table survives. Readability measures on real pairs that contain tables will move and must be re-measured, not assumed.
- **Accepted limits (v1):** table and cell formatting changes are not reported; a merge/split/layout change — including a vertical merge changed, and a row merged across the table when a column is added through it — reads as a whole row deleted and re-inserted; a moved table reads as delete + insert; one change item per table, the report printing the whole table both sides; tables in headers, footers and footnotes stay flat; RenderCore does not decorate row marks (`trPr` `ins`/`del`) or `cellIns`/`cellDel`, so a deleted row previews as an ordinary row with struck text; how Word desktop accepts and rejects `cellIns`/`cellDel` is unverified until the live-host smoke test.

## Measured (2026-10-18)

- **Seeded sweep** (`word-tables.sweep.test.ts`, 200 cases in the suite). Each case is a random table built from multi-paragraph cells, a nested table, a footnote reference in a cell, a vertical merge, and a row merged across the table, then given 1–3 random edits of 10 kinds. Every case must reconstruct both ways, mint unique ids, report, and revise in place. Run once at 2,000, it found two defects that the hand-written cases had not:
  1. **A blank cell paragraph written `<w:p/>`** took a replace's runs outside it and lost a footnote reference on reject. The fix opened the self-closed form in all three paragraph helpers, which also closed an open body-level follow-up (a blank paragraph inserted whole was wrapped `<w:ins><w:p/></w:ins>`).
  2. **A column added beside a row deleted and a row inserted** paired every row by position.
- **Guard ablation:** each of 14 guards, reverted alone, turns a named test red. Four were first green under ablation and got tests of their own. One of those exposed that the self-check could not see merges at all.
- **Real pair with tables** (local, never committed; shape only).
  - **Before task 10:** 2 tables → 3. One 5×5 table was revised in place with 16 cells edited. A 4-column table gained a column and had its merged bottom row split; it was struck and re-inserted row by row, 0 of 7 rows paired, because the column map required every row unmerged and the sweep's column edits only ran on unmerged grids.
  - **After:** 4 of 7 rows take the new column cell by cell. The other 3 are a vertical merge added (×2) and the split merged row.
  - **Throughout:** zero fallbacks, and both directions reconstruct table by table.
  - **Found measuring:** the flattener left an empty `<w:tbl>` shell where every row was rejected. Fixed, and the signature now counts tables.
