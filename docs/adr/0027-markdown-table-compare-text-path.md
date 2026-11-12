# ADR-0027: Markdown-table compare in the Text-vs-text path (structured row/cell redline)

- **Status:** Accepted  <!-- user-approved 2026-06-19, in the same change as the scope register + spec §2.2/§4.1 update below -->
- **Date:** 2026-06-19
- **Builds on / extends:** feature 12 (Compare — additional sources, `spec/features/10-compare-additional-sources.md`), **ADR-0018** (Text-vs-text paste compare), **ADR-0008** (categorise, never suppress), and the table-diff design already specified in **feature 08 §9.3 / §9.6**.
- **Partially supersedes:** **ADR-0018**'s blanket "Rich-text / HTML / **Markdown** paste parsing ❌" rejection — *only* its Markdown-**table** clause, and *only* for the Text-vs-text path. Broader Markdown (headings, emphasis, lists, links, blockquotes) **remains cut**.

## Context

A user pastes two versions of a Markdown **table** into Text-vs-text Compare and wants to see what changed — a renamed column, an added row, an edited cell amount. Today (ADR-0018) the paste path is a deterministic one-`<w:p>`-per-line split (`io/text/synthesise-docx.ts`): a Markdown table is destroyed at ingestion. Every row — including the `|---|---|` separator line — becomes an independent paragraph, and the diff runs line-by-line / word-by-word. An added column or a reordered row reads as intra-line **"diff salad"**, not a clean structural change. This is the exact failure the spec's structured table diff exists to prevent.

Three facts shape the decision:

1. **The structured table diff is already built, to spec, but dormant.** `diff/tables-core.ts` (`diffGrid`) implements feature 08 §9.3/§9.6 *verbatim*: align rows by row-hash LCS, then cells by cell-hash LCS, so an added column surfaces as `structural` cell inserts rather than a wholesale rewrite. `diff/tables.ts` parses `<w:tbl>` into that grid. **Neither is wired into the live pipeline** — `diffTables` is imported only by its own test.
2. **The live redline author flattens tables entirely.** `output/author.ts` rebuilds a flat `<w:body>` by concatenating per-paragraph `rawXml`; `preprocess/pararecord.ts` discards the `<w:tbl>` wrapper (it records only a boolean `isTableCell`, never the `tableCell:rXcY` partId the model and `ChangeList` UI already anticipate). So even **Word-vs-Word** tables currently flatten in the redline. Authoring a *table-preserving, cell-level* redline is genuinely new code.
3. **RenderCore already renders `<w:tbl>`.** `render-core/input/parse-body.ts` parses `<w:tbl>/<w:tblGrid>/<w:tr>/<w:tc>` into its layout model and decorates `<w:ins>/<w:del>` inside cells. Once the authored body contains a real table, the paginated preview and the vector PDF (feature 16) render it for free.

So the missing pieces are narrow: a deterministic Markdown-table **parser**, **activation** of the dormant `diffGrid`, and a structured table **redline author**. Everything else is reuse.

This is recorded separately from ADR-0018 because it reverses a specific cut it made, and it is a *structural* (not styling) capability — exact, deterministic, testable — which is precisely why it sits on the right side of ADR-0018's "non-deterministic styling rabbit hole" line.

## Decision

Add **Markdown-table-aware** comparison to the **Text-vs-text** path, tables only:

1. **Detect** GitHub-style pipe tables deterministically: a header row containing `|`, immediately followed by a delimiter row matching `^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$`, then zero or more body rows containing `|`. A literal pipe inside a cell is written `\|` (mirrors the AI-export emitter convention). Anything that does not match stays one-`<w:p>`-per-line, exactly as today. A pasted blob becomes an ordered list of **blocks** — `paragraph` lines and `table` grids.
2. **Synthesise** a real `<w:tbl>` (with `<w:tblPr>` borders, `<w:tblGrid>`, `<w:tr>`/`<w:tc>`, one `<w:p>` per cell) for each table block; paragraph blocks stay `<w:p>`. Both synthesised packages are valid OOXML that Word opens.
3. **Diff & author** structurally. A table is treated as **one alignment unit** (a record keyed by its content hash) so the surrounding paragraphs ride the *unchanged* structural → word → move → classify → cascade path. When two aligned table units differ, the new table author runs the dormant **`diffGrid`** (row-then-cell LCS) and emits a tracked-changes `<w:tbl>`:
   - **cell value change** → run-level `<w:ins>`/`<w:del>` from a cell word diff (always substantive, §7.3);
   - **added / removed column** → `structural` cell `<w:ins>`/`<w:del>`;
   - **added / removed row** → whole-row insert/delete with the OOXML row-mark (`<w:trPr><w:ins>/<w:del>>`);
   - **whole table added / removed** → all cells wrapped, row-marks set.
4. **Output:** the editable tracked-changes `.docx` remains primary (accept/reject in native Word); the RenderCore paginated **preview** and **vector PDF** render the redline table from the same authored body.

### Boundaries

- **Tables only.** Headings, `**bold**`, lists, links, blockquotes and all other Markdown stay **literal plain text** — ADR-0018's rejection of *styling* inference is unchanged. This ADR moves only the table clause.
- **Text-vs-text path only.** Word-vs-Word and the PDF modes are untouched. The new logic is reached only when a record carries a parsed table grid, which only the Markdown parser sets; the shared pipeline produces **byte-identical** Word-vs-Word output. Completing structured tables for the Word-vs-Word path is a deliberate follow-up (it changes shipped behaviour and overlaps the deferred ADR-0026 table-level threshold). The table-diff/author modules are built **reusable** so that follow-up adopts them rather than reimplementing.
- **Deterministic, no library, no ML, client-side, never-suppress.** A hand-rolled pipe parser (no `marked`/`remark`) keeps the render bundle and the zero-egress guarantee intact (ADR-0001/0003). Row/cell alignment is exact LCS — no fuzzy or semantic cell matching (ADR-0008). The redline carries **every** cell/row/column change; the separator row is consumed as table structure, never silently dropped.
- **v1 granularity.** In the change **summary / list**, a changed table counts as **one** change item (the same granularity model as a multi-paragraph replace) — *categorisation, not suppression*: the redline itself shows every cell change. Per-cell change-list enumeration (using the `tableCell:rXcY` partId the UI already supports) is a noted follow-up.

## Rationale

- **Maximum reuse / minimum new surface.** The expensive part — row/cell LCS alignment (`diffGrid`) — already exists, as does the table renderer (RenderCore) and the whole paragraph pipeline. The genuinely new code is a constrained parser plus a structured table author.
- **Right side of the ADR-0018 line.** ADR-0018 cut Markdown because *style inference* is non-deterministic. A pipe table is a **structure**, parsed by an exact, testable rule — the opposite of a styling guess.
- **Contained and reversible.** Text-path-only, additive, inert-unless-a-table-is-present; Word-vs-Word output is provably unchanged.
- **Real editable redline.** The user gets native accept/reject on cell/row/column changes — the thing lawyers actually want — not a read-only view.

## Rejected alternatives

- **Broader Markdown parsing (headings / emphasis / lists / links).** Rejected — remains the non-deterministic styling rabbit hole ADR-0018 cut. This ADR moves only the *table* clause.
- **Complete structured tables for Word-vs-Word in this change.** Rejected for now — it changes shipped Word redline behaviour, has a broad regression surface, and overlaps the deferred ADR-0026 table-level threshold. Deferred; the new modules are reusable so it is a wiring exercise later, not a rewrite.
- **A Markdown library (`marked`, `remark`, `markdown-it`).** Rejected — bundle-size pressure on the ≤5 MB lazy render path, an offline-bundling/egress audit burden, and non-determinism. A bounded hand-rolled pipe parser is deterministic and testable.
- **Treat the table as plain text (status quo).** Rejected — that *is* the "diff salad" the user is complaining about.
- **ML / fuzzy semantic cell matching.** Remains ❌ (ADR-0008). Cell alignment is exact-hash LCS only.
- **Authatically detecting a table without the delimiter row.** Rejected — heuristic table detection risks silently reshaping a user's plain text. Requiring the `---|---` delimiter row keeps detection deterministic and false-positive-free.

## Consequences

- **New code** confined to the text path: `io/text/parse-markdown-table.ts` (parser), `output/author-table.ts` (structured table redline author, reusing `diffGrid`), a text-aware record extractor, and driver/pipeline wiring.
- **Small additive, Word-inert edits** to shared compare modules: an optional table-grid field on `ParaRecord` (`model.ts`); a table branch in `output/author.ts`; skip-when-table guards in the word-diff (`pipeline.ts`), move (`diff/move.ts`) and classify stages; a pipeline core that runs from pre-built `ParaRecord[]` (mirrors the existing PDF path), shared by the Word and text paths. Each guard fires only when `ParaRecord.table` is set, which Word-vs-Word never produces.
- **the scope register feature 12** and **`spec/features/10-compare-additional-sources.md` §2.2/§4.1** are updated in the same change to record the moved boundary (Markdown tables ✅ in Text-vs-text; broader Markdown still ❌).
- Confidentiality, jurisdiction, scope discipline, and categorise-never-suppress are all preserved; this ADR + the user's approval + the scope register update **is** the scope gate.
- Acceptance lives in the project's internal test plan.
</content>
</invoke>
