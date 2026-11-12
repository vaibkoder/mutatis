# 10 — Compare: Additional Sources (Compare tab)

> **Status:** ✅ v1 (this feature) — extends feature **08-compare** (Word-vs-Word). Governed by **ADR-0017** (PDF-vs-PDF, Word-vs-PDF) and **ADR-0018** (Text-vs-text). Canonical scope: the project scope for feature 12.
>
> Heading scheme: `#` doc title, `##` numbered top sections, `###` sub-sections.

---

## 1. Purpose & user value

Feature 08 compares two `.docx` versions. But lawyers also need to compare versions that are **not** two saved Word files:

- a counterparty's **execution copy** arrives as a **signed PDF** ("does this PDF match the draft we sent?"), or two **PDF versions** come back ("what changed between these?"); and
- a quick check of **two pasted snippets** (a clause two ways, pasted counterparty wording, two e-mailed paragraphs) where saving each as a `.docx` first is pure friction.

All three are confidentiality-sensitive — today they mean manual reconciliation or a **cloud** compare/diff tool that uploads the content off-machine. This feature does all of it **entirely client-side in the task pane**, reusing the shipped Compare engine.

The defining invariant carries over from 08 and ADR-0008 unchanged: **the classifier categorises but never suppresses.** No change is ever silently dropped; uncertainty is a **badge**, never a hidden filter.

---

## 2. Scope

### 2.1 In scope ✅ v1

| Mode | Input | Output | Authors a redline? |
|---|---|---|---|
| **Text-vs-text** (ADR-0018) | two pasted plain-text blobs (Original / Modified) | **a standalone tracked-changes `.docx`**, generated instantly, **opened only on click** | **Yes** — real `<w:ins>`/`<w:del>`, accept/reject in native Word |
| **PDF-vs-PDF** (ADR-0017) | two text-extractable `.pdf` files | findings + read-only side-by-side text diff; approximate low-confidence table cell diff; per-page image-presence notice | No (PDFs have no tracked-changes target) |
| **Word-vs-PDF** (ADR-0017, **gated**) | active Word doc vs an uploaded PDF | findings + side-by-side + alignment-confidence badge | No |

### 2.2 Cut / out ❌

- **OCR / scanned-image PDFs** ❌ (permanent) — digitally-generated, text-extractable PDFs only; scanned pages are **reported as "no extractable text — cannot compare,"** never processed (ADR-0017).
- **ML semantic-equivalence / any suppression** ❌ — carries over from ADR-0008.
- **Authoring a redline *into* a PDF, or writing back into the host Word doc for the PDF modes** ❌ — PDFs have no tracked-changes target (ADR-0017). (Text-vs-text *does* author a standalone redline — that is its whole value, ADR-0018.)
- **Rendered-page pixel-diff overlay** ❌ (dropped 2026-06-08) — distinct from the in-scope image-*presence* notice.
- **Rich-text / HTML / broader Markdown (headings, emphasis, lists, links) paste parsing** ❌ — Text-vs-text treats these as literal plain text (deterministic line-split). **Exception (ADR-0027): GitHub-style Markdown _tables_ ✅** — detected by a header + `---|---` delimiter row and compared as real tables (structured row/cell redline), not as plain-text lines.
- **Exact/guaranteed-structure PDF table diff** ❌ — PDF tables get an *approximate, low-confidence* cell diff only.

---

## 3. Dependencies & reuse

This feature adds **almost no diff logic** — it adds **input adapters and one output path** around the shipped Compare engine.

- **Reused unchanged from feature 08:** the diff core (`diff/lcs`, `diff/structural`, `diff/word`, `diff/atomic-tokens`, `diff/move`), the entire SG/UK classifier (`classify/**`), the data model (`model.ts`), the text-normalisation/hash primitives (`preprocess/normalise`, `preprocess/hash`), the `CompareConfig` resolver, and — **for Text-vs-text only** — the redline author (`output/author.ts`) and standalone-document materialiser (`output/writeout.ts → materialise({mode:'newDocument'})`).
- **New, per mode:** see §4. PDF.js (`pdfjs-dist`, Apache-2.0, bundled offline) for the PDF modes; a plain-text→`.docx` synthesiser for Text-vs-text.
- **OOXML Service** ([ARCHITECTURE](../ARCHITECTURE.md#the-ooxml-service)) — owns `Word.createDocument(base64).open()` for the Text-vs-text redline; the PDF modes do not author a package.
- **Config Layer** — the classifier reads the same SG/UK `compare.*` keys; **no US conventions** are introduced by any adapter (ADR-0002).

The source seam already exists: `compare/io/load.ts` carries a discriminated `CompareSource` union (`activeDocument | docxArrayBuffer`) with documented seams for the new kinds.

---

## 4. The three modes

### 4.1 Text-vs-text — paste → instant standalone redline (ADR-0018)

- **Input:** two text areas (Original / Modified), plain-text paste.
- **Adapt:** split each blob into **blocks** (deterministic) — paragraph lines (**one `<w:p>` per line**) plus any GitHub-style Markdown **tables** (ADR-0027), each synthesised as a real `<w:tbl>` (`<w:tblPr>` borders, `<w:tblGrid>`, `<w:tr>`/`<w:tc>`, one `<w:p>` per cell). The result is a **minimal valid `.docx` per side** (`[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `word/styles.xml` + `word/settings.xml`).
- **Reuse:** the surrounding paragraphs ride the **existing** structural/word diff, SG/UK classifier, cascade collapse and move detection verbatim; a table is one alignment unit whose changes are emitted by the dormant-now-activated **row-then-cell LCS** (`diff/tables-core.ts`, §9.3/§9.6) and a structured `<w:tbl>` redline author.
- **Output:** a **real tracked-changes `.docx`**, **generated eagerly** on Compare and held in memory, **opened only on the user's click** (`materialise({mode:'newDocument'})` / `.docx` download). Never auto-opened. The RenderCore preview/vector-PDF render the redline table from the same authored body.
- **v1 boundary:** plain text **+ Markdown tables**; no rich-text/HTML; no broader-Markdown style inference. **Tables only** for the structured path — text-path; Word-vs-Word table structuring is a separate follow-up (ADR-0027).
- **Main risk:** the synthesised package must be a **valid** OOXML doc that Word opens cleanly *and* renders revisions. Bounded + testable (OOXML validation + open-in-Word).

### 4.2 PDF-vs-PDF (ADR-0017)

- **Input:** two uploaded text-extractable `.pdf` files.
- **Adapt:** PDF.js (offline) text extraction → geometry paragraph reconstruction → PDF→`ParaRecord[]` adapter; reuse the diff + classifier; **no redline authored**.
- **Output:** findings + read-only side-by-side; approximate low-confidence table cell diff; per-page image-presence notice; scanned PDFs reported unsupported.
- Full task breakdown, reliability assessment, and acceptance (P1–P13) are tracked with the implementation.

### 4.3 Word-vs-PDF — gated (ADR-0017)

- **Input:** the active Word document (clean side, via `readDocx`→`ParaRecord`) vs an uploaded PDF (reuses 4.2's extraction).
- **Output:** findings + side-by-side + an alignment-confidence badge; **no redline**.
- **Gate:** ships **only if** the measured **balanced** alignment bar passes — anchor-match ≥ 0.90, unmatched-paragraph rate ≤ 0.05, over N ≥ 3–5 private test pairs + synthetic pairs, measured on the fuzzy-fallback pipeline. Otherwise deferred with the measured result recorded. (Phase E).

---

## 5. Shared invariants (all modes)

- **Confidentiality (ADR-0001/0003).** Everything is client-side. The Text-vs-text synthesised package and redline are in-memory (Blob/base64); `createDocument().open()` is a local Office API. The PDF modes bundle all PDF.js assets locally with a zero-off-origin-network guard. No backend, no upload, no telemetry.
- **Jurisdiction (ADR-0002).** All adapters feed the existing SG/UK `CompareConfig` + classifier; no US-isms introduced.
- **Categorise, never suppress (ADR-0008).** Text-vs-text authors a redline carrying **every** change. The PDF modes badge low-confidence/alignment/image-presence but never hide a diff; "no extractable text" is an explicit can't-compare, not a silent drop.
- **No OCR, no ML.** Permanent (ADR-0017, ADR-0008).

---

## 6. Output policy matrix

| Mode | Redline `.docx` | Findings panel | Side-by-side | Table cell diff | Image-presence | Confidence badge |
|---|---|---|---|---|---|---|
| Text-vs-text | ✅ (click-to-open) | ✅ | ✅ | n/a | n/a | n/a |
| PDF-vs-PDF | ❌ | ✅ | ✅ | approximate, low-confidence | ✅ | ✅ |
| Word-vs-PDF | ❌ | ✅ | ✅ | approximate, low-confidence | ✅ | ✅ (gated) |

---

## 7. Acceptance & build sequence

- **Acceptance.** Text-vs-text → `TT1–TT8` in the implementation plan. PDF-vs-PDF / Word-vs-PDF → `P1–P13` + the Phase-E alignment gate in the implementation plan.
- **Build order (recommended).** **Text-vs-text first** (`13-text-compare.md`) — simplest, no PDF.js, validates the standalone-redline path quickly; then **PDF-vs-PDF** (`12-pdf-compare.md` A–D); then **Word-vs-PDF** (`12` Phase E) behind its measured gate.
- **Config keys.** All three reuse the feature-08 `compare.*` keys; no new locale keys. Any new feature-local knobs (e.g. Text-vs-text large-paste cap, PDF input caps, reconstruction thresholds, alignment-bar constants) are **named constants**, not magic literals.

---

## 8. Reference architecture (not code)

As with feature 08 (§11, LibreOffice Writer comparison), any external tool is **reference architecture only** — ideas, never ported code; all reuse passes reuse-before-build, attribution, and the scope gate.
