# ADR-0017: PDF comparison — PDF-vs-PDF (and Word-vs-PDF) via offline PDF.js

- **Status:** Accepted  <!-- user-approved 2026-06-08, with the scope register update (feature 12) in the same change. Decisions Q1 (table cell diff) and Q2 (image-presence notice) folded in below. -->
- **Date:** 2026-06-08
- **Supersedes / amends:** the v2 deferral of **Word-vs-PDF** in the scope register and `spec/features/08-compare.md` §2.2. Extends, does not reopen, **ADR-0008** (categorise, never suppress).

## Context

Two related comparison workflows are currently outside v1 scope:

1. **PDF-vs-PDF** — compare two `.pdf` versions and show what changed. **Not listed in the scope register at all** (neither in, deferred, nor cut).
2. **Word-vs-PDF** — compare the active Word draft against an uploaded PDF. **Deferred to v2** in the scope register and `spec/features/08-compare.md` §2.2.

Both are real SG/UK legal workflows. Execution copies routinely arrive as **signed PDFs**; counsel needs to answer "does this executed PDF match the Word draft we sent?" and "what changed between these two PDF versions?" Today this means either manual reconciliation or uploading documents to a server-side compare tool — a confidentiality (and often regulatory) problem the product exists to remove.

The v1 **Compare** feature (feature 08, already shipped) already implements a source-agnostic diff pipeline. Investigation confirms ~70–75% of it is reusable for a PDF source: the diff core (`diff/lcs.ts`, `diff/atomic-tokens.ts`, `diff/word.ts`, `diff/structural.ts`, `diff/move.ts`), the entire rule-based classifier (`classify/**`), the data model (`model.ts`), the text-normalisation/hash primitives (`preprocess/normalise.ts`, `preprocess/hash.ts`), and the `CompareConfig` resolver all operate on plain text / `ParaRecord[]` and carry **no Word-specific dependency**. A `PdfSource` discriminator seam was already left in `compare/io/load.ts` for exactly this.

The hard constraints set:
- **Confidentiality is the #1 risk.** PDF libraries frequently fetch a worker, fonts, or cmaps from a CDN and some carry telemetry. Any library that cannot run 100% offline with all assets bundled locally is disqualified (the project principles, ADR-0001, ADR-0003).
- **No OCR.** v1 handles **digitally-generated, text-extractable PDFs only.** Scanned / image-only PDFs are detected and reported as *"no extractable text — cannot compare,"* never silently mis-handled. (OCR is also an offline/confidentiality problem — a second reason it is out.)
- **Permissive licences only** (MIT / Apache-2.0; ISC acceptable if added to the allowlist).
- **ADR-0008 still binds:** categorise/filter, never suppress; no ML semantic-equivalence.

## Decision

Move both features into scope, **sequenced and phase-gated**, built on the existing Compare pipeline:

1. **PDF-vs-PDF — IN (build first).** Extract text from two PDFs with **PDF.js (`pdfjs-dist`, Apache-2.0)**, reconstruct paragraphs by geometry, feed the reconstructed text through the **existing** normalise → structural-diff → word-diff → classify pipeline, and surface the result as **findings + a read-only side-by-side text diff** in the task pane.

2. **Word-vs-PDF — IN, immediately after, behind a go/no-go gate.** Reuse PDF extraction from (1) for the PDF side; extract the active Word document's paragraphs via the existing Doc Model / `readDocx` path for the (clean) Word side; run the **same** pipeline. Ship only if the Phase-E alignment-quality acceptance bar (below) is met on real text-PDF pairs; otherwise it stays deferred with the measured reason recorded. The Word side is structurally clean, so the only new risk over (1) is asymmetry, and exact-hash anchor matching absorbs it wherever text is verbatim.

**Output surfaces (both):**
- **Findings in the shared panel** + a **read-only side-by-side text diff view** (reuses `ChangeSummary`, `ChangeList`, `WordDiff` rendering).
- **Approximate table-region cell diff — IN (Q1, user-approved 2026-06-08).** Rather than collapsing PDF tables to prose, detect x-column bands → reconstruct a coarse row/cell grid → reuse the cell-LCS logic of `diff/tables.ts` (extracted into a structure-agnostic core), **always low-confidence**; an un-griddable region falls back to flagged prose. This never asserts an *exact* cell diff (consistent with the Reliability assessment below) — it converts the worst noise case (numeric schedules → prose churn) into an aligned, badged cell diff.
- **Per-page image-presence notice — IN (Q2, user-approved 2026-06-08).** A **presence flag, not a pixel diff**: detect pages containing embedded image content the text diff cannot see (signatures, seals, stamps) via PDF.js page resources, **dedupe repeated letterhead/logos**, and surface a per-run "N pages contain image content not text-compared" disclosure. Closes the dangerous "no differences" false-confidence on execution copies. No rendering, no pixel comparison.
- **Rendered-page visual *pixel-diff* overlay — DROPPED (2026-06-08, user decision).** A pixel overlay (PDF.js render-to-canvas + pixelmatch/resemblejs) was considered to *diff* image/stamp/signature changes. **Not pursued** — see Rejected alternatives. (Distinct from the image-*presence* notice above, which is IN.) The text-diff surfaces stand alone.
- **Explicitly NOT a redline.** A PDF has no tracked-changes format to author into, and writing a redline back into the Word doc is out of scope and fragile. The Word-specific output layer (`output/author.ts`, `output/writeout.ts`, `preprocess/flatten.ts`, `preprocess/quarantine.ts`) is **not reused**.

**Scanned/no-text handling:** on load, extraction measures per-page text coverage. A PDF (or page) with no extractable text is reported as **"no extractable text — cannot compare (appears scanned/image-only; OCR is out of scope)"**, naming the affected pages. The tool never diffs a silently-truncated subset.

**Library choice:** `pdfjs-dist` (Apache-2.0) for extraction (and page rendering only if ever needed). The optional overlay libraries (`pixelmatch` ISC / `resemblejs` MIT) are **not adopted** — the overlay was dropped (2026-06-08). **`pdf-lib` is not adopted** — no output PDF is authored, so it has no role here.

## Confidentiality verification (the disqualifying criterion)

| Library | Licence | Telemetry / phone-home | Network risk & mitigation |
|---|---|---|---|
| **pdfjs-dist** (Mozilla) | Apache-2.0 | None. `getDocument()` reads bytes you hand it; the document never leaves the WebView. | Risk is **asset URLs only** (worker, `cMapUrl`, `standardFontDataUrl`, v5+ `wasmUrl`). All are configurable. Mitigation: bundle the worker and every referenced asset directory locally via Vite; set no `*Url` to a remote origin; self-host or disable the WASM paths (an unset `wasmUrl` with `useWasm` on causes 404 *attempts* — no leakage, but must be eliminated). **Verified offline-capable.** |
| **pixelmatch** (Mapbox) | ISC | None | Operates on raw RGBA typed arrays from a canvas. No assets, no network. |
| **resemblejs** (alt.) | MIT | None | Canvas-based; no mandatory network. |

**Guardrail (new, mirrors the existing three invariant lints):** add a build-time/test assertion that PDF.js asset URLs resolve to local bundled paths and that a full parse+render+diff cycle issues **zero network requests** (DevTools network assertion in the live-verify checklist; a unit/integration guard for the config). This makes the offline property a wall, not a convention.

## Invariant analysis

- **Confidentiality (ADR-0001/0003):** preserved. The only new attack surface is PDF.js asset-fetching, bounded by local bundling + a no-network guard (above). No backend, no upload, no telemetry.
- **Jurisdiction (ADR-0002):** orthogonal. The PDF source reuses the existing SG/UK `CompareConfig`, atomic tokens (`Clause`/`Schedule`/`paragraph`), and classifier unchanged. No US-isms introduced.
- **Scope (ADR-0005):** this ADR + user approval + the scope register update **is** the gate. PDF-vs-PDF moves from unlisted → ✅; Word-vs-PDF moves from 🔜 → ✅ (phase-gated). No cut feature is revived (OCR and ML semantic-equivalence remain ❌).
- **Categorise, never suppress (ADR-0008):** preserved and explicitly extended to PDF. The reused classifier only labels/filters. Two honest notes:
  - The **"no extractable text" stop is not suppression** — nothing is diffed at all; it is an explicit can't-compare, the opposite of a silent drop.
  - **Alignment uncertainty (esp. Word-vs-PDF)** can mis-*present* a real difference (e.g. show a moved-and-reflowed paragraph as delete+insert) but must never *hide* one. Any confidence/alignment indicator is a **badge/annotation, errs toward visibility**, and must **not** be a default-on filter that conceals low-confidence substantive diffs. This is the same "show more, not less" stance ADR-0008 already takes.

## Reliability assessment (honest)

PDFs store **positioned glyphs, not logical paragraphs.** PDF.js `getTextContent()` returns text items with geometry but **no reading-order, paragraph, column, or table semantics** — array order is not reading order. Reconstruction is therefore a **heuristic geometry-clustering pass** (cluster by y for lines, sort by x, infer paragraph breaks by vertical gap, detect multi-column x-bands).

- **Single-column legal prose** (most contract/deed bodies): reconstruction is reliable; paragraph-level hashing + anchor matching is robust to repagination because it matches on normalised, numbering-stripped body text, not page position.
- **Bounded failure modes:** headers/footers (already excluded by default — `compare.excludeHeaders/Footers`), signature/execution blocks (already forced always-substantive — §7.2), footnote/cross-ref renumbering (already cascade-collapsed). These existing rules map directly onto the known PDF noise sources.
- **Genuinely shaky:** **multi-column layouts and tables.** Pure PDF.js table reconstruction is unreliable. v1 bound: detect multi-column / tabular regions and **report them as lower-confidence / structurally-approximate rather than asserting a precise cell-level diff** — never silently mis-diff them. (PDFium via `@embedpdf/pdfium`, MIT + Apache/BSD WASM, is the permissive fallback engine to evaluate *only if* PDF.js text quality proves insufficient. MuPDF-based libraries are **AGPL — disqualified**.)

This is why Word-vs-PDF is **gated on a measured alignment-quality bar**, not assumed.

A reference-architecture review of an open-source contract-analysis project (2026-06-08) confirmed our reconstruction is **not under-built** relative to a shipped comparator — that project deliberately punts on paragraph reconstruction (one line = one paragraph). Three small tactics were borrowed into the ingest design within this confidentiality envelope (no OCR, offline, no suppression): pdfjs `hasEOL` as the line-segmentation primitive, a per-page font-size-median heading heuristic, and hard input caps with typed errors. No algorithm or boundary change resulted.

## Rejected alternatives

- **OCR for scanned PDFs.** Rejected (an explicit decision): out of scope, and an offline/confidentiality problem. Scanned PDFs are detected and reported as unsupported, never processed.
- **MuPDF / mupdf.js for higher-quality extraction.** Rejected: **AGPL-3.0** — incompatible with the permissive open-source core.
- **Authoring a PDF or Word redline output.** Rejected: PDFs have no tracked-changes target; writing a redline back into Word is fragile and out of scope. Output is findings + side-by-side, not a materialised redline. (`output/**` not reused.)
- **A default-on "hide low-confidence differences" filter.** Rejected: would suppress substantive changes — a direct ADR-0008 violation. Confidence is surfaced as a badge only.
- **Rendered-page visual (pixel) overlay.** Considered as an additive signal for signature/stamp/image changes; **dropped (2026-06-08, user decision)**. Whole-document pixel diff is dominated by rendering noise (font substitution, anti-aliasing, DPI, repagination), and even the scoped execution-page variant needs human eyeballing on real rendered pages to tune — a poor fit for the autonomous build budget and a low reliability-per-effort ratio versus the text diff, which already carries the core value. May be revisited later only as a narrowly-scoped, human-validated execution-page check if a concrete need arises.
- **Commercial/proprietary PDF SDKs** (Apryse/PDFTron, Nutrient/PSPDFKit). Rejected: non-permissive licences.
- **Building PDF-vs-PDF and Word-vs-PDF as one undifferentiated drop.** Rejected: Word-vs-PDF's reliability is unproven until PDF extraction is hardened; sequencing PDF-vs-PDF first de-risks it and gives Word-vs-PDF a real go/no-go gate.

## Consequences

- The combined feature spec (`spec/features/10-compare-additional-sources.md`) and the scope register update (feature 12) landed **with** this ADR's acceptance.
- New dependency: `pdfjs-dist` (lazy-loaded/code-split per the Architecture performance guidance). No overlay/image-diff dependency is taken (overlay dropped). PDF.js assets are vendored locally and covered by a no-network guard.
- ADR-0008's no-suppression invariant now explicitly governs the PDF surfaces.
- Word-vs-PDF ships **only** if the Phase-E alignment bar passes on real pairs; if not, it remains deferred with the measured result recorded — no silent best-effort shipping.
- If approved, the next step is to write the feature spec and executes the implementation plan drafted alongside this ADR.
