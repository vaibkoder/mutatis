# ADR-0024: RenderCore — an own client-side paginated layout/rendering engine for redlines (worker-resident HarfBuzz core, dual preview + vector-PDF output)

- **Status:** Accepted  <!-- commissioned and pre-approved by the maintainer via a planning brief; the scope register row 15, spec/features/16-render-core.md, spec/README build sequence, and the ADR index updated in the SAME change. -->
- **Date:** 2026-06-10
- **Extends:** **ADR-0001/0003** (fully client-side — RenderCore renders entirely in-browser, zero network egress at render time), **ADR-0008** (categorise, never suppress — the renderer styles every change and hides none), **ADR-0022** (host-portable core — RenderCore is host-free behind the same discipline as Compare-on-web).
- **Partially supersedes:** **ADR-0020**. Two of its decisions are reversed for the **Word-vs-Word / Text-vs-text "rendered redline" snapshot**: (a) the Tier-1 `window.print()` generation path is replaced by RenderCore's programmatic vector PDF, and (b) the Tier-3 rule "do not adopt a shipped runtime PDF generator" is lifted (`@pdfme/pdf-lib` + `fontkit` become shipped, vendored dependencies). **Everything else in ADR-0020 stands:** the artefact naming/framing (read-only "rendered redline" snapshot that *complements, never replaces*, the editable `.docx`; "candidate differences report" for the PDF modes), full-change-set-always export (ADR-0008), the disclosure/banner requirements, and the **PDF-modes candidate-differences report itself, which keeps its current `window.print()` path** — the PDF modes are out of RenderCore's scope.

## Context

Mutatis's Compare engine is complete and strong: per-paragraph diff ops with word-level sub-diffs, 8-class change classification, move linkage, renumbering-cascade handling, and `authoredBodyXml` — a merged `<w:body>` carrying native `<w:ins>/<w:del>/<w:moveFrom*>/<w:moveTo*>` revision markup, the same bytes used to build the editable `.docx`. The **diff is solved**. What is weak is **rendering** it: the current read-only output (ADR-0020 Tier 1) is a semantic HTML flow printed via `window.print()`. Its known weaknesses: numbering regex-scraped from literal text (fails on Word auto-numbered documents), run formatting (`rPr`) discarded, bold only via a heading heuristic, no real pagination, footnotes never placed, and no programmatic PDF (the browser print dialog is the only exit; no browser API for silent print-to-PDF exists — verified).

The product goal, in the user's explicit priority order:

1. **Robust and reliable** — every change visibly and unambiguously marked; nothing suppressed or silently dropped (ADR-0008).
2. **Conventional-redline quality** — a professional paginated artefact: real multi-level auto-numbering with hanging indents, bold headings, justified body text, preserved run formatting, footnotes at the bottom of the correct page with a separator, A4 pagination, and a downloadable **vector PDF with selectable text and embedded subset fonts**.
3. **Responsiveness** — no UI freeze ever (layout off the main thread); ~150-page document in low single-digit seconds; bundle delta in the low single-digit MB.
4. **Aesthetic exactness is NOT a goal** — the tracked-change theme is simple, configurable, standard.
5. **Line-for-line identity with Microsoft Word is a NON-GOAL** — investigated, found mostly achievable but meaningless to the reader. Never an acceptance criterion; the bar is *professional, internally consistent* pagination.

A deep engineering investigation (2026-06-10, five parallel research streams: Word layout internals; WASM typesetters; HTML paginators + PDF extraction; browser office engines; fonts & PDF backends) settled the technology questions. Its conclusions are summarised here because they carry the decision:

- **Word's layout is reproducible, not proprietary magic.** Word line-breaking is greedy first-fit (no Knuth-Plass); justification is inter-word space expansion plus (Word 2013+) an undocumented shrink of inter-word spaces up to ~20% — reverse-engineered by LibreOffice (tdf#119908, landed LO 24.2), after which LO achieves line/page parity with Word on justified Latin DOCX. **Kerning, ligatures and hyphenation are OFF by default in Word**, so Latin line breaking reduces to summing `hmtx` advance widths — deterministic given the font bytes. Word "single" line height = `(OS/2.usWinAscent + usWinDescent) / unitsPerEm × fontSize` (typo metrics if the USE_TYPO_METRICS flag is set) — different from CSS `line-height: normal`, so an engine must compute it. Pagination semantics (widowControl, keepNext, keepLines, pageBreakBefore, cantSplit, footnote reservation) are documented OOXML.
- **Fonts are the real fidelity bottleneck — solved by OFL metric-clones + embedded-font extraction.** Microsoft fonts cannot be redistributed; Liberation Serif/Sans are advance-width-identical to Times New Roman/Arial (→ identical line breaks), Carlito↔Calibri and Caladea↔Cambria near-identical. ≈1.5–3 MB as WOFF2 for 4 families × 4 styles. Embedded `.odttf` fonts deobfuscate trivially (XOR first 32 bytes with the reversed fontKey GUID); re-embedding a subset into an output PDF from the user's own document is the intended use of the embedding-permission bits.
- **No off-the-shelf engine fits** (verified negative result, June 2026): nothing is simultaneously light, offline, Safari-safe, footnote-correct, vector-PDF-capable, and permissively licensed. See Rejected alternatives.

## Decision

Build **RenderCore**: a custom, worker-resident, host-free layout and rendering engine that turns the already-solved comparison output into (a) a paginated on-screen preview and (b) a programmatic, downloadable vector PDF — both painted from **one shared layout result**, so they are pixel-locked to each other.

**Architecture (settled by the investigation; recorded in `spec/features/16-render-core.md` §3):**

1. **Input:** `authoredBodyXml` (native revision markup) + `styles.xml` + the existing numbering-resolver output + `footnotes.xml` (+ theme/embedded fonts when present). The existing comparison engine, per-paragraph model, and numbering resolver are the foundation — reused, not rebuilt.
2. **Shaping/measurement:** `harfbuzzjs` (~0.3–1 MB wasm, MIT) in a **Web Worker** — no DOM dependency, so the entire layout runs off the main thread (avoids the main-thread lock-up failure mode) and is deterministic across Chromium/WebKit (reads font bytes directly).
3. **Line breaking:** Word semantics — greedy fill, space-expansion justification + the ~20% shrink rule, kerning/ligatures off, the usWin line-height formula.
4. **Pagination:** widow/orphan (2 lines), keepNext/keepLines, pageBreakBefore, footnote bottom-of-page reservation via fixpoint iteration (the single hardest item), basic table row splitting (cantSplit).
5. **Fonts:** vendored OFL metric-clones (Liberation Serif/Sans, Carlito, Caladea) + `.odttf` extraction of embedded fonts; everything same-origin, offline.
6. **Output A — preview:** absolutely-positioned spans/SVG per line painted from the computed boxes; the tracked-change theme is plain CSS, simple and configurable.
7. **Output B — PDF:** `@pdfme/pdf-lib` (maintained fork; the original Hopding/pdf-lib is abandoned) + `fontkit` subset embedding; decorations drawn directly as vector operations.

**Build order — two long-horizon phases** (`spec/features/16-render-core.md` §6). **Phase 1 — the engine:** open with the de-risk spikes (A — the harfbuzz+Liberation thesis test against a real contract; B — the footnote-reservation fixpoint; C — pdf-lib subset-font emission on Chromium and WebKit) behind a GO/NO-GO gate, then — once GO — build the whole engine in the same phase (fonts → input model → layout core → paginator → footnotes → tables → decoration → PDF → preview), landing as tested library code. **Phase 2 — integration + hardening:** wire it behind the Compare/preview surfaces in both hosts, Text mode, retire the old path, and run the cross-browser/perf/QA sweep. Each phase is a single long-horizon unit of work. The spike gate still commits nothing further until it passes; NO-GO escalates the fallback ladder.

**Fallback ladder (recorded so a blocker never triggers a re-survey):** if Spike A surprises badly, or a later phase uncovers a genuine blocker — STOP, write the evidence, and escalate to **Typst.ts** (~12 MB gz, Apache-2.0; complete engine with native footnotes and in-browser vector PDF; needs an OOXML→Typst translation layer and is its own typesetting model, not Word-look). If Typst.ts also fails, the documented last resort is the **OnlyOffice sdkjs + x2t-wasm stack** (best Word facsimile; 50–100 MB, AGPL + disputed logo clause, 5/5 integration effort) — adopting it would require its own ADR and an explicit licence decision, since AGPL is currently disqualifying. Either escalation supersedes the relevant parts of this ADR with user approval; neither is taken silently.

**Non-goals (binding):**
- **No line-for-line Word parity** — never an acceptance criterion or test assertion. Residual divergence vs Word (~5–10% of line breaks by default) is a calibration budget, not a bug bar; Word is not even self-consistent across its own platforms.
- **No aesthetic exactness** — no exotic decorations (double underline etc.), no theme gallery; one simple configurable theme (e.g. coloured underline + highlight for inserts, coloured strikethrough for deletes, paired styling + anchors for moves, an inline change anchor).
- **Not an editor.** RenderCore renders read-only artefacts. ADR-0022's cut — no rebuilding of Word's in-canvas editing surface in the standalone target — is untouched.
- **PDF-vs-PDF mode is out of scope** (it produces a candidate-differences report, not a redline; its existing export path stands).

**Constraint check:**
- **Confidentiality (ADR-0001/0003):** all layout and PDF generation happen in the browser/worker; all dependencies and fonts are vendored same-origin; **zero network egress at render time**. The office-free grep discipline extends to every chunk the render path loads.
- **Licences:** permissive only — harfbuzzjs (MIT), @pdfme/pdf-lib (MIT), fontkit (MIT), Liberation/Carlito/Caladea (OFL). No AGPL.
- **Host-free (ADR-0022):** RenderCore lives in the host-free core, loadable from both the standalone web app (primary target) and the task pane; no `Office.*`/`Word.*` imports.
- **Never suppress (ADR-0008):** acceptance criterion 1 is a property test that every diff op in the change-set appears in the rendered artefact, styled by its class.
- **Bundle:** total shipped delta ≈ 2 MB code/wasm + 1.5–3 MB fonts, code-split and lazy-loaded (within the ARCHITECTURE.md heavy-dep policy; harfbuzzjs is the "provably CPU-bound" WASM case the spec reserved for).

## Rationale

- **The diff is solved; rendering is the remaining gap between Mutatis and a professional-grade artefact.** The current HTML-flow + `window.print()` output discards run formatting, mis-derives numbering, cannot paginate, cannot place footnotes, and cannot produce a programmatic PDF. Those are exactly the user's priority-2 items, and no incremental patch to a CSS flow renderer can deliver pagination physics or footnote reservation.
- **A custom core is justified only because the survey came back negative** — and it did, with evidence (see Rejected alternatives). Once HTML/print routes are excluded (no silent print-to-PDF API) and AGPL engines are excluded (licence wall), what remains is shaping + arithmetic over documented OOXML semantics: greedy breaking with known justification rules, a published line-height formula, enumerable pagination flags. The hard parts (footnotes fixpoint, tables) are bounded and spiked first.
- **Worker-resident HarfBuzz kills both the performance and the determinism problems at once:** layout never touches the DOM (no main-thread lock-up), and measurements come from font bytes, not from whichever browser text engine happens to be underneath (Chromium vs WebKit consistency).
- **One layout result, two painters** guarantees the preview and the PDF paginate identically — an acceptance criterion no print-dialog route can meet.
- **Metric-clone fonts make line breaks right even when the user lacks the original fonts**, and `.odttf` extraction covers documents that carry their own. The preview may *display* a locally-installed real Microsoft font while positions stay valid, because the clones are advance-width-identical.
- **Spikes-first ordering** caps the cost of being wrong: the thesis (Spike A), the hardest algorithm (Spike B) and the output backend (Spike C) are each validated before any production phase, and the fallback ladder is pre-recorded so failure escalates instead of re-opening the survey.

## Rejected alternatives

All surveyed in that investigation. Do **not** re-survey; if a phase uncovers a genuine blocker, use the fallback ladder above.

- **Keep/extend the existing HTML-flow + `window.print()` renderer (ADR-0020 Tier 1).** Rejected as the redline export path: no real pagination, no footnote placement, run formatting discarded, numbering scraped by regex, and the print dialog is a manual, browser-chrome-dependent step that cannot produce a programmatic PDF. It remains the bar to clear, and remains in service for the PDF-modes candidate-differences report (out of RenderCore's scope).
- **Any HTML print-dialog route generally** (print stylesheets, hidden iframes). Rejected: **no browser API exists for silent print-to-PDF and none is coming** (verified). Fails the programmatic-PDF requirement categorically.
- **LibreOffice WASM (ZetaOffice).** Rejected: ~250 MB, Safari untested — disqualified on weight alone.
- **OnlyOffice sdkjs + x2t-wasm.** Rejected as primary: best Word facsimile available, but 50–100 MB, AGPL + a disputed logo clause, and 5/5 integration effort (CryptPad/ZIZIYI pattern). Kept as the documented **last-resort fallback**, gated on its own ADR + licence decision.
- **Typst.ts.** Rejected as primary: ~12 MB gz, Apache-2.0, a complete engine with native footnotes and in-browser vector PDF — but it is its own typesetting model (not Word-look) and needs an OOXML→Typst translation layer. Kept as the **first fallback** (plan-B) on the ladder.
- **Vivliostyle.** Rejected: AGPL, and PDF only via the print dialog.
- **paged.js.** Rejected: frozen at 0.4.3, buggy footnotes, main-thread layout (re-introduces the lock-up failure mode).
- **docx-preview.** Rejected: no real pagination.
- **SuperDoc.** Rejected: AGPL, no PDF export.
- **@react-pdf/renderer.** Rejected: Knuth-Plass line breaking (≠ Word semantics), weak footnotes and tables.
- **Commercial WASM layout engines.** Rejected: proprietary licences — incompatible with the open-source, permissive-only constraint.
- **jsPDF as the PDF backend.** Rejected: no font subsetting. (`pdfkit` is the recorded acceptable alternative backend if the pdf-lib forks disappoint.)
- **Cloud/server-side rendering of any kind.** Rejected outright — violates the confidentiality invariant (ADR-0001/0003). Non-negotiable, as ever.
- **Making line-for-line Word parity the goal.** Rejected: investigated and found *mostly achievable but meaningless* — the reader gains nothing, Word itself is not self-consistent (Mac CoreText vs Windows DirectWrite, compat modes), and chasing the last ~5–10% of break parity converts a bounded project into an unbounded one. The bar is professional, internally consistent pagination.
- **Rendering from the diff ops / flow model instead of `authoredBodyXml`.** Rejected: the merged OOXML with native revision markup is the same bytes as the editable `.docx` output, keeping rendered and editable artefacts in lock-step and giving the renderer every run property the flow model discards.

## Consequences

- **Scope moves in the same change (per the scope register "Changing scope"):** the scope register gains row 15 (RenderCore, `spec/features/16-render-core.md`); the cut list gains "line-for-line Word-parity acceptance bar" and "decoration/theme gallery"; `spec/README.md` gains the feature row + build-sequence Phase 8; the ADR index records ADR-0020's partial supersession.
- **A committed roadmap of two long-horizon phases** (`spec/features/16-render-core.md` §6, ADR-0022 cadence): **Phase 1** consolidates the spikes + the whole engine (fonts, input model, layout core, paginator, footnotes, tables, decoration, PDF, preview) into one phase; **Phase 2** is integration + Text mode + hardening. The full component list is preserved — consolidated, not cut. Each phase is planned and executed as one long-horizon unit of work.
- **New shipped dependencies** (harfbuzzjs, @pdfme/pdf-lib, fontkit, OFL font binaries) enter the bundle, vendored same-origin, code-split and lazy-loaded; the zero-egress and office-free grep disciplines extend to all render-path chunks. Bundle budget: ≤ ~5 MB delta including fonts.
- **On completion of Phase 2 (integration)**, the Word/Text rendered-redline snapshot is generated by RenderCore (programmatic PDF download) and the inline-redline preview (feature 14) upgrades to the paginated preview; the old flow-renderer path for those surfaces is retired. **Until then, the ADR-0020 Tier-1 path stays in service** (the PDF-modes candidate-differences report keeps it permanently).
- **Execution model, not a multi-week estimate.** This is built as two long-horizon phases rather than a many-week programme. The footnote, table, and PDF work is the point of the engine and is not dropped to fit a schedule; consolidation reduces handoffs, not capability.
- **Risk posture:** Phase 1's spike gate caps the bet; the fallback ladder (Typst.ts → OnlyOffice stack) is pre-recorded with triggers, so a blocker escalates with evidence instead of re-opening the engine survey.
- **What gets harder:** font licensing hygiene (OFL attribution files ship with the bundle), a real calibration/QA burden across Chromium + WebKit (the project's WebKit lessons apply), and a much larger rendering test surface (property tests against the comparison engine's output, pagination fixtures, PDF assertions).
