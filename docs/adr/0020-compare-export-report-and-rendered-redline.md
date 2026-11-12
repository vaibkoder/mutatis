# ADR-0020: Exporting Compare output — a read-only "candidate differences report" (PDF modes) and a rendered redline snapshot (Word/Text modes)

- **Status:** Accepted — **partially superseded by ADR-0024 (2026-06-10):** the Tier-1 `window.print()` generation path for the **Word/Text rendered-redline snapshot** and the Tier-3 rule against shipping a runtime PDF generator are superseded by RenderCore (programmatic vector PDF; takes effect when feature 16's integration phase lands — until then this path stays in service). **Standing:** the artefact naming/framing (read-only snapshot complements, never replaces, the editable `.docx`; "candidate differences report" for the PDF modes), full-change-set-always export (ADR-0008), the disclosure/banner requirements, and the PDF-modes candidate-differences report path itself.  <!-- accepted by user 2026-06-09; the scope register (rows 10 & 12), spec/README build sequence, and the ADR README index updated in the SAME change. -->
- **Date:** 2026-06-09
- **Extends (does not supersede):** **ADR-0017** (PDF-vs-PDF / Word-vs-PDF — findings + side-by-side, no redline authored), **ADR-0018/0019** (Text-vs-text standalone `.docx` redline + inline preview), **ADR-0008** (categorise, never suppress), **ADR-0001/0003** (client-side-only confidentiality). It moves one boundary each ADR drew about **output**, and is recorded as its own decision per the scope register → "Changing scope".

## Context

Every Compare mode produces its result **in-pane only**, with no way to take it off-screen:

- **PDF-vs-PDF** (ADR-0017): findings + read-only side-by-side. No export.
- **Text-vs-text** (ADR-0018/0019): an editable tracked-changes `.docx` (click-to-open) + a read-only inline preview. The `.docx` *is* a downloadable artefact, but there is no flat/printable rendering.
- **Word-vs-Word** (feature 08): authors a tracked-changes redline into a new/active Word document. No standalone export.

The user wants two things:

1. **(A) Export the PDF-vs-PDF result as a downloadable artefact** — ideally a PDF, possibly also `.docx` — so the comparison can be filed, e-mailed to a colleague, or kept as a record. Today it is trapped in the task pane.
2. **(B/C) Let the Word-vs-Word and Text-vs-text modes "output PDF redlines"** — i.e. a PDF rendering of the marked-up changes, alongside their existing `.docx`.

Both re-open settled output decisions and both introduce a **new client-side generation surface**, so they go through the ADR gate. This ADR's job is to draw two lines the prior ADRs left implicit: **what an exported artefact is called and is**, and **how it is generated without breaching confidentiality**.

### The terminology line (the core of this decision)

ADR-0017's cut reads: *"Authoring a redline **into** a PDF, or writing back into the host Word doc **for the PDF modes** ❌ — PDFs have no tracked-changes target."* That cut is about **authoring an editable redline whose target is a PDF (or the host doc)** — putting accept/reject revision markup *into* a PDF, which is impossible, or writing PDF-derived changes back into the live Word document, which is fragile and out of scope.

A **standalone, read-only report that the user downloads for their records is a different artefact**: it asserts nothing editable, has no tracked-changes target, and is never written back into any document. It is the export equivalent of the side-by-side view that already ships — pixels/text on a page, not revision markup. So request (A) is **not** the cut thing; it is adjacent to it and must be named so the distinction holds:

- **PDF-vs-PDF / Word-vs-PDF export = a "candidate differences report"** — read-only, explicitly **not a redline**, carrying the same disclosures the in-pane panel does.

For requests (B/C), the same physics applies in reverse: a PDF cannot carry accept/reject revisions, so a "PDF redline" for the Word/Text modes is necessarily a **flat, read-only rendering** of the marked-up text — it **loses the editable accept/reject value that is those modes' whole point**. The editable `.docx` must remain primary and unchanged; the PDF is an *additional, read-only snapshot*, never a replacement. So:

- **Word-vs-Word / Text-vs-text export = a "rendered redline" (printable snapshot)** — read-only; the editable `.docx` (or live-doc write-back) stays the primary output.

### The confidentiality line (hard invariant)

Whatever generates the artefact must be **100% client-side, in-pane, zero network** — the same wall as the PDF.js offline hardening (ADR-0017, the zero-off-origin guard). **No cloud PDF/render service, ever.** This disqualifies every "HTML→PDF as a service" option outright and constrains the tech choice below.

## Decision

Add a **read-only export capability** to the Compare feature, uniform in framing across all modes, generated entirely client-side. Concretely:

1. **Two named artefacts, both read-only:**
   - PDF modes → a **"candidate differences report"** (never "redline").
   - Word/Text modes → a **"rendered redline" snapshot** that **complements, never replaces**, the editable `.docx` / live-doc redline.

2. **Every exported artefact carries the in-pane safety surface verbatim** (ADR-0017 / ADR-0008):
   - the **"what is NOT compared"** disclosure (images/signatures/seals/stamps; visual formatting; scanned/image-only pages) for the PDF modes;
   - the **per-run image-presence notice** ("N pages contain image content not text-compared") for the PDF modes;
   - a **read-only / not-a-tracked-changes** banner for the Word/Text rendered snapshots.

3. **No suppression in the export (ADR-0008).** The artefact contains the **full classified change set** regardless of the on-screen navigation filters (`hideFormattingOnly`, `hideWhitespacePunct`, `collapseRenumbering`). Filters are **navigation only** and **never** subtract from what is exported. (If a future "export the filtered view" option is ever wanted, it must print an explicit "filtered — N changes hidden; unfilter to export all" banner — but the **default and recommended behaviour is export-everything**.)

4. **Generation is client-side, with a tiered tech choice (recommended below):** prefer **zero new shipped dependency**; never a network call.

### Recommended technical approach (tiered, reuse-first)

The task brief lists three options; evaluated against reuse-before-build + the confidentiality wall + bundle weight:

**Tier 1 — primary recommendation: browser print-to-PDF (`window.print()` + a print stylesheet), GATED on a feasibility spike.**
- All three modes already render the differences in-pane: PDF → `SideBySideDiff` + `PdfFindings`; Text → `InlineRedline`; Word-vs-Word → `ops`/`summary` exist (currently written to the live doc, but renderable inline like the others). A **print-only stylesheet** over a dedicated print view + the disclosures gives a true "Save as PDF" with **zero new dependency, zero network, maximal reuse**.
- **Risk / open question:** `window.print()` may be blocked or misbehave inside the **Office task-pane iframe**, especially on **Word-on-web / WebKit (Mac)** — the same engine that broke PDF.js streaming. This **must be verified by a short spike before committing** (see Open questions Q1). If it works, it is the cleanest answer by a wide margin.

**Tier 2 — fallback if the spike fails: emit a `.docx` "candidate differences report" via the existing OOXML Service.**
- Reuses the **exact precedent** the brief points at — `output/writeout.ts → materialise({mode:'newDocument'})` over a synthesised package (the Text-vs-text `io/text/synthesise-docx.ts` pattern). The report is a normal document: a heading, the disclosures, then the classified findings / a flat rendering of the marked-up text. The user does **Save As → PDF** from Word for a PDF.
- **Zero new shipped dependency** (OOXML Service + PizZip already ship), deterministic, fully offline — no print-iframe uncertainty. The trade-off is it is a `.docx`, not a one-click PDF.

**Tier 3 — explicitly deferred/avoided: adopting a runtime PDF generator (`pdf-lib` / `jsPDF`) as a *shipped* dependency.**
- **Honest correction to ADR-0017:** `pdf-lib` *is* in the repo — but only as a **`devDependency`** used by `tests/fixtures/pdf/generate.ts` to synthesise test PDFs. It is **not bundled into the add-in**, so ADR-0017's "pdf-lib is not adopted [as a shipped dep]" still holds. Adopting it (or `jsPDF`) at runtime would add **bundle weight** and a **new offline-hardening surface** that must pass the same zero-off-origin guard as PDF.js. `pdf-lib` is also low-level (manual text placement — no flow/wrap/pagination), so rendering a flowing redline is real work.
- **Recommendation: do not adopt a runtime PDF generator unless both Tier 1 and Tier 2 prove insufficient.** If ever adopted, it inherits the full PDF.js offline-hardening regime (local assets, zero-network guard/test).

**Sequencing (per the one-feature-at-a-time cadence and the original ask):** build **(A) the PDF-vs-PDF candidate-differences report first** — it is the literal request and the highest-value gap (PDF results are the only mode with no downloadable artefact at all). Then add the Word/Text rendered-redline snapshots, which are lower-value (those modes already have an editable `.docx`) and carry the "flat PDF loses accept/reject" caveat.

## Rationale

- **Maximum reuse, minimum new surface.** Tier 1 adds a stylesheet + a print view over components that already exist; Tier 2 reuses the shipped OOXML write-out path. Neither adds a runtime dependency or a network call.
- **The terminology holds the line ADR-0017 drew.** "Candidate differences report" and "rendered redline snapshot" are *read-only artefacts*, not redlines authored into a PDF target and not write-backs — so this extends ADR-0017 without contradicting its cut.
- **Confidentiality preserved by construction.** `window.print()` renders the local DOM; the OOXML path is the proven local materialiser. No bytes leave the machine; no cloud render service is ever in the design.
- **ADR-0008 preserved and carried into the artefact.** The export contains every classified change and the same blind-spot disclosures the panel shows — the danger ADR-0017 guards against (false "no differences" confidence on an execution copy) cannot be reintroduced by a quieter exported PDF.
- **The editable `.docx` value is intact.** For the Word/Text modes the PDF/print export is additive; the accept/reject `.docx` remains primary, so no lawyer loses the thing they act on.

## Rejected alternatives

- **Calling the PDF-mode export a "redline".** Rejected — a PDF has no tracked-changes target; naming it a redline invites a lawyer to treat a read-only text/pixel artefact as authoritative accept/reject markup. It is a **candidate differences report**.
- **Any cloud / server-side HTML→PDF or compare-to-PDF service.** Rejected outright — violates the confidentiality invariant (ADR-0001/0003). Non-negotiable.
- **Exporting the *filtered* on-screen view by default.** Rejected — a filtered export can silently omit a substantive change (ADR-0008). The default exports the full change set; a filtered export, if ever added, must print an explicit "N changes hidden" banner.
- **A PDF "redline" for Word/Text that replaces the editable `.docx`.** Rejected — a flat PDF cannot carry accept/reject; replacing the `.docx` would destroy those modes' core value. The PDF is an additional read-only snapshot only.
- **Adopting `pdf-lib`/`jsPDF` as a shipped runtime dependency up front.** Rejected for v1 — bundle weight + a new offline-hardening surface for a capability `window.print()` or the OOXML `.docx` path likely already cover. Reconsider only if both fail (Tier 3).
- **Writing the PDF-derived report back into the host Word document.** Rejected — that *is* ADR-0017's cut; out of scope and fragile.

## Consequences

- A **small, presentational** capability lands on the existing Compare engine: a print stylesheet + a print/export view (Tier 1) or a synthesised report document (Tier 2). The diff, classifier, and `output/author.ts` are untouched.
- On acceptance, **the scope register rows 10 and 12** gain a read-only-export note, and the **ADR README index** gains this entry, **in the same change** (per "Changing scope").
- A separate piece of work covers it, sequenced PDF-report-first.
- The confidentiality guarantee extends to the export path: any generation surface is covered by the existing zero-off-origin posture; if a runtime PDF generator is ever adopted it inherits the PDF.js offline-hardening regime + guard test.

## Resolved decisions (user, 2026-06-09)

The open questions below were put to the user and resolved as follows; the implementation is bound to these:

1. **Tech path → print spike first.** Tier 1 (`window.print()` + print stylesheet) is the target, **gated on a feasibility spike** verifying it produces a PDF inside the task-pane iframe on **Word desktop *and* Word-on-web/WebKit (Mac)**. If the spike fails, fall back to **Tier 2 (`.docx` report)**. (Implies: a true one-click PDF is **not** a hard v1 requirement — the `.docx`/Save-As-PDF fallback is acceptable — so Tier 3, a shipped PDF-generator dependency, stays off the table for v1.)
2. **"PDF redline" reframing confirmed.** For Word-vs-Word & Text-vs-text the export is a **read-only rendered snapshot** that **complements, never replaces**, the editable `.docx`.
3. **Export = full change set always** (ADR-0008), independent of on-screen navigation filters. (Adopted as the default; no filtered-export option in v1.)
4. **Sequencing → PDF report first**, then the Word/Text rendered-redline snapshots.

**Accepted by the user on 2026-06-09.** In the same change: the scope register rows 10 & 12 gained the read-only-export note, the `spec/README.md` build sequence gained the compare-export entry, and the ADR README index flipped this entry to Accepted. The implementation plan is authored next (Phase 2, strong model), opening with the `window.print()` feasibility spike that decides Tier 1 vs Tier 2 before any UI is built.
