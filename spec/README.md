# Mutatis — Open-Source Legal Word Add-in (SG / UK)

> **Product name:** **Mutatis** (ADR-0012).
> **Status:** Scaffold and shared infrastructure in progress; feature docs ready for implementation.
> **Audience:** Contributors implementing the add-in.

---

## Vision

Mutatis is an **open-source Microsoft Word task-pane add-in** that assists legal drafting and document review for lawyers in **Singapore** and the **United Kingdom (England & Wales)** — explicitly **not** the US market. It is modelled loosely on commercial proofing/clause/compare add-ins, but built on a **fully client-side, zero-backend architecture**: the entire application runs inside the Word task pane's embedded WebView and **never transmits document content to any server**.

This is the product's core positioning — **confidentiality-first**. Because documents never leave the user's machine, the add-in is compatible with the strict confidentiality and data-protection obligations that bind its target users (the UK **SRA Code of Conduct** and the Singapore **PDPA** / banking-secrecy expectations), and it requires **no Docker, no server, no account, and no IT setup** to run. The only permitted network calls are (a) loading the add-in's own static assets, and (b) explicit, opt-in storage adapters (GitHub / SharePoint) that carry only user-authored library data — never the working document — and are off by default.

## Target users & positioning

- **Primary users:** transactional and disputes lawyers, paralegals, and legal PAs in SG and E&W firms and in-house teams, drafting commercial agreements, deeds, schedules, and court documents.
- **Positioning:** the confidentiality-safe, jurisdiction-correct alternative to cloud "document inspector" / comparison SaaS. Every convention defaults to SG/UK practice (Clause not Section, day-first dates, £/S$ not bare-$, en-GB spelling). US conventions are recognised but never defaulted to — see [CONVENTIONS.md](CONVENTIONS.md).

---

## Status legend

| Tag | Meaning |
|---|---|
| ✅ | **v1 in-scope** — specified and slated for the first release. |
| 🔜 | **v2 deferred** — acknowledged, designed-around, not built in v1. |
| ❌ | **Cut** — deliberately not built (often a platform impossibility or a malpractice-risk decision), not a temporary gap. |

**Rationale convention:** every 🔜 and ❌ item carries a one-line reason in the feature doc and in the consolidated lists below. A cut is never silent; a deferral always names what replaces it in v1.

---

## Feature list

| Feature | Tab | Status | One-line description | Doc |
|---|---|---|---|---|
| **Defined Terms Checker** | Proof | ✅ | Used-but-not-defined, defined-but-not-used, defined-multiple-times, capitalisation drift — tuned for SG/UK double-bold-quote definitions. | [features/01-defined-terms.md](features/01-defined-terms.md) |
| **Cross-Reference & Static Numbering** | Proof | ✅ | Validates prose cross-references (OK / Not Found / Ambiguous), badges manually-typed refs, and flags hand-typed clause numbers. Also defines the shared **Section Index**. | [features/02-cross-reference-and-static-numbering.md](features/02-cross-reference-and-static-numbering.md) |
| **Factual Consistency** | Proof | ✅ | Weekend execution dates, number↔word/currency mismatches, unfilled placeholders. | [features/03-factual-consistency.md](features/03-factual-consistency.md) |
| **Interactive Reference Lookup** | Review | ✅ | Selection-triggered sidebar preview of the target clause/definition — cursor never moves. | [features/04-interactive-lookup.md](features/04-interactive-lookup.md) |
| **Document Repair** | Repair | ✅ | Metadata/comment/DMS-XML scrubbing, rsid strip, style/profile conformance, section-break and blank-paragraph cleanup, accept-all (gated). | [features/05-repair.md](features/05-repair.md) |
| **Numbering** | Numbering | ✅ | Apply/save/share legal numbering schemes; promote/demote; post-change cross-ref highlight review. **+ Numbering upgrade (ADR-0023):** assisted static-number removal via tracked changes + human-driven re-apply. | [features/06-numbering.md](features/06-numbering.md) |
| **TOC / Clause Library / Templates** | Content | ✅ | Native TOC + Index of Schedules; vetted clause library with jurisdiction warnings; tagged-variable templates + shared data sets. | [features/07-toc-precedent-templates.md](features/07-toc-precedent-templates.md) |
| **Compare** | Compare | ✅ | Word-vs-Word OOXML redline (`<w:ins>`/`<w:del>`) with rule-based classification and move detection. | [features/08-compare.md](features/08-compare.md) |
| **Compare — additional sources** | Compare | ✅ | **Text-vs-text** paste→instant standalone redline; **PDF-vs-PDF**; **Word-vs-PDF** (gated). All client-side. | [features/10-compare-additional-sources.md](features/10-compare-additional-sources.md) |
| **AI-Readable Export ("Copy for AI")** | All tabs | ✅ | Per-panel "Copy for AI" buttons + one-click "AI Briefing Pack" (Proof tab). Exports JSON or Markdown: resolved glossary, clause tree, before/after change set, deal facts, cast list. Client-side only; zero network. | [features/11-ai-export.md](features/11-ai-export.md) |
| **Standalone Implementation** (platform track) | n/a (separate app) | ✅ | Host-portable feature core shared by the Word add-in and a **backendless web app / installable PWA**. Standalone target is "analyse & export": import `.docx`/`.pdf` → checkers → findings → export corrected doc / redline. Same zero-upload guarantee. Sequenced **post-v1-features** (Phase 7). | [features/12-standalone-implementation.md](features/12-standalone-implementation.md) |
| **RenderCore** (rendering track) | Compare | ✅ | Own client-side paginated redline rendering engine: worker-resident layout core (harfbuzzjs, Word line-break semantics, footnote reservation, basic tables) + OFL metric-clone fonts → paginated preview + **programmatic vector PDF** from one layout result. Supersedes the `window.print()` rendered-redline path when its integration phase lands. Spike-gated; two long-horizon phases (engine, then integration). Build-sequence Phase 8. | [features/16-render-core.md](features/16-render-core.md) |
| **Compare at scale** (1:N & M:N) | Compare | ✅ | Batch redlining in the **standalone/web target only**: one template against *N* variants, and two bundles paired document-to-document. Output is ***N* separate redlines** + a compact index — **no deviation matrix** (deferred to a separate product). M:N pairing is **proposed and user-confirmable, never silent**. Word add-in untouched. Build-sequence Phase 9. | [features/17-compare-batch.md](features/17-compare-batch.md) |

### What was cut and why (❌)

Consolidated from all drafts:

- **Word's native Compare** — not exposed in Office JS; we author our own OOXML redline instead (Compare).
- **ML / semantic-equivalence "this change doesn't matter, hide it"** — categorically cut as a malpractice risk. The Compare classifier *categorises* but **never suppresses**.
- **One-shot static→dynamic numbering conversion** — OOXML `abstractNum`/`num` injection is too corruption-prone; **stays cut**. Static numbering is detect-and-flagged (Cross-Reference Feature B), and the **Numbering upgrade** (ADR-0023, supersedes ADR-0006) adds an *assisted, tracked-changes, human-driven* removal + re-apply path — typed-API only, no injection.
- **Automatic cross-reference rewriting after renumbering** — error-prone; replaced by warn-and-highlight (Numbering) / Assign-source text replacement (Cross-Reference).
- **Live REF-field insertion in Assign-source repair** — v1 replaces with static text; v2 seam left.
- **Cross-document / cross-file resolution** (reading sibling `.docx`/external agreements) — Office JS cannot read other files. Detected and labelled "external", never resolved (Cross-Reference, Lookup, Defined Terms). *(This is about **resolving a cross-reference into another agreement from inside the add-in**. It does **not** cut the standalone target's batch Compare (feature 17 / ADR-0037), where the user explicitly uploads the files and no Office JS is involved.)*
- **Split-screen / docked target view** (Lookup) — no Office JS API to split the document viewport; sidebar-first is deliberate.
- **Real-time per-keystroke feedback** (Lookup) — `DocumentSelectionChanged` must be debounced; replaced by selection-triggered "feedback on demand".
- **Text-box detection/removal** and **per-author tracked-change accept** (Repair) — out of scope / not exposed by the API (bulk accept only).
- **PDF-to-Word structural repair** (Repair) — not part of the product.
- **Table of Authorities / Bluebook / OSCOLA / citation indexing** (Content) — US/citation-tooling territory; out of scope.
- **Drag-and-drop insertion** (Content) — task panes cannot initiate HTML5 drag-and-drop; click-to-insert only.
- **True repeating-section content controls / OOXML cloning** (Templates) — not exposed by Office JS; replaced by fixed slots (default 4, max 6).
- **KMS integrations (iManage / NetDocuments / OpenText)** in core — enterprise plugin adapter only.
- **OneDrive/SharePoint *scheme* sync** (Numbering) — replaced by JSON export/import.
- **Table-cell & footnote numbering**, **style-linked list modification** (Numbering) — unreliable / template-breaking; skipped or detect-and-warn.
- **US conventions everywhere** — `$`-as-USD default, month-first dates, "Section §" as the only clause noun, "Inc./LLC/EIN/State-of-incorporation", US-litigation outline as default. Recognised where sensible, never defaulted.
- **Pyodide / heavy WASM** in v1 — available only if a future feature is provably CPU-bound.

### Deferred to v2 (🔜)

- **Public-holiday calendars** for the weekend check: SG gazetted public holidays + E&W bank holidays (Scotland/NI differ and are out of scope), plus a **defined-"Business Day" cross-check** (Factual Consistency).
- ~~**Word-vs-PDF compare** via PDF.js~~ — **moved into ✅ scope** (feature 12, ADR-0017; spec `features/10-compare-additional-sources.md`), gated. **Fuzzy moved-and-edited paragraph** detection remains v2 for the Word-vs-Word path (the PDF path has a scoped fuzzy fallback).
- **Live REF-field insertion + bookmark creation** for Assign-source repair (Cross-Reference).
- **Header/footer/text-box coverage** for cross-reference and numbering scanning (Cross-Reference).
- **Endnote scanning** (Defined Terms scans footnotes only in v1).
- **Multi-column (3+) definition tables** (Defined Terms handles the 2-content-column case).
- **Page-number field auto-insertion** for court page-number positioning (Repair flags, does not insert).

---

## Build sequence (recommended)

Build the **shared infrastructure first**, then features in dependency order so each unlocks the next. (Synthesised from [ARCHITECTURE.md](ARCHITECTURE.md) §Build and per-feature milestones.)

**Phase 0 — Scaffold.** Yeoman/Office scaffold → migrate to Vite; manifest declaring **WordApi 1.5** floor; HTTPS dev cert; Preact + signals; `Office.onReady` gate; `shared/office` Capabilities module + `context.sync` batching helpers.

**Phase 1 — Core shared infra (unlocks everything).**
1. **Config/Localisation layer** (`shared/config`) — every checker reads it first.
2. **Document Model / Paragraph Index** (`shared/doc-model`) — single-pass loader + event invalidation.
3. **Findings model** (`shared/findings`) + **Task Pane UI Shell** (`taskpane/shell`) — renderable spine.
4. **Storage abstraction** (`shared/storage`) — local + document settings + JSON import/export; adapters stubbed.

**Phase 2 — Structure + first features.**
5. **Section Index builder** (`shared/section-index`) — authored by Cross-Reference; depends on doc-model + config.
6. **Factual Consistency** (dates/currency/placeholder) — validates the Findings→UI pipeline end to end (doc-model + findings + config only).
7. **Defined Terms Checker** (doc-model + findings + config).
8. **Cross-Reference & Interactive Lookup** (depend on Section Index).
9. **Numbering** (Section Index + config presets; consumes the static-numbering flag).

**Phase 3 — OOXML-dependent features.**
10. **OOXML Service** (`shared/ooxml`).
11. **Repair** (primary OOXML Service consumer).

**Phase 4 — Content/library features.**
12. **TOC, Clause Library, Templates** (depend on storage + adapters; wire opt-in GitHub/SharePoint here).

**Phase 5 — v2.**
13. **Compare** (diff-match-patch + OOXML base64 round-trips, worker-offloaded), then **PDF import/compare** (PDF.js).

**Phase 5b — Compare export (read-only; builds on Compare).**
13b. **Compare export** (**ADR-0020**) — a read-only download for every Compare mode: a **"candidate differences report"** (PDF-vs-PDF / Word-vs-PDF) and a **"rendered redline" snapshot** (Word-vs-Word / Text-vs-text) that complements, never replaces, the editable `.docx`. 100% client-side: **print-to-PDF (`window.print()` + print stylesheet), spike-gated**, with a **`.docx` report fallback** via the OOXML write-out path; no shipped PDF-generator dependency in v1. Carries the in-pane blind-spot disclosures and the **full classified change set regardless of on-screen filters** (ADR-0008). Sequenced **PDF-report-first**. *(The Word/Text snapshot generation path and the no-shipped-PDF-generator rule are superseded by RenderCore — takes effect when feature 16's integration phase lands; until then this path stays in service. The PDF-modes candidate-differences report path is unaffected. ADR-0024.)*

**Phase 5c — Numbering upgrade (assisted static-number removal; builds on feature 06).**
13c. **Numbering upgrade** (`spec/features/06-numbering.md` §11.1, **ADR-0023**, supersedes ADR-0006) — a two-step, human-gated power-up on the Numbering tab: (Step 1) delete hand-typed clause-number prefixes + whitespace **as tracked changes** over a batch selection (batch-guarded, whitespace auto-rule + single global override, change-tracking-availability gated); (Step 2) the user selects which freshly-cleared paragraphs to re-number and runs the existing M2 apply engine + §5.3 cross-ref review. **One-shot OOXML conversion stays cut** — typed-API only, natively reversible, no `abstractNum` injection. Milestones M9–M12.

**Phase 6 — AI export (no new scanning; build after source features).**
14. **AI-Readable Export** (`src/features/ai-export/`) — a pure serialisation layer. Depends on features 01, 02, 03, 08/10, 09 being in memory. Adds "Copy for AI" buttons to each panel and a "Copy Briefing Pack" composite to the Proof tab header. See `spec/features/11-ai-export.md` and ADR-0021.

**Phase 7 — Standalone implementation (platform track; post-v1-features).**
15. **Standalone Implementation** (`spec/features/12-standalone-implementation.md`, **ADR-0022**) — restructure to a **host-portable core** behind a `HostPort` seam, then ship a **backendless web app / installable PWA** that shares one feature core with the add-in. Order within the track: (a) host-neutral `Finding.location` + `HostPort` + extract `host-office` (add-in unchanged); (b) `host-web` + `apps/web` running Factual Consistency end-to-end as proof; (c) the `numbering.xml` list-number renderer → light up Section-Index features; (d) Compare + export; (e) PWA packaging; (f) optional Tauri binary. This is a *new product surface*, not a v1 feature — it does **not** block or alter the v1 add-in; do v1 features first.

**Phase 8 — RenderCore (rendering track; post-Compare/standalone surfaces).**
16. **RenderCore** (`spec/features/16-render-core.md`, **ADR-0024**, partially supersedes ADR-0020) — the own client-side paginated redline rendering engine: worker-resident layout core (harfbuzzjs shaping, Word greedy+shrink line breaking, usWin line height, widow/orphan/keepNext, footnote bottom-of-page reservation via fixpoint, basic tables) + vendored OFL metric-clone fonts + `.odttf` extraction, painting **one shared layout result** to a paginated preview **and** a programmatic vector PDF (`@pdfme/pdf-lib` + fontkit subsetting). Host-free; consumes `authoredBodyXml` + the existing numbering resolver; replaces the HTML-flow + `window.print()` renderer for the Word/Text rendered-redline snapshot and the inline preview when its integration phase (roadmap Phase 2) lands (PDF-modes report unaffected). **Spike-gated** (go/no-go at the top of Phase 1; fallback ladder Typst.ts → OnlyOffice stack). Built in **two phases** — Phase 1 = spikes + the whole engine, Phase 2 = integration + hardening (work consolidated, not reduced) — see the roadmap in the feature spec §6.

**Phase 9 — Compare at scale (batch track; standalone-target only; post-Compare/standalone surfaces).**
17. **Compare at scale** (`spec/features/17-compare-batch.md`, **ADR-0037**) — **one-to-many (1:N)** then **many-to-many (M:N)** batch redlining, built entirely in the **standalone/web target** (`src/web/**`, `src/host-web/**`); the **Word add-in is untouched**, because ADR-0003's two-documents-in-memory `getFileAsync` ceiling makes an *N*-way queue in the task pane structurally wrong. Preprocess the template **once** and reuse it across all *N* (`runComparePipeline` already takes pre-built `ParaRecord[]`); run pairs in a **bounded worker pool** (the repo's first — copy the RenderCore transport shape); **stream results** and never retain *N* `PipelineResult`s. Output is ***N* separate redlines*** + a flat, filter-independent batch index — the **aggregate "deviation matrix" is cut here** and deferred to a separate desktop product. M:N pairing uses content Jaccard over the existing paragraph hashes ▸ version-token-stripped filename similarity ▸ defined-term/party-cast overlap **as a tie-breaker only**, assigned by Hungarian with no-match padding, and is **always proposed in a confirmable grid before anything runs** — a mis-pairing is the pairing analogue of suppression (ADR-0008). **Cross-format *pairing* is designed for; cross-format *comparison* stays out** (ADR-0017 Phase E's alignment gate was never measured). **Prerequisite:** `compare/driver.ts` must become instantiable — it is a module-level-signal singleton today.

**Dependency summary:** `config` → everything; `doc-model` → `section-index` → cross-ref / lookup / numbering; `findings` + `shell` → all checkers; `ooxml` → repair / compare; `storage` → library / templates / datasets. **Standalone track (Phase 7)** depends on all features being expressible through `HostPort`; it restructures the seams (`doc-model`, `office`, `ooxml`, `findings.location`) but adds no new checker logic. **RenderCore (Phase 8)** depends on Compare's `authoredBodyXml` output + the numbering resolver; it adds no checker or diff logic — rendering only. **Compare at scale (Phase 9)** depends on the shipped Compare pipeline + the standalone target; it adds no diff, classification or rendering logic — batching, pairing and memory policy only.

---

## How to use this spec

Read in this order before building any feature:

1. **README.md** (this file) — what exists, status, build order.
2. **[ARCHITECTURE.md](ARCHITECTURE.md)** — the client-side model, shared modules, OOXML Service, technology stack, and the authoritative **Office JS Constraints Reference**. Feature docs deep-link to its headings; do not restate them.
3. **[CONVENTIONS.md](CONVENTIONS.md)** — the SG/UK legal + localisation single source of truth (the concrete contents of the Config layer). No feature hardcodes a convention that lives here.
4. **The feature doc** you are building (`features/NN-*.md`) — feature-specific algorithm, data model, edge cases, tests.

The shared layer is **authoritative**. Where a feature doc and a shared doc disagree on a shared concern, the shared doc wins.
