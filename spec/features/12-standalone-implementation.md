# 12 — Standalone Implementation (web app / PWA, host-portable core)

**Status:** ✅ in scope as a **platform track** (post-v1-features). Established by **ADR-0022**. the project scope. Sequenced as **Phase 7** in `spec/README.md`.

> Deliver every host-independent Mutatis feature as a **standalone, backendless web application** (installable as a PWA; optionally wrapped as a signed Tauri binary later) that **shares one feature core with the Word add-in**, so a single feature edit ships to both products. The standalone target is an **"analyse & export" tool**: import `.docx`/`.pdf` → run the existing checkers → show findings → export a corrected `.docx` / rendered redline. It does **not** rebuild Word's editing surface. Backendless and in-browser, by construction — the confidentiality invariant is the hard constraint, not a goal.

---

## 0. How to use this doc

This is the design input to the work, not a task plan. Read it together with ADR-0022 and ARCHITECTURE.md before starting.

**Read first:** `docs/adr/0022-standalone-host-portable-implementation.md` (the decision + rejected alternatives), then `spec/ARCHITECTURE.md` (the shared modules and the Office seams this restructures), then this doc.

---

## 1. Purpose & user value

Mutatis's value is its checkers — defined-terms, cross-references, factual consistency, party-role, numbering analysis, Compare (incl. PDF). Today they are reachable **only from inside Word**. A user without Word open, or reviewing a `.docx`/`.pdf` they were sent, cannot use any of it.

The standalone target makes the checkers reachable **without Word**: open the app (a desktop-installed PWA or a web URL), drop in a document, get the same findings, export a corrected document or a redline. Same confidentiality guarantee — nothing leaves the machine.

**User value in one line:** "Run Mutatis's checks on a document without needing Word open — same client-side, zero-upload guarantee."

---

## 2. Scope

### 2.1 In scope ✅

- **A host-neutral feature core** — all feature logic + shared modules, free of `Office.*`/`Word.*` imports, depending only on the `HostPort` interface (§3).
- **A `host-web` adapter** implementing `HostPort` over browser primitives (file-picker ingest, OOXML parse, in-browser model + view for navigation, `Blob` download for write-back).
- **The `host-office` adapter** — today's Office JS code, refactored behind `HostPort` (no behaviour change to the add-in).
- **Two app targets** (`apps/web`, `apps/word-addin`) sharing `core` + `ui`, each with its own build + packaging.
- **PWA packaging** of `apps/web`: manifest + service worker → installable, own-window, offline-capable.
- **Analyse-and-export workflow** in the web target: import `.docx`/`.pdf` → checkers → Findings UI → export corrected `.docx` / rendered redline (download).
- **The host-neutral location refactor** of `Finding.location` (token, not `Word.Range`) — prerequisite, lands first.
- **The lint boundary**: `core/**` and `ui/**` may not import host symbols (extends `eslint-plugins/invariants.cjs`).

**Per-feature standalone compatibility matrix.** Three tiers, set by how much each feature depends on Word's *live editing surface* (its cursor/selection, its revision engine, its native fields) versus pure analysis + file I/O. The product's centre of gravity has shifted toward **export-centric** features (AI-export 11, the compare-export artefacts ADR-0020, templates-generate) — these are the *best* standalone fits. The features that resist porting are the ones welded to Word's editing engine (native updatable TOC field; the ADR-0023 numbering upgrade). Confirm each during planning.

**Tier ✅ — Ports fully (analysis and/or export; the standalone sweet spot):**

| Feature | Standalone behaviour |
|---|---|
| Factual Consistency (03) | Pure analysis; ports as-is. **Recommended proof-of-concept first feature.** |
| Defined Terms (01) | Pure analysis; navigation via host-neutral location into the in-app view. |
| Cross-Reference + Static-Numbering flag (02) | Prose-xref validation + Section Index + static-number flagging all host-agnostic. REF/bookmark *reads* (ADR-0014) come from the parsed OOXML. Assign-source fix applied as an export edit. |
| Party-Role Consistency (09) | Pure analysis; ports as-is. |
| Compare — Word-vs-Word (08) | Diff/classify/redline authoring all host-free; output is a downloaded redline `.docx` instead of in-place write-back. |
| Compare — Text-vs-text & PDF (10) | **Best-fit features** — Text paste needs no document at all; PDF path is pure PDF.js. Zero Word dependency. |
| Compare — export reports + rendered redline (**ADR-0020**) | Read-only "candidate differences report" (PDF modes) + "rendered redline" snapshot (Word/Text) are pure client-side generation → download. Natural standalone fit. |
| AI-Readable Export (11) | Pure serialisation; clipboard/download already host-neutral. Ports unchanged. |
| Templates (07) | Tagged-variable fill + data sets → **generate a filled `.docx` download**. A natural standalone fit. |
| Document Repair (05) | Metadata/comment/DMS-XML scrub, rsid strip, style conformance, section-break/blank-para cleanup are OOXML surgery → output a **cleaned `.docx` download**. Tracked-change accept-all becomes an OOXML-level accept on export (no WordApi-1.6 gate). |

**Tier ⚠️ — Works, but the interaction model is reframed (no Word cursor/canvas):**

| Feature | Reframing |
|---|---|
| Interactive Reference Lookup (04) | Needs the in-app document/outline view to drive selection→preview. Works as "click a cross-ref → preview its target"; the add-in's "Word cursor never moves" value prop is moot here. |
| Numbering — apply / promote-demote (06) | *Detection* ports cleanly. *Applying* a scheme is doable as an OOXML export transform, but interactive promote/demote wants a live editor; standalone does apply-then-export, not live restructuring. |
| Clause Library (07) | Storage + browse + fuzzy search port (localStorage/JSON/Fuse). "Insert at cursor" becomes copy-to-clipboard or insert-into-exported-doc. |
| TOC / Index of Schedules (07) | Standalone can **compute and write a *static* TOC** into the exported `.docx`, but **not** a live updatable Word field — see Tier ❌. |

**Tier ❌ — Does not port (welded to Word's live editing surface; stays add-in-only):**

| Feature | Why |
|---|---|
| Native updatable **TOC field** (07) | `Range.insertField(FieldType.toc)` + F9 refresh is a Word engine feature. Standalone can emit a *static* table of contents into the exported doc, but not a field that re-paginates/updates inside Word. |
| **Numbering upgrade** (ADR-0023) | Built entirely on **Word's native tracked-changes engine** (`changeTrackingMode = trackAll`, human accept/reject of revisions *in Word*) + **live Word selection** for the human-driven re-apply. No host equivalent. Standalone can still *detect* static numbers and even strip prefixes on export, but the "assisted tracked-changes removal, reviewed in Word, then human-selected re-apply" loop has no standalone form. |
| Live in-place editing / Word-driven navigation generally | The standalone target is analyse-and-export by design (ADR-0022), not a Word replacement. |

### 2.2 Deferred / cut 🔜❌

- ❌ **Any backend / server-side processing of document content.** The hard red line (ADR-0001/0003/0022). The web target is static and in-browser only.
- ❌ **Rebuilding Word's in-canvas editing surface** — live track-changes authoring/acceptance (incl. the ADR-0023 numbering-upgrade loop), native updatable TOC fields, in-place redline insertion. The standalone target is analyse-and-export; these stay Word-add-in-only. (ADR-0022.)
- 🔜 **Signed desktop binary (Tauri).** Optional later layer for formal distribution; reuses the same bundle. Not part of the first standalone milestone.
- 🔜 **In-browser *editing* of the document** beyond apply-and-export (i.e. a true web editor). Out for this track; revisit only with a new ADR.
- 🔜 **Numbering *apply* / Repair *write-back* / Templates / Compare as live in-app mutation** — the web target **exports a modified/generated `.docx`** rather than mutating a live host document. Live in-place mutation stays the add-in's job.

### 2.3 Invariant check (must hold throughout)

- **Confidentiality** — web target makes **no network call carrying document content**; only its own static assets + opt-in storage adapters (user-authored library data only), exactly as the add-in. Reinforced by offline-capable PWA. (ADR-0001/0003.)
- **Jurisdiction** — Config layer ports unchanged; SG/UK only. (ADR-0002.)
- **Scope** — build only what this doc + the project scope mark in. Anything beyond → STOP, propose, ask. (ADR-0005.)
- **Compare never suppresses** — classifier ports unchanged; categorise, never hide. (ADR-0008.)

---

## 3. Architecture — the `HostPort` seam

The core depends on **one interface**. Each app supplies an implementation. This is the entire portability story.

```ts
// host-neutral location — replaces Word.Range in Finding.location
interface DocLocation {
  paraId: string;
  charStart?: number;
  charEnd?: number;
  label: string;            // human-readable, e.g. "Clause 4.2"
}

interface HostPort {
  // ingest + model
  loadDocument(): Promise<DocModel>;          // → existing ParaNode[] model
  onDocumentChanged(cb: () => void): Disposable;

  // navigation (replaces range.select())
  reveal(loc: DocLocation): Promise<void>;    // office: range.select(); web: scroll/highlight in own view

  // OOXML read-modify-write (already host-agnostic in shared/ooxml)
  readDocx(): Promise<DocxBundle>;
  writeBack(bundle: DocxBundle): Promise<void>; // office: insertFileFromBase64; web: Blob download

  // storage (mirrors existing Store/StorageProvider)
  storage: StorageProvider;
}
```

**The four concrete changes** (in dependency order — this is the spine of the plan):

1. **Host-neutral location.** Change `Finding.location.getRange(ctx): Word.Range` → a `DocLocation` token + a `HostPort.reveal(loc)` resolver. Update the Findings model and every consumer (the click-to-navigate in `taskpane/shell`). Lands first; unblocks everything.
2. **`HostPort` interface + extract `host-office`.** Define the interface; move today's Office JS code (`shared/office`, `shared/doc-model` loader, the `getFileAsync`/`insertFileFromBase64` ends of `shared/ooxml`) behind it. The add-in keeps working unchanged.
3. **`host-web` adapter.** File-picker ingest; **OOXML→`ParaNode[]` loader that computes rendered list numbers from `numbering.xml`** (the one genuinely-hard component — no Word renderer to supply `listItem.listString`); in-browser view for `reveal`; `Blob` download for `writeBack`.
4. **Workspace + `apps/web` + PWA.** Split into `packages/{core,ui,host-office,host-web}` + `apps/{word-addin,web}` (or the env-switched single-Vite fallback, ADR-0022); add the PWA manifest + service worker; wire CI to build both.

**Suggested layout** (ADR-0022; confirm in planning):

```
packages/
  core/         # all feature logic + HostPort interface + host-neutral models
  ui/           # Preact components; take a HostPort, host-neutral
  host-office/  # HostPort via Office JS (today's code)
  host-web/     # HostPort via file-picker + OOXML parse + in-browser view + download
apps/
  word-addin/   # wires core+ui+host-office; ships Office manifest
  web/          # wires core+ui+host-web; ships PWA manifest + service worker
```

**Lint boundary:** extend `eslint-plugins/invariants.cjs` so `core/**` and `ui/**` cannot import `Office.*`/`Word.*`; only `host-*` may. Joins the existing no-`fetch`-of-document-content rule.

---

## 4. The hard parts (call these out in the plan)

1. **List-number rendering in `host-web`.** Word *renders* `"4.2.1"`/`"(b)"`; raw OOXML only has `numPr` references into `numbering.xml`. The web `doc-model` adapter must compute the rendered string itself (abstractNum/level/format/start/lvlOverride). Bounded — Repair/Numbering already parse `numbering.xml` — but the single largest new component. The Section Index and cross-ref/numbering features depend on it.
2. **Round-trip fidelity with no host.** No Word managing the package: the web write-back owns 100% of fidelity. Apply the OOXML Service's existing caveats (ADR-0003) and validate exported `.docx` opens cleanly.
3. **Navigation without a live Range.** `reveal(loc)` in the web target needs an in-browser rendering (or at least a located, scrollable text view) to highlight. Decide the minimum viable view during planning — a structured text/outline view may suffice for analyse-and-export; a faithful document render is more work.
4. **Feature write paths.** Numbering-apply and Repair-write must be reframed as "produce a modified `.docx` for download" in the web target, not live mutation. Confirm each feature's web behaviour explicitly.

---

## 5. Suggested phasing (to be refined into tasks)

1. **Seam prep** — host-neutral `Finding.location` + `HostPort` interface; extract `host-office`; add the lint rule. Add-in still green. *(no new product behaviour yet)*
2. **Proof of concept** — `host-web` + `apps/web` running **Factual Consistency only** end-to-end (import `.docx` → findings → export). Validates the whole pipeline on the simplest, fully host-free feature.
3. **DocModel completeness** — the `numbering.xml` list-number renderer; light up Section-Index-dependent features (cross-ref, defined-terms, numbering *detect*).
4. **Compare + export** — Word/Text/PDF compare in the web target; rendered-redline / corrected-`.docx` download.
5. **PWA packaging** — manifest + service worker + offline asset caching; installable.
6. **(optional, later)** Tauri binary for signed distribution.

Each phase ends green (lint/typecheck/test) and reviewed against this spec.

---

## 6. Acceptance (definition of done for the track)

- `apps/web` runs as a static, **offline-capable installable PWA**; makes **zero network calls carrying document content** (assert with a Playwright zero-network spec, mirroring the add-in's).
- A user can **import a `.docx`, get the same findings the add-in produces, and export** a corrected `.docx` / redline — for every feature listed in §2.1.
- The **add-in is unchanged in behaviour** — same features, same tests green — now running on `host-office` behind `HostPort`.
- The **no-host-import lint** passes; `core/**` and `ui/**` are provably host-free.
- A single feature edit, made in `core`, is observable in **both** `apps/web` and `apps/word-addin` after rebuild (the propagation guarantee).
- All four invariants hold; review against this spec passes.
