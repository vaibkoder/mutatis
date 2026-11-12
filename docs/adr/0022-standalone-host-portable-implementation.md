# ADR-0022: Standalone host-portable implementation — a backendless web app / PWA sharing one feature core with the Word add-in, via host-adapter layering

- **Status:** Accepted
- **Date:** 2026-06-09
- **Extends (does not supersede):** **ADR-0001 / ADR-0003** (fully client-side, no backend — reaffirmed and made the hard constraint of the new target), **ADR-0002** (SG/UK jurisdiction — unchanged), **ADR-0008** (Compare categorises, never suppresses — unchanged), **ADR-0004 / ADR-0012** (tech stack — the web target reuses the same Preact + TS + WASM bundle).

## Context

Mutatis today is a Microsoft Word task-pane add-in: the entire app runs inside Word's embedded WebView and reaches the document through Office JS (`Word.run`, `getFileAsync`). The maintainer wants a **second delivery surface** — a standalone application a user can run *without Word* — that offers every feature that can stand on its own outside the Word host, and that **stays in sync with the add-in feature-for-feature** so a single feature edit lands in both products.

Two forces shape the decision:

1. **The confidentiality invariant is non-negotiable (ADR-0001/0003).** Mutatis's entire positioning — SRA/PDPA-compatible, "documents never leave the machine" — depends on there being **no backend**. Any standalone target that uploads document content to a server is categorically out.
2. **The codebase is already mostly host-agnostic.** A coupling audit (2026-06-09) found the feature "brains" — recognition, classification, diff, detection, the Config layer, the Findings model, the Section Index — largely free of Office JS. The Office dependency concentrates in a few seams: `shared/doc-model` (the single-pass `Word.run` loader), `shared/office`, `shared/ooxml` (the `getFileAsync`→`insertFileFromBase64` loop), and the `Word.Range` return type leaking through `Finding.location.getRange(ctx)`. Per-feature, the write/navigate edges (`repair`, `numbering`, `cross-reference`) are the only heavily-coupled feature code.

The question is *how* to deliver a standalone target without forking the codebase or breaking the invariant.

## Decision

Adopt a **ports-and-adapters (hexagonal) architecture** and ship the standalone target as a **static, backendless web app installable as a PWA**. Concretely:

1. **One host-neutral feature core.** All feature logic and the shared modules (`config`, `findings`, `section-index`, `doc-model` types, checker/diff/classify code) become **host-agnostic** — they import no `Office.*` / `Word.*` symbols. The core depends only on a **`HostPort`** interface.

2. **A `HostPort` interface (the single seam).** It declares the small set of capabilities the core needs from its environment: load the document into the existing `ParaNode[]` model; resolve a **host-neutral location** to something navigable; read/write the document package (OOXML in / OOXML out); persist key/value + document-scoped storage. `Finding.location` stops returning `Word.Range` and instead carries a **host-neutral location token** (e.g. `{ paraId, charStart, charEnd }`) that each adapter resolves its own way.

3. **Two host adapters implement `HostPort`.**
   - **`host-office`** — backed by Office JS, exactly as today (`Word.run`, `getFileAsync`, `insertFileFromBase64`, `range.select()`).
   - **`host-web`** — backed by the browser: a File-picker for ingest, the existing OOXML Service (PizZip/JSZip + fast-xml-parser, which is already host-agnostic) for parse, an **in-browser document model + view** for navigation, and a `Blob`-download for write-back. The web `doc-model` adapter must **compute rendered list numbers** (`"4.2.1"`, `"(b)"`) from `numbering.xml` itself, because there is no Word renderer to supply `listItem.listString`.

4. **Two thin app targets, one workspace.** `apps/word-addin` and `apps/web` each wire the core + UI to their adapter and own their build (Vite) and packaging (Office manifest vs PWA manifest). They share the `core` and `ui` packages. A feature edit is, by construction, an edit to both.

5. **The standalone target is an "analyse & export" tool, not an "edit-in-place" editor.** It imports `.docx`/`.pdf`, runs the existing checkers, shows findings, and **exports** a corrected `.docx` / rendered redline as a download. It does **not** attempt to rebuild Word's editing surface (live track-changes authoring/acceptance, native TOC fields, in-canvas redline).

6. **Packaging ladder: SPA → PWA → optional Tauri.** Ship as a static SPA; add a PWA manifest + service worker for an installed-app feel (desktop icon, own window, offline). A signed **Tauri** binary (system WebView, not Electron) is an *optional later layer* for formal distribution, reusing the same bundle unchanged.

7. **The no-host-import boundary is lint-enforced.** Extend the custom invariant ESLint rules: `core/**` and `ui/**` may not import `Office.*` / `Word.*` — only the adapter packages may. This turns "stay portable" from a discipline into a build error, alongside the existing no-`fetch` confidentiality lint.

## Rationale

- **Web-first beats binary-first** because the add-in is *already* a browser SPA targeting an evergreen + WASM baseline. Lifting it onto a static host is the smallest jump, needs no per-OS build/sign/notarise pipeline, and — critically — a static client-side SPA satisfies ADR-0001/0003 *by construction*, the same way the task pane does. A PWA delivers the "feels like a native app" experience (own icon, own window, offline) for a fraction of the cost of a binary.
- **Ports-and-adapters beats a fork** because the audit shows the core is already ~host-agnostic; the seams are few and named. Sharing one core is the only structure where "update once, ships to both" is true by construction rather than by vigilance.
- **"Analyse & export" beats "rebuild the editor"** because Word's editing surface (list rendering, track-changes, TOC fields, in-place redline) was *free* only because Word provided it. Rebuilding it standalone is an open-ended scope explosion and the most likely way this effort dies. Analyse-and-export preserves the product's value (the checkers) while keeping the new surface bounded.
- **Tauri-optional, not Electron-first** because a system-WebView shell is ~3–10 MB and reuses the WKWebView/WebView2 engines the add-in already runs in, versus Electron's ~100 MB bundled Chromium — and neither is needed at all for a PWA.

## Rejected alternatives

- **A web app with a backend (upload → server-side processing).** Rejected outright: violates ADR-0001/0003 and destroys the SRA/PDPA positioning that *is* the product. This is the one true red line. The standalone target must be backendless and in-browser.
- **Fork the codebase / copy features into a separate web repo.** Rejected: guarantees drift. Every feature fix would have to be applied twice, by hand, forever — exactly what the maintainer asked to avoid.
- **Rebuild Word's editing surface in the browser (a full in-browser .docx editor).** Rejected: open-ended scope (list-number rendering, track-changes authoring/acceptance, native fields, in-canvas redline), high fidelity risk, and not required to deliver the checkers' value. The standalone target is analyse-and-export.
- **Desktop binaries first (Electron/Tauri `.exe`/`.dmg` as the primary target).** Rejected as the *first* step: adds code-signing, Apple notarisation, auto-update, and per-OS build pipelines with no functional gain over a PWA for a backendless client-side app. Kept as an optional later layer (Tauri preferred) for formal distribution only.
- **Two parallel `vite` configs in a single non-workspace project (env-switched adapter alias) instead of a workspace.** A viable lighter alternative and explicitly allowed as a fallback, but not preferred: separate `package.json`s make the "core cannot import a host" boundary *physical* rather than conventional, which is worth the extra tooling for a load-bearing invariant.
- **Leave `Finding.location.getRange(ctx): Word.Range` as-is and special-case the web build.** Rejected: that return type *is* Office leaking into the shared model; special-casing it re-introduces host knowledge into the core and defeats the single-seam design. The location must become host-neutral.

## Consequences

**Easier / preserved:**
- All four invariants hold unchanged. The web target *reinforces* confidentiality (static, offline-capable, no network). Jurisdiction and "Compare never suppresses" are logic-level and port untouched.
- **The product's newer features are export-centric and port well.** AI-Readable Export (feature 11) is pure clipboard/download serialisation; the Compare export artefacts (ADR-0020 — read-only "candidate differences report" and "rendered redline" snapshot) are pure client-side generation to a downloadable file; Templates generate a filled `.docx`. These are *better* suited to an analyse-and-export web target than to the task pane. The Text/PDF Compare modes (ADR-0017/0018/0019) need no Word document at all and are the single best standalone fit.
- **Conversely, one feature is explicitly Word-only and does not port:** the Numbering upgrade (ADR-0023) is built on Word's native tracked-changes engine plus live Word selection (assisted removal as tracked revisions reviewed *in Word*, then human-selected re-apply). It has no standalone equivalent and stays add-in-only; the standalone target may still *detect* static numbers and strip prefixes on export, but not run that loop. The native updatable TOC field (feature 07) is likewise Word-only (standalone emits a static TOC instead). See `spec/features/12` §2.1 for the full per-feature compatibility matrix.
- Feature work continues to live in one core; both products ship it on next build. CI can build and deploy both the add-in static host and the web/PWA from one commit.
- An estimated ~70%+ of current logic ports with little change (`factual-consistency`, `party-role`, `defined-terms`, the Compare diff/classify and PDF path are already host-free).

**Harder / follow-on obligations:**
- **`Finding.location` must become host-neutral** (token, not `Word.Range`) — a cross-cutting refactor through the Findings model and every consumer. This is the load-bearing change and should land early.
- **The web `doc-model` adapter must compute rendered list numbers** from `numbering.xml` (no Word renderer). Bounded — Repair/Numbering already manipulate `numbering.xml` — but the single genuinely-hard new component.
- **Round-trip fidelity is now 100% on the web target** (no Word managing the package). Expect a validation burden; the OOXML Service's existing fidelity caveats (ADR-0003) apply with no host safety net.
- **This is a new product surface, sequenced after the v1 add-in features**, not a v1 feature. It is gated into scope by this ADR + the scope register (row 14) with explicit maintainer approval. It should be migrated **incrementally**: introduce `HostPort`, move one Office-free feature (`factual-consistency`) behind it as a proof, then convert the write/navigate-heavy features.
- **The custom-lint surface grows** by one rule (no host imports outside adapters), joining the existing invariant rules in `eslint-plugins/invariants.cjs`.

The planning and execution of this surface is specified in `spec/features/12-standalone-implementation.md` and sequenced as **Phase 7** in `spec/README.md`.
