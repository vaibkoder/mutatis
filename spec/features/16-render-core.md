# 16 — RenderCore (own client-side redline rendering engine)

**Status:** ✅ in scope. Established by **ADR-0024**. the project scope (row 15). Sequenced as **Phase 8** in `spec/README.md`.

> A custom, worker-resident, host-free layout/rendering engine that turns Mutatis's already-solved comparison diff into a professional paginated redline: (a) an on-screen paginated preview and (b) a downloadable **vector PDF** with selectable text and embedded subset fonts — both painted from **one shared layout result**. Replaces the HTML-flow + `window.print()` renderer as the Word-vs-Word / Text-vs-text rendered-redline path when the integration phase (§6 Phase 2) lands — until then the ADR-0020 path stays in service. 100% client-side; zero network egress at render time.

---

## 0. How to use this doc

This is the design input to the two delivery phases (§6). Phase 1 builds the engine; Phase 2 integrates and hardens it.

**Read first:** `docs/adr/0024-rendercore-own-client-side-redline-rendering-engine.md` (the decision + rejected alternatives + fallback ladder), then this doc.

---

## 1. Purpose & priority order

The redline's core value is **utility: a lawyer must reliably know exactly what changed**. Everything else serves that. Priority order (explicit from the user — binding on every phase):

1. **Robust and reliable** — every change visibly and unambiguously marked; nothing suppressed or silently dropped (ADR-0008).
2. **Quality** — real multi-level auto-numbering with hanging indents, bold headings, justified body text, preserved run formatting (bold/italic/underline/superscript), footnotes at the bottom of the correct page with a separator, A4 pagination, output as a real downloadable vector PDF with selectable text.
3. **Fast** — no UI freeze ever (layout off the main thread); ~150-page document in low single-digit seconds; shipped bundle delta in the low single-digit MB.
4. **Aesthetic exactness is NOT a goal in itself** — a simple, configurable, standard tracked-change theme; no exotic decorations.
5. **Line-for-line identity with Word is a NON-GOAL** — never an acceptance criterion. The bar is *professional, internally consistent* pagination.

## 2. Scope

### 2.1 In scope ✅

- **The layout engine** (`harfbuzzjs` shaping in a Web Worker; greedy first-fit breaking with Word's justification semantics — space expansion + ~20% inter-word shrink; usWin line-height formula; kerning/ligatures off, matching Word defaults).
- **The paginator** — A4 geometry from `sectPr`, widow/orphan (2 lines), keepNext/keepLines, pageBreakBefore, explicit breaks; **footnote** bottom-of-page reservation via fixpoint iteration, separator, oversized-footnote split with continuation; **basic table** layout with row splitting + cantSplit.
- **The font subsystem** — vendored OFL metric-clones (Liberation Serif/Sans ↔ TNR/Arial, Carlito ↔ Calibri, Caladea ↔ Cambria; ≈1.5–3 MB WOFF2), `.odttf` embedded-font extraction (XOR deobfuscation), `FontFace` registration for the preview, raw bytes retained for shaping + PDF subset embedding.
- **Input model** — `authoredBodyXml` (native `<w:ins>/<w:del>/<w:moveFrom*>/<w:moveTo*>` markup) + `styles.xml` (full style-cascade resolution: docDefaults → style chain → numbering-level pPr → direct formatting, with toggle-property semantics) + the **existing numbering resolver's** labels + `footnotes.xml` + theme/embedded fonts when present.
- **Output A — paginated preview**: absolutely-positioned spans/SVG per line painted from the computed boxes; virtualised page list; progress reporting; theme via CSS.
- **Output B — programmatic vector PDF**: `@pdfme/pdf-lib` + `fontkit` subset embedding; decorations drawn as vector ops; selectable text; opens in Acrobat/Preview/Chrome. **Not** via the browser print dialog.
- **Tracked-change decoration** — one simple configurable theme: e.g. coloured underline + highlight for inserts, coloured strikethrough for deletes, paired styling + anchors for moves, an inline change anchor (e.g. `*`), left-margin change bar. Styling may use the classifier's 8-class output as annotations; styling never filters (ADR-0008).
- **Integration** — replace the old flow renderer behind the existing surfaces: Word/Text rendered-redline snapshot export (programmatic PDF download), the inline-redline preview (feature 14) upgraded to the paginated preview; wired in **both hosts** (standalone web app — primary target — and the task pane) through the existing driver/HostPort seams.
- **Text mode** — TEXT compares render via the **same engine** with synthesised default styles (no separate budget).
- **Host-free, offline, lazy-loaded** — no `Office.*`/`Word.*` imports; all assets vendored same-origin; code-split so non-render flows pay nothing.

### 2.2 Out of scope / deferred 🔜❌

- ❌ **PDF-vs-PDF / Word-vs-PDF modes** — they produce a *candidate differences report*, not a redline; their existing ADR-0020 export path stands unchanged.
- ❌ **Line-for-line Word pagination parity** as an acceptance bar or test assertion (ADR-0024; cut in the project scope). Residual divergence is a calibration budget.
- ❌ **Aesthetic/decoration features beyond the one simple configurable theme** (no theme gallery, no double underline et al.) (ADR-0024; cut in the project scope).
- ❌ **Any editing surface** — RenderCore renders read-only artefacts; ADR-0022's no-in-browser-editor cut is untouched. The editable `.docx` / host-doc redline remains the primary, unchanged output of the Word/Text modes.
- ❌ **AGPL dependencies; any network dependency at render time; any cloud rendering.**
- 🔜 **Headers/footers rendering** — body + footnotes only in v1 (the per-paragraph model's `isHeaderFooter` flag exists; rendering headers/footers on each page is deferred).
- 🔜 **Endnotes; comments/annotations rendering; OMML math; drawing layer/charts.**
- 🔜 **Images** — v1 renders a visible placeholder box (never silently drops content occupying space); full raster/vector image painting is deferred.
- 🔜 **Hyphenation, kerning/ligatures-on documents, CJK/RTL line-breaking rules** — Word's Latin defaults (all off) are the v1 regime; HarfBuzz shapes what it is given, but non-Latin breaking correctness is deferred.
- 🔜 **Nested tables beyond one level; exotic table features** (vertical merge fidelity beyond basics) — flag-and-degrade gracefully, never crash.
- 🔜 **Migrating the PDF-modes candidate-differences report onto RenderCore** — possible later quality upgrade, not in this track.

### 2.3 Invariant check (must hold every phase)

- **Confidentiality** — zero network egress at render time: no CDN, font, or package fetches; everything vendored same-origin, offline (ADR-0001/0003). The office-free/zero-egress grep discipline extends to **all** chunks the render path loads (see memory: grep every loaded chunk, not just the entry).
- **Jurisdiction** — nothing US-hardcoded; the renderer is convention-neutral (it renders what the document says); any configurable defaults (e.g. page-size fallback when `sectPr` is absent) read from the Config layer (A4, not US Letter).
- **Scope** — build only what this doc marks ✅; blockers escalate via the ADR-0024 fallback ladder, never via silent improvisation.
- **Compare never suppresses** — the renderer categorises and styles changes but renders **every** one (ADR-0008); acceptance criterion 1 is the property test.

## 3. Architecture (settled — ADR-0024 / brief §4.4; do not re-litigate)

```
                      ┌──────────────────────── Web Worker ────────────────────────┐
authoredBodyXml ─┐    │  Input model        Layout core           Paginator        │
styles.xml ──────┼──► │  (style cascade,    (harfbuzzjs shaping,  (widow/orphan,   │
numbering labels ┤    │  rPr/pPr resolve,   greedy+shrink break,  keeps, footnote  │
footnotes.xml ───┤    │  revision spans,    justification, usWin  fixpoint, table  │
fonts (clones +  ┘    │  numbering attach)  line height)          row split)       │
 odttf extract)       └──────────────┬─────────────────────────────────────────────┘
                                     │  ONE layout result (pages → lines → positioned runs + decorations)
                          ┌──────────┴──────────┐
                          ▼                     ▼
                   Preview painter        PDF painter
                   (positioned spans/     (@pdfme/pdf-lib + fontkit
                   SVG, CSS theme,        subset embed, vector
                   virtualised pages)     decorations, metadata)
```

- **Primary input is `authoredBodyXml`** — the same merged bytes used for the editable `.docx`, so rendered and editable artefacts stay in lock-step. The renderer does not consume the diff-op flow model; change classes arrive as optional annotations keyed to revision spans/paragraphs, defined in RenderCore's own types (the compare feature maps into them — keeps `shared → features` dependency direction clean).
- **Reuse, don't rebuild:** the comparison engine, per-paragraph model (rawXml carries every `rPr`), and the host-neutral, tested **numbering resolver** (lvlText %k substitution, numFmt, lvlRestart, startOverride/lvlOverride, isLgl) are the foundation. The resolver's known deferred bounds (`<w:isLgl/>` applied one level too far, so a `decimalZero` sub-level renders `1.1` where Word renders `1.01` — 08c's to fix; bullet glyphs; footnote-body lists; and, in RenderCore itself rather than in the resolver, `parse-body.ts` not yet reading the `pStyle` → `numPr` chain) are closed **as needed** during the Phase 1 build (the tables/numbering step). **Two former entries on this list are closed:** the resolver follows `<w:numStyleLink>` to a list style's definition, and counters belong to the abstract definition rather than to the `<w:num>` — each was closed because, while open, it made a real change invisible to Compare, which shares this resolver (ADR-0042). *"Style-linked numbering"* named both of those and the `parse-body` gap; that ambiguity is why the first read as somebody else's problem for twenty-three review passes.
- **Worker boundary:** the entire parse→layout pipeline runs in the worker (no DOM dependency anywhere in it); the main thread only paints. Message protocol with transferables, progress events, and cancellation.
- **Determinism:** measurements come from font bytes via HarfBuzz, not the browser text engine → identical layout on Chromium and WebKit, and identical between preview and PDF by construction (one layout result).
- **Fonts:** requested-family → metric-clone mapping (TNR→Liberation Serif, Arial→Liberation Sans, Calibri→Carlito, Cambria→Caladea; documented no-clone cases like Garamond get the nearest visual substitute *for display* with explicit non-metric flagging); embedded `.odttf` fonts win over clones when present. Preview may *display* a locally-installed real Microsoft font (metrics match the clones, so positions stay valid); shaping and PDF embedding always use the vendored/extracted bytes.
- **Suggested home** (confirm in phase planning): host-free engine under `src/shared/render-core/` (pure; no feature imports; no DOM in the layout pipeline), worker entry under `src/workers/`, vendored fonts under `public/fonts/` (or equivalent same-origin asset path) with their OFL licence files. Lazy-loaded behind a dynamic import from the Compare surfaces.
- **Word's layout rules being implemented** (the de-facto public spec is LibreOffice `sw/source/core/text/`; key facts verified by investigation): greedy first-fit breaking; justification = inter-word space expansion + up-to-~20% inter-word shrink (tdf#119908); kerning/ligatures/hyphenation off by default; line height "single" = `(usWinAscent + usWinDescent) / unitsPerEm × fontSize` (typo metrics iff USE_TYPO_METRICS); modern Normal = Calibri 11pt, 1.08 line, 8pt space-after; pagination flags are documented OOXML (widowControl default-on 2 lines, keepNext on heading styles, keepLines, pageBreakBefore, cantSplit, footnote reservation); relevant compat flags are enumerable.

## 4. The hard parts (call these out in each phase's plan)

1. **Footnote bottom-of-page reservation** — a fixpoint iteration (placing a footnote shrinks the body area, which can move the reference off the page, which moves the footnote…), plus oversized-footnote splitting with a continuation notice. The single hardest component; Spike B (the opening move of Phase 1) de-risks the algorithm before the engine is built around it.
2. **Style-cascade resolution** — OOXML effective-property computation (docDefaults → basedOn chain → numbering-level pPr → direct formatting, toggle properties XOR semantics) is fiddly and load-bearing for everything downstream. Pure data; test exhaustively against real fixtures.
3. **Sub-pixel rounding accumulation** — small advance-width rounding choices compound across a justified line; calibrate empirically in Spike A and keep the rounding policy fixed and documented thereafter.
4. **Table layout** — column widths (tblGrid/tblW), cell wrapping, row splitting, cantSplit; degrade gracefully on exotic structures (render content; never crash, never drop).
5. **Performance at ~150 pages** — incremental layout, transferable buffers, progress + cancellation; the worker keeps the main thread under ~50 ms blocks.
6. **WebKit** — the project's prior lessons apply (e.g. the ReadableStream async-iteration polyfill); verify wasm + FontFace + Blob-download behaviour on WebKit in the spikes and at every integration point.

## 5. Acceptance criteria (derive tests from these — never from Word screenshots)

1. **Never-suppress (property test):** every diff op in the change-set appears in the rendered artefact, styled by its class — property-test against the comparison engine's output (ADR-0008).
2. **Auto-numbering:** auto-numbered documents render correct labels and hanging indents (no regex scraping of literal text).
3. **Run formatting:** bold/italic/underline/superscript preserved from `rPr`.
4. **Body typography:** justified body text; bold headings; A4 geometry from `sectPr`.
5. **Footnotes:** superscript reference in the body; footnote body at the bottom of the page containing the reference, with separator; splits correctly when oversized.
6. **Pagination semantics:** widow/orphan (2) + keepNext honoured (no orphaned heading at a page bottom).
7. **PDF:** vector, selectable text, fonts embedded as subsets, opens in Acrobat/Preview/Chrome; **PDF pagination identical to the preview** (one layout result).
8. **Performance:** no main-thread block > ~50 ms during render; ~150-page document in low single-digit seconds; progress reporting.
9. **Cross-engine:** works on Chromium and WebKit.
10. **Bundle & egress:** delta within ~5 MB including fonts; **zero network egress** (extend the office-free grep discipline to ALL chunks the render path loads).
11. **Text mode:** renders via the same engine with synthesised default styles.

## 6. Roadmap — two phases (committed)

RenderCore ships in **two phases**: Phase 1 builds the engine and Phase 2 integrates and hardens it. When Phase 2 is done the engine ships and the track is closed.

**The scope is unchanged** — every component (fonts, style cascade, layout core, paginator, footnotes, tables, decoration, PDF, preview, integration, Text mode, hardening) is still here. It is consolidated into two buckets at the natural seam: **build the engine** (Phase 1), then **wire it in and harden** (Phase 2). Each phase lands green (lint/typecheck/test). Phase 1 ships tested library code with no user-visible surface; the wiring arrives in Phase 2.

### Phase 1 — The engine: spikes → layout core → both outputs



**1a — De-risk first (opening move; gates the rest of the phase).** Before building, prove the thesis with three throwaway spikes under `spikes/render-core/` (outside the shipped bundle; excluded from the office-free grep):
- **Spike A (thesis test):** harfbuzzjs + vendored Liberation Serif in a Web Worker; lay out a sample of justified paragraphs with the greedy+shrink breaker and the usWin line-height formula; overlay against Word's rendering of the same `.docx` (user-assisted check).
- **Spike B:** footnote-reservation fixpoint on a synthetic doc, incl. a footnote longer than a page (split + continuation).
- **Spike C:** @pdfme/pdf-lib + fontkit subset-font emission of a few laid-out pages; verify on Chromium **and** WebKit; selectable text + file size.
- **GO/NO-GO gate:** GO needs all three spikes' evidence (overlay result, fixpoint output, PDF artefacts + sizes) recorded with the phase; Spike A alone can NO-GO. **NO-GO ⇒ STOP, record evidence, escalate the ADR-0024 fallback ladder (Typst.ts → OnlyOffice) — do not build.** Fold the calibration findings (rounding policy, shrink behaviour) into the build as fixed policy.

**1b — Build the engine** (dependency order, same session, once GO):
- **Font subsystem** — vendor OFL clones (Liberation Serif/Sans, Carlito, Caladea × 4 styles; WOFF2 + raw bytes for shaping/embedding) + licence files, same-origin; requested-family→clone mapping + fallback chain + no-clone flagging; `.odttf` extraction (XOR deobfuscation) with embedded fonts winning over clones; `FontFace` registration (Safari family-name quoting); lazy-load packaging; zero-egress guard/test extended.
- **Layout input model** — parse `authoredBodyXml` + `styles.xml` + `sectPr` into a resolved, host-free render model: full style cascade (docDefaults → style chain → numbering-level pPr → direct; toggle-prop semantics), effective rPr/pPr (indents incl. hanging, spacing, justification, keeps), A4 geometry (Config default when absent), revision spans from native `<w:ins>/<w:del>/<w:moveFrom*>/<w:moveTo*>` + class annotations (RenderCore-owned types), numbering-resolver labels + level indents, footnote refs + `footnotes.xml` bodies, image placeholders.
- **Layout core (worker)** — message protocol (transferables, progress, cancellation); HarfBuzz shaping (kerning/ligatures off) + advance-width measurement; greedy first-fit breaker; justification = space expansion + ~20% shrink; usWin/typo line height; space before/after; tabs + hanging indents. Spike-A calibration as fixed policy. Output: line boxes with positioned runs, off the main thread.
- **Paginator** — page filling; widow/orphan (2); keepNext/keepLines; pageBreakBefore; explicit breaks; per-section page geometry.
- **Footnotes** — bottom-of-page reservation fixpoint (per Spike B); separator; oversized split + continuation; superscript refs; fixpoint termination bound asserted.
- **Tables (+ numbering-resolver bound-closure as needed)** — tblGrid/tblW widths; cell wrapping via the same breaker; row split + cantSplit; basic borders/shading; graceful degradation (render content, never crash/drop). Close resolver deferred bounds real fixtures hit — **still open:** bullet glyphs and footnote-body lists. (Per-abstractNum counters and `<w:numStyleLink>` are **closed**, see §3.) As needed, not speculatively.
- **Tracked-change decoration + theme** — insert underline+highlight, delete strikethrough, move pairing + anchors, inline change anchor, left-margin change bar; one simple configurable theme (colours, highlight on/off) — no gallery.
- **PDF backend** — @pdfme/pdf-lib + fontkit subset embedding; paint positioned runs; decorations as vector ops; superscript; metadata; verified on Chromium + WebKit; opens in Acrobat/Preview/Chrome; selectable text.
- **Preview painter** — absolutely-positioned spans/SVG per line from the **same** layout result; page chrome; virtualised page list; progress UI; theme via CSS variables (author the stylesheet section).

**Phase 1 exit:** as tested library code, the engine turns a real Compare change-set into a paginated preview **and** a vector PDF painted from one shared layout result. Acceptance criteria **1–7 evidenced at engine/output level** — never-suppress property test (every change present + decorated by class), auto-numbering labels/indents, run formatting, justified body + bold headings + A4, footnote placement incl. oversized split, widow/keepNext, PDF vector/selectable/subset-embedded **with preview↔PDF pagination parity** (both painters exist by end of phase). Lint/typecheck/test green; bundle delta recorded; zero render-time egress. No user-visible surface yet.

### Phase 2 — Integration + hardening (replace the old renderer; ship)



- **Wire RenderCore behind the existing surfaces in both hosts** (standalone web — primary — and the task pane) via the driver/HostPort seams: Word/Text rendered-redline snapshot → programmatic PDF download (replaces `window.print()` for these modes); inline-redline preview (feature 14) → paginated preview. The Compare feature maps its change classes into RenderCore's annotation types.
- **Text mode** via synthesised default styles (acceptance 11).
- **Retire** the old flow-renderer path for these surfaces; the PDF-modes candidate-differences report is untouched.
- **Hardening sweep:** ~150-page perf (low single-digit seconds; no >50 ms main-thread block) measured + tuned; memory/transferables audit; cross-browser (Chromium + WebKit, the project's WebKit lessons); bundle budget ≤ ~5 MB delta; zero-egress + office-free grep across **all** render-path chunks.
- **Full acceptance sweep 1–11 with pasted evidence** — the artefact-level re-check of criteria 1 and 7 now that the surfaces exist, plus 8/9/10 (perf, cross-engine, bundle/egress) measured here.

**Phase 2 exit:** all 11 acceptance criteria evidenced; old path retired without regression (full suite green); editable `.docx`/host-doc redline provably untouched and primary; feature marked complete; track closed.

**Sizing note.** The split is at the build/integrate seam so Phase 1 is self-contained (a tested engine) and Phase 2 is the regression-sensitive wiring + hardening. Do **not** drop the footnote, table, or PDF work — they are the point of the engine. If Phase 1's spikes NO-GO, the whole track escalates to the fallback ladder before any build.

## 7. Anti-goals / do-not-do list (binding on every phase)

- Do NOT re-survey rendering engines or re-litigate the architecture (ADR-0024 §Rejected alternatives is evidence-backed). A genuine blocker ⇒ STOP, surface it with evidence, use the recorded fallback ladder (Typst.ts → OnlyOffice stack).
- Do NOT make line-for-line Word parity an acceptance criterion or test assertion.
- Do NOT add any network dependency, AGPL library, or Word-host dependency.
- Do NOT build aesthetic features beyond the one simple configurable theme.
- Do NOT discard or work around the existing comparison engine, `authoredBodyXml` path, numbering resolver, or per-paragraph model — they are the foundation.
- Do NOT drop the footnote, table, or PDF work when consolidating — they are the point of the engine.

## 8. References

- LibreOffice Word-justification parity (tdf#119908): numbertext.org/typography · LO 24.2 release notes. LibreOffice `sw/source/core/text/` = the de-facto public spec for Word layout.
- OOXML compat settings: MS-DOCX compatibilityMode docs; OOXML Part 4 compat topic.
- Line metrics: SIL FDBP Line Metrics; OS/2 usWin vs typo metrics.
- Fonts: github.com/liberationfonts/liberation-fonts · Carlito/Caladea (Google, OFL) · Carlito mismatch: google/fonts#9720.
- odttf deobfuscation: gist.github.com/dungsaga/ab8d2379bb566c9925b27df3bc82ca8b.
- harfbuzzjs: github.com/harfbuzz/harfbuzzjs (v1.2.1, 2026-06).
- PDF backend: github.com/pdfme/pdfme (@pdfme/pdf-lib fork; Hopding/pdf-lib unmaintained); pdfkit = recorded alternative backend.
- Fallbacks: typst.ts github.com/Myriad-Dreamin/typst.ts (~12 MB gz, Apache-2.0) · OnlyOffice pattern: github.com/cryptpad/onlyoffice-x2t-wasm (AGPL — needs its own ADR + licence decision if ever reached).
