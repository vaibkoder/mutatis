# ARCHITECTURE — Technical Foundation

> Consolidated technical foundation for Mutatis. This is the **single source of truth** for the client-side model, shared modules, technology stack, and Office JS constraints. Feature docs reference these headings by anchor and must not restate them. Where a feature doc conflicts with this document on a shared concern, **this document wins**.
>
> See also [README.md](README.md) (product index) and [CONVENTIONS.md](CONVENTIONS.md) (SG/UK config values).

---

## Client-Side Model

The add-in is a single-page web application hosted **entirely inside the Word task pane**, an embedded browser control whose engine varies by host (WebView2/Chromium on Windows, WKWebView on Mac, Chrome WebView on Android, the host browser on Word Online). All of these support modern JavaScript and WebAssembly, so the codebase targets an evergreen-browser baseline. All document interaction goes through the **Office JS API** (`Word.run(...)` and `Office.context.document`), gated by **WordApi requirement sets**.

**Confidentiality-first, fully client-side principle.** The entire application runs in the WebView; document content **never** transmits to any server. This is the product's core positioning and is compatible with SRA (UK) and PDPA (SG) confidentiality obligations.

**Data-residency guarantee (by construction, not policy).**
- All document reads (`getFileAsync`, typed API) deliver bytes **into JS memory in the WebView**.
- All processing (unzip, parse, diff, edit) happens **in the WebView's JS/WASM sandbox**.
- The only writes are **back into the same document** via `insertFileFromBase64` / `createDocument`.
- The app makes **no fetch/XHR/WebSocket calls that carry document content**. The only permitted network calls are (a) loading the add-in's own static assets from its hosting origin, and (b) **explicit, opt-in** storage adapters (GitHub/SharePoint) carrying only user-authored library data, never the working document, and off by default.

This must be enforced as a code-review **invariant** and, ideally, a build-time lint rule flagging any `fetch`/`XMLHttpRequest` outside the sanctioned storage-adapter modules.

**Two interaction models.** Choosing per operation is a recurring architectural decision:
- **(A) Typed Office JS API** — the default. `Word.run(async ctx => …)` for paragraphs, ranges, styles, lists, footnotes, tracked changes, fields, tables — everything the typed API can express.
- **(B) Client-side OOXML read-modify-write loop** — the escape hatch for what the typed API cannot express (stripping `w:rsid*`, deleting section breaks, surgical `numbering.xml` edits, removing orphaned bookmarks, authoring redline `<w:ins>`/`<w:del>`). See [The OOXML Service](#the-ooxml-service). No bytes leave the machine during this loop.

---

## The OOXML Service

`shared/ooxml` encapsulates the entire read-modify-write loop behind a clean API so feature code never hand-rolls `getFileAsync` slice plumbing or base64. **Consumed by Repair (v1) and Compare (v2).**

**Authoritative call sequence:**

```text
1. Office.context.document.getFileAsync(
       Office.FileType.Compressed, { sliceSize: <bytes> }, cb)
   → returns the LIVE document as .docx bytes, delivered in slices.
     Read each slice via file.getSliceAsync(i, cb); concatenate to a Uint8Array;
     file.closeAsync().

2. Unzip in-memory (PizZip default; JSZip for async/streaming — see Technology Stack).
   Edit the relevant XML parts (word/document.xml, word/numbering.xml,
   word/styles.xml, word/settings.xml, word/footnotes.xml, header*.xml …)
   as strings or via fast-xml-parser. Re-zip. Base64-encode.

3. Word.run(ctx => {
       ctx.document.body.insertFileFromBase64(base64, Word.InsertLocation.replace);
       return ctx.sync();
   });
   → writes the modified document back IN PLACE.
```

**Fallback path.** `insertFileFromBase64(..., replace)` has documented edge-case artifacts (e.g. a stray leading empty paragraph / blank page in some builds). When an operation is sensitive to this, fall back to:

```text
Word.createDocument(base64).open();   // opens the cleaned result in a NEW document
```

**Service responsibilities.** Drive the slice loop; unzip via PizZip and expose parts as editable strings; re-zip/base64/`insertFileFromBase64`; detect in-place artifacts and (when configured) fall back to `createDocument().open()`; enforce the two-documents-in-memory limit by serialising concurrent `getFileAsync` requests and calling `closeAsync()` promptly.

**Interface (sketch):**
```ts
interface OoxmlService {
  readDocx(opts?: { sliceSize?: number }): Promise<DocxBundle>;
  writeBack(bundle: DocxBundle, opts?: {
    mode?: 'replace' | 'newDocument';                 // default 'replace'
    onArtifactDetected?: 'fallbackToNewDoc' | 'throw';
  }): Promise<void>;
}
interface DocxBundle {
  part(path: string): string | undefined;             // e.g. 'word/document.xml'
  setPart(path: string, xml: string): void;
  hasPart(path: string): boolean;
  listParts(): string[];
  removePart(path: string): void;
  stripAttributes(partPath: string, attrNames: string[]): void;  // e.g. w:rsid*
}
```

**Known fidelity caveats.** `insertFileFromBase64` can drop/alter package-level constructs (certain styles, numbering links, section properties) when inserting a built package into an existing document. For freshly-authored output packages (Compare), prefer materialising a **standalone new document** (`createDocument().open()`) and use in-place insert only for insert-into-active-doc; validate the round-trip and fall back on detected fidelity loss. Some package-level XML (`fontTable.xml`, `theme/theme1.xml`, certain `styles.xml` internals) may not be safely editable client-side in v1 — do not attempt to rewrite these.

---

## Shared Modules

These are built first and depended upon by multiple features. Each is the **single source of truth** for its concern; feature docs reference them by heading. `src/shared/**` must **not** import from `src/features/**` — dependencies flow features → shared, never the reverse.

### Document Model / Paragraph Index

`shared/doc-model`. A **single-pass loader** that reads the entire document once into an in-memory, queryable structure every feature reads from. Most performance-critical shared module.

- **Inputs:** the live document (one batched `Word.run` load).
- **Captures per paragraph:** index/order, plain `text`, named `style` + key direct-format flags, `listItem.listString` (rendered number e.g. `"4.2.1"`, `"(b)"`; guarded — not every paragraph is a list item) and list level, table membership (`parentTableCell`), and a `getRange(ctx)` to re-resolve a live Range on demand (never hold detached proxies across syncs).
- **Footnotes:** `document.body.footnotes` (WordApi 1.5) loaded as a parallel collection tagged `container: 'footnote'`.
- **Outputs:** `paragraphs: ReadonlyArray<ParaNode>`, `footnotes`, `getById(id)`, and a `(paragraphId, charOffset) → Word.Range` resolver for navigation.
- **Caching & invalidation:** cached after `load()`; registers `onParagraphChanged/Added/Deleted` (WordApi 1.4), **coalesced and debounced** (400–600 ms); marks itself stale and re-reads targeted or full; features subscribe via `onInvalidate`. There is **no `document.onChanged`** — this granular-event path is the only one. Selection-driven UI uses `DocumentSelectionChanged` separately (see [Task Pane UI Shell](#task-pane-ui-shell)).

```ts
interface ParaNode {
  id: string; index: number; text: string; styleName: string;
  listString?: string; listLevel?: number;
  container: 'body' | 'footnote' | 'tableCell';
  tableInfo?: { tableId: string; row: number; col: number };
  getRange(ctx): Word.Range;
}
```

### Section Index Builder

`shared/section-index`. Derives the logical legal structure — clauses, sub-clauses, schedules, parts, recitals, annexures, definitions — from the Document Model's paragraphs by combining **named styles** with **`listString`** auto-numbers. **Authored by the Cross-Reference feature; consumed by Cross-Reference, Interactive Lookup, and Numbering.** Feature docs must not re-specify it.

- **Inputs:** `DocModel.paragraphs` + the active jurisdiction profile (terminology + numbering presets — see [CONVENTIONS.md](CONVENTIONS.md)).
- **Outputs:** a tree of `SectionNode`s + a flat `byKey` map (`number → SectionIndexEntry[]`; an **array** so duplicate keys drive Ambiguous classification) + a defined-terms index. Each entry carries a normalised key (`clause:4.2`, `schedule:2/para:3`, `recital:B`, `part:II`), display label, raw number, parent key, level, paragraph id, lazy range, `isStaticNumbered`, and bookmark names (for REF resolution). Statically-numbered units are still indexed (fallback leading-number regex) so they don't produce false "Not Found".
- Rebuilt on `DocModel` invalidation.

### Storage Abstraction

`shared/storage`. A uniform key/value + document-scoped store with import/export and pluggable remote adapters. **Consumed by Numbering schemes, Clause Library, Templates, and Shared Data Sets.**

- **Tiers:** `localStorage` (machine-local user prefs / scheme library); `Office.context.document.settings` (per-document, persisted inside the `.docx`, `saveAsync()`, ~2 MB budget shared across features); **JSON import/export** (user-initiated file download/upload for portability/sharing without a server); **pluggable adapters** `GitHubAdapter` / `SharePointAdapter` (opt-in, user-configured, carry only user-authored library content — never the working document; default disabled).

```ts
interface Store<T> {
  get(key: string): Promise<T | undefined>;
  set(key: string, value: T): Promise<void>;
  list(prefix?: string): Promise<string[]>;
  export(): Promise<Blob>;            // JSON
  import(blob: Blob): Promise<void>;
}
interface StorageProvider { local: Store<unknown>; docSettings: Store<unknown>; remote?: RemoteAdapter; }
```

### Findings Model

`shared/findings`. A **single shared structure** for every issue surfaced in the Proof/Review tabs, so all checkers feed **one uniform UI renderer**. A checker's only obligation is to emit `Finding`s; it owns no UI.

```ts
type Severity = 'error' | 'warning' | 'info';   // features also use 'critical'/'low'/'medium' labels — map to this enum
interface Finding {
  id: string;
  category: string;                 // e.g. 'cross-reference', 'defined-terms/used-not-defined'
  featureId?: string;               // hard discriminator to keep distinct features ungrouped (see Cross-Reference A vs B)
  severity: Severity;
  title: string; detail?: string;
  location: { paraId?: string; getRange(ctx): Word.Range; label: string };
  suggestedAction?: { label: string; apply(ctx): Promise<void> };
}
```

The store aggregates `Finding[]` per category, filters by severity/category, and exposes counts for tab badges.

### Localisation / Config Layer

`shared/config`. A **single central config object** holding **Singapore** and **UK (E&W)** drafting conventions. **Every checker, formatter, and structure builder reads its conventions from here. No feature may hardcode US conventions** (no `$` default, no `mm/dd/yyyy`, no "Section" as the only clause noun). US-centric defaults are a **bug**.

The concrete values — terminology, defined-term conventions, currency table, dates, placeholders, numbering presets, entity/template fields, court formatting profiles, en-GB spelling — are **specified in [CONVENTIONS.md](CONVENTIONS.md)**. This module is the runtime carrier; CONVENTIONS.md is the data. A default profile (UK or SG) is selected at first run and persisted via Storage; switching jurisdictions means selecting a profile, never editing checkers.

### Task Pane UI Shell

`taskpane/shell`. The chrome hosting all features:
- **Tab navigation** across feature groups (Proof / Review / Cross-Refs / Numbering / Repair / Content / Compare / Settings) with **badge counts** from the Findings store.
- **Findings list with click-to-navigate**: each row renders a `Finding`; clicking resolves its range and calls `range.select()` (lazy resolution to avoid stale ranges). Optional one-click `suggestedAction.apply`.
- **Preview region** (used by Interactive Lookup) for read-only target previews.
- **Settings panel**: jurisdiction profile switch (UK/SG), storage/adapter configuration, capability/host diagnostics (which WordApi sets are available), import/export of libraries and schemes.
- **Office lifecycle gate**: mounts only after `Office.onReady`; reflects `Office.context.officeTheme`; subscribes to `DocumentSelectionChanged` (debounced) for "current clause" context.

---

## Technology Stack

- **Language — TypeScript (mandatory)**, `strict` mode (`strictNullChecks`, `noImplicitAny`, `exactOptionalPropertyTypes`). Static typing is load-bearing given the nested `ctx.load`/`ctx.sync` orchestration and OOXML string surgery.
- **UI framework — Preact (≈3 kB) + `@preact/signals` (recommended).** Narrow sidebar (~320–400 px) needs fast first paint and a small bundle; signals give fine-grained Findings-list reactivity. *Acceptable alternative:* plain TS + lit-html. **Avoid** Angular (weight) and full Fluent UI React; use Fluent UI **Web Components** or CSS tokens matching `Office.context.officeTheme` for visual consistency.
- **Build tooling — Vite, scaffolded from the Office Add-in generator.** Bootstrap manifest + dev-cert + sideloading via `yo office` once, then migrate the bundler to **Vite**. Keep `office-addin-debugging` / `office-addin-manifest` / `office-addin-dev-certs` as dev deps. Output is static HTML/JS/CSS over HTTPS; production is any static host (GitHub Pages, Cloudflare Pages, Netlify) serving only the app shell.

**Key libraries:**

| Concern | Library | Notes |
|---|---|---|
| Unzip/zip `.docx` (sync) | **PizZip** | **Default** for the synchronous OOXML loop; standardise on one ZIP lib across the bundle. |
| Unzip/zip `.docx` (async/streaming) | **JSZip** | For streaming very large packages (Compare §large-doc). |
| XML parse/serialize | **fast-xml-parser** | Structured edits (numbering.xml, styles.xml). For simple attribute stripping (`w:rsid*`) prefer targeted regex to preserve byte fidelity. |
| Text diff (Compare) | **diff-match-patch** | Word-mode diff inside changed paragraphs. |
| Short-field diff (Compare) | **fast-diff** | Cells/headings where DMP is overkill. |
| NL date parsing | **chrono-node** | Configured **little-endian (day-first)** — never US month-first. |
| Spelled-number → numeric | **words-to-numbers** | Wrapped for British inserted "and", currency-word strip, suffix strip, sub-units. |
| Library fuzzy search | **Fuse.js** | Client-side clause-library index (≈2000-entry ceiling — see Clause Library). |
| PDF rendering | **PDF.js** | **v2 only** — Word-vs-PDF compare / import; client-side, zero-upload. |
| Heavy CPU fallback | **Pyodide (~6.4 MB WASM)** | **Not used in v1**; lazy-loaded only if a future feature is provably CPU-bound. |

**WordApi requirement-set floor: 1.5.** Rationale: footnote bodies (`document.body.footnotes`) and native TOC field insertion (`Range.insertField(..., FieldType.toc)`) are core. WordApi 1.3 (`listItem.listString`) and 1.4 (granular `onParagraph*` events) are subsumed. **Features needing WordApi 1.6** (`body.getTrackedChanges()` → `acceptAll()/rejectAll()`, cross-platform) must be **feature-detected and degraded gracefully**, not gate the whole add-in. Requirement sets are declared in the manifest **and** re-checked at runtime via `Office.context.requirements.isSetSupported('WordApi', '1.6')`.

---

## Repository Structure

```
mutatis/
├── manifest/
│   ├── manifest.xml              # unified/legacy manifest (or manifest.json)
│   └── manifest.dev.xml          # localhost sideload variant
├── public/                       # static assets (icons, fallback html)
├── src/
│   ├── taskpane/
│   │   ├── index.html
│   │   ├── main.tsx              # mounts Preact app, Office.onReady gate
│   │   ├── shell/                # tab nav, layout, settings, theming, preview region
│   │   └── components/           # shared presentational components (FindingRow…)
│   ├── features/                 # one folder per feature (UI + logic)
│   │   ├── defined-terms/
│   │   ├── cross-reference/
│   │   ├── factual-consistency/
│   │   ├── interactive-lookup/
│   │   ├── repair/               # uses OOXML Service
│   │   ├── numbering/
│   │   ├── content/              # TOC / clause-library / templates
│   │   └── compare/              # v1 Word-vs-Word; v2 PDF
│   ├── shared/                   # shared infra — NO feature-specific logic
│   │   ├── doc-model/            # Document Model / Paragraph Index
│   │   ├── section-index/        # clause/schedule/recital structure builder
│   │   ├── ooxml/                # OOXML Service
│   │   ├── storage/              # storage abstraction + adapters
│   │   ├── config/              # Localisation/Config layer (see CONVENTIONS.md)
│   │   ├── findings/            # Findings model
│   │   ├── office/              # thin Office JS wrappers, batching, Capabilities
│   │   └── types/
│   ├── workers/                  # web workers (diff, zip) — optional, perf
│   └── styles/                   # global CSS tokens, Office theme bridge
├── tests/
│   ├── unit/  fixtures/  integration/
├── spec/                         # this specification
├── package.json  tsconfig.json  vite.config.ts  README.md
```

**Hard rule:** `src/shared/**` must not import from `src/features/**`.

---

## Office JS Constraints Reference

Fact-checked against official Microsoft docs. **Authoritative — feature docs reference this table instead of restating these facts. Do not "rediscover" them.**

| # | Constraint | Implication for builders |
|---|---|---|
| 1 | `document.body.footnotes` available (**WordApi 1.5**); each footnote has `.body`. | Proofing/cross-ref must scan footnote bodies, not just main body. **Endnotes** are a separate collection (`body.endnotes`). |
| 2 | `paragraph.listItem.listString` available (**WordApi 1.3**) → rendered number (`"4.2.1"`, `"(b)"`). | Primary signal for Section Index, cross-refs, numbering. Guard: not all paragraphs are list items (`getOrNullObject`). |
| 3 | `body.shapes` (`Word.ShapeCollection`) is **desktop-only** (`WordApiDesktop` 1.2). | Out of scope. Do not depend on it. |
| 4 | `body.getTrackedChanges()` → `TrackedChangeCollection`; `.acceptAll()/.rejectAll()` in **WordApi 1.6, cross-platform**. | Tracked-changes feature gated on 1.6; feature-detect — but it is **NOT** desktop-only. Per-author accept is **not** exposed (bulk only). |
| 5 | `Range.insertField(InsertLocation, FieldType.toc)` inserts a native TOC (`FieldType` enum, **WordApi 1.5**). | Use native field; don't synthesise text. Switches set via `field.code`; refresh via `field.updateResult()`. |
| 6 | `clearOrResetDirectFormat` **DOES NOT EXIST**. | To clear inline overrides: iterate runs/paragraphs and reset `font`/`paragraphFormat`, and/or apply a named style. |
| 7 | **No `document.onChanged` event.** | Use `onParagraphChanged` / `onParagraphAdded` / `onParagraphDeleted` (1.4) + `DocumentSelectionChanged` (1.1). |
| 8 | `body.search()` has a **255-char limit** per query (throws `InvalidArgument` beyond). | Chunk queries or operate over the cached Document Model text; never pass long/compound strings. |
| 9 | `SectionCollection` is **read-only for structure**; section breaks cannot be deleted via the typed API. | Use the OOXML loop to delete/merge section breaks. |
| 10 | **Cross-document reading impossible** except base64 workarounds (`createDocument`, `retrieveStylesFromBase64`). | Compare/template features use base64 round-trips / File-API upload; cannot read other open/closed files. |
| 11 | `insertFileFromBase64(..., replace)` can leave artifacts (stray leading paragraph / blank page). | OOXML Service offers `createDocument().open()` fallback. |
| 12 | Only **one `getFileAsync` document in memory** at a time (effective two-doc limit overall). | Serialise OOXML reads; `closeAsync()` promptly. Second compare doc → File-API upload, not parallel `getFileAsync`. |
| 13 | Always **batch `ctx.sync()`**; never sync inside per-item loops. | Single-pass loads; bounded sync count; chunked writes. |
| 14 | `DocumentSelectionChanged` fires on **mouse-up + every keystroke** (no document click event; fires on selection settle, not mouse-down). | **Debounce** (~400 ms) selection-driven work; round-trip ~200–600 ms in Word Online. |
| 15 | **Repeating Section Content Control is NOT exposed** in Office JS. | Templates use fixed slots; no OOXML cloning. |
| 16 | Task panes **cannot initiate HTML5 drag-and-drop** into the canvas. | Click-to-insert only. |
| 17 | Word's **native Compare is not in Office JS**. | Author redline at the OOXML level (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`). |

---

## Cross-Platform Matrix

| Capability | Word Win (WebView2) | Word Mac (WKWebView) | Word Online | Word Android | Notes / Degradation |
|---|---|---|---|---|---|
| Core typed API (paras, styles, ranges, lists) | ✅ | ✅ | ✅ | ✅ | Baseline; WordApi ≤1.5. |
| `listItem.listString` (1.3) | ✅ | ✅ | ✅ | ✅ | Load-bearing; available everywhere. |
| Footnote bodies (1.5) | ✅ | ✅ | ✅ | ✅ | Proofing scans footnotes. |
| TOC field insert (1.5) | ✅ | ✅ | ✅ | ✅ | TOC feature; fallback to `insertOoxml` `w:fldSimple` + F9 if 1.5 absent. |
| Granular `onParagraph*` (1.4) | ✅ | ✅ | ✅ | ✅ | Drives model invalidation. |
| `DocumentSelectionChanged` (1.1) | ✅ | ✅ | ✅ | ✅ | Selection context (debounced). |
| Tracked-changes accept/reject all (1.6) | ✅ | ✅ | ✅ | ✅ (if 1.6) | **Cross-platform in 1.6** — feature-detect; hide UI if unsupported. |
| OOXML loop (`getFileAsync` → `insertFileFromBase64`) | ✅ | ✅ | ✅ | ⚠️ verify | Repair, Compare. Validate slice behaviour on Online/Android; degrade Repair to "unsupported" with a clear message if constrained. |
| Numbering `setLevelNumbering`/`setLevelIndents` | ✅ | ✅ | ⚠️ partial | ⚠️ | **Word Online has incomplete support** for compound formatStrings/indents — best-effort apply, non-blocking banner, per-level error tolerance. |
| `body.shapes` (`WordApiDesktop` 1.2) | ✅ desktop | ✅ desktop | ❌ | ❌ | Out of scope. |
| Cross-document reading | ❌ | ❌ | ❌ | ❌ | Impossible except base64 workarounds. |

**Graceful degradation policy.** The Capabilities module (`shared/office`) is the single source of truth. Each feature declares a minimum requirement set; the shell **hides or disables** unsupported features with an explanatory tooltip rather than throwing. Never assume desktop-only APIs.

---

## Performance Guidelines

Large legal documents (200–400 pages, deeply nested numbering, many cross-references) are the design point.

- **Single-pass loading.** Features must **not** independently traverse the document — they query the cached Document Model. One traversal feeds all checkers per analysis run.
- **Batch `context.sync()`.** Queue many `ctx.load`/property sets and sync **once** (or a small bounded number of times). Never `await ctx.sync()` inside a per-item loop. Chunk large writes in windows (e.g. 200 paragraphs per sync).
- **Debounce reactive work.** Coalesce `onParagraph*` (400–600 ms) and `DocumentSelectionChanged` (~400 ms) before recomputing. Don't rebuild the Section Index on every keystroke. Pre-build dictionaries/indexes at open; lookups are then in-memory and instant.
- **255-char `body.search()` limit.** Operate over the Document Model's already-loaded text; reserve `search()` for short literal tokens (navigation re-location).
- **Two-documents-in-memory `getFileAsync` limit.** Serialise OOXML reads; for a second compare document use File-API upload (streams from disk) and free each `ArrayBuffer` after parsing.
- **Large-document handling.** Chunk/yield CPU work (process N paragraphs, yield to the event loop); show progress and support cancellation; hash-first / lazy-detail (compute alignment before running per-paragraph word diffs); virtualise long lists; **lazy-load heavy deps** (PDF.js, diff-match-patch, future Pyodide are code-split). Consider **web workers** for diff and zip to keep the UI thread responsive.
