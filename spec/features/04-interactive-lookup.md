# 04 — Interactive Reference Lookup (Review tab)

**Status:** ✅ v1 — read-only "feedback on demand" preview of the target of a selected cross-reference or defined term.

> When the user selects or clicks a cross-reference ("Clause 14.2", "paragraph 3 of Schedule 2") or a defined term ("Completion Date"), the task pane *instantly* shows a **preview of the target text** in the sidebar. **The cursor never moves.** A **"Jump to"** button is offered but navigation is opt-in, never automatic. This is the natural companion to the Cross-Reference checker (which *validates* references) — here we *explain* them on demand. Read-only; modifies nothing.

## 1. Purpose & user value

Legal documents are dense webs of internal cross-references and defined terms. To check what a reference *means*, a drafter normally has to stop typing, scroll (or Ctrl+F) to the target, read it, scroll back, and recover their place — the single most frequent micro-interruption in clause drafting, especially painful in long agreements where the target may be 80 pages away and scrolling risks losing the cursor.

Interactive Reference Lookup removes the round-trip: the drafter reads the preview in peripheral vision and keeps writing.

## 2. Scope

### 2.1 v1 behaviour (in scope) ✅

- Fires on **debounced document selection change** (~400 ms). Framed as **"feedback on demand"** — select a reference, get a preview — not live typing feedback.
- Classifies the current selection (or the word under a bare cursor) as **Cross-reference** ("Clause 4.2", "clause 4.2(a)", "Schedule 2", "paragraph 3 of Schedule 2", "Recital (B)", "the Recitals", "Annexure 1"), **Defined term** (present in the pre-built dictionary), or **Neither** (idle).
- **Resolves** cross-references against the shared [Section Index Builder](../ARCHITECTURE.md#section-index-builder) (including auto-numbered units via `listString`); compound Schedule references ("paragraph 3 of Schedule 2") to the **Schedule paragraph**, not a body clause "3"; defined terms against the in-memory dictionary built at open; and definition-by-cross-reference (follow one hop to preview the substantive clause as well as the definition sentence).
- **Renders** a preview panel: target label + scrollable target text + **Jump to** button.
- Handles graceful states: *no definition found*, *external reference*, *ambiguous — multiple targets*.
- **External-document references** are *detected* and surfaced ("Points to an external document: *[name]*"), never silently failed.

### 2.2 Explicit cut/deferred list 🔜❌

| Item | Status | Reason |
|---|---|---|
| **Split-screen / docked target view** | ❌ | **No Office JS API** to split the document view or scroll a second viewport (per [Cross-platform matrix](../ARCHITECTURE.md#cross-platform-matrix)). The sidebar is the only UI surface. Framed as deliberate sidebar-first UX. |
| **Cross-file resolution** (reading the target in another `.docx`) | ❌ | Office JS **cannot read other files** ([Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 10). We detect the *pattern* and show "external document"; never resolve content. |
| **Real-time, per-keystroke typing feedback** | ❌ | `DocumentSelectionChanged` fires on every keystroke and must be debounced ([Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 14); replaced by selection-triggered "feedback on demand." |
| Auto-number reconstruction | **NOT cut** ✅ | `listString` is available (WordApi 1.3) and resolves auto-numbered sections (Clause 5, Clause 10). |
| Editing / fixing references | ❌ here | The Cross-Reference checker's job (feature 02). This feature is read-only. |
| Definitions in linked/embedded sub-documents | ❌ | Same limitation as cross-file; treated as external. |

See [README.md](../README.md#feature-list) for the consolidated lists.

## 3. Dependencies

Read these first.

- [Section Index Builder](../ARCHITECTURE.md#section-index-builder) — **reused wholesale** for cross-reference resolution (Schedule/Recital/Annexure handling, `listString`-based numbering). This feature adds **no** section-parsing logic; it only *queries* the index. Authored by feature 02.
- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — normalised paragraph list with stable indices; source for the defined-terms dictionary build and for selection-window reads.
- [Localisation / Config layer](../ARCHITECTURE.md#localisation--config-layer) — jurisdiction profile (SG/UK), reference vocabulary, definition quote style (Clause 6).
- [Task Pane UI Shell](../ARCHITECTURE.md#task-pane-ui-shell) — the preview region this feature renders into, and the shell's debounced `DocumentSelectionChanged` subscription.
- [Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) — row 14: `DocumentSelectionChanged` fires on mouse-up + **every keystroke** (no document click event; settles, not mouse-down), round-trip ~200–600 ms in Word Online ⇒ **debounce required**. row 2: `listItem.listString` for auto-number resolution. row 10: cross-document reading impossible. This feature does **not** restate these.

### 3.1 Requirement-set floor

- **WordApi 1.1** — `addHandlerAsync(DocumentSelectionChanged)`, `getSelectedDataAsync` / `getSelection()`, `range.select()`.
- **WordApi 1.3** — `paragraph.listItem`, `listItem.listString`, `Range.paragraphs`, `Body.search`.

Manifest declares `MinVersion="1.3"`; at runtime feature-probe `isSetSupported("WordApi", "1.3")` and degrade (Clause 8) if absent.

## 4. Patterns: debounce, selection read, cancellation

### 4.1 SelectionChanged + debounce

Registered once at task-pane init (after `Office.onReady`):

```ts
Office.context.document.addHandlerAsync(
  Office.EventType.DocumentSelectionChanged,
  onSelectionChangedRaw,
  (res) => { if (res.status !== Office.AsyncResultStatus.Succeeded) logRegistrationFailure(res); }
);
```

**Debounce (~400 ms trailing):**

```ts
let debounceTimer: number | null = null;
const DEBOUNCE_MS = 400;   // config: lookup.debounceMs

function onSelectionChangedRaw() {
  if (debounceTimer !== null) clearTimeout(debounceTimer);
  debounceTimer = window.setTimeout(runLookup, DEBOUNCE_MS);
}
```

400 ms is chosen so a *settling* selection (clicked and stopped) triggers exactly one lookup, and continuous typing produces *zero* lookups until the user pauses. Exposed as `lookup.debounceMs` for tuning. (Why a debounce is mandatory: [Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 14.)

### 4.2 getSelection + word-boundary expansion

```ts
await Word.run(async (context) => {
  const sel = context.document.getSelection();
  sel.load("text, isEmpty");
  const expanded = sel.getRange(Word.RangeLocation.whole);
  …
});
```

- **Selected span** (e.g. highlighted "Clause 4.2"): use the selection text directly.
- **Bare cursor** (empty range, the "I clicked on a word" case): expand to the **word under the cursor**, then *grow* to capture a multi-token reference by inspecting the **enclosing sentence** (capped at ~240 chars, config `lookup.sentenceWindowChars`) so that "Clause", "4.2", "of", "Schedule", "2" — which the cursor sits inside *one* token of — can be reassembled by the classifier (Clause 5). We read a **bounded window** (current sentence), never the whole paragraph, to bound latency and avoid mis-classifying distant text.

### 4.3 Cancellation / abort (seq guard)

Because Word Online round-trips are slow (up to ~600 ms) and the user may select several references quickly, an in-flight `Word.run` can resolve *after* a newer selection started. We version every lookup and discard stale results:

```ts
let lookupSeq = 0;

async function runLookup() {
  const mySeq = ++lookupSeq;
  let result: LookupResult;
  try {
    result = await Word.run(async (context) => { /* read + resolve */ });
  } catch (e) {
    if (mySeq !== lookupSeq) return;       // superseded mid-flight — drop
    return renderError(e);
  }
  if (mySeq !== lookupSeq) return;         // a newer lookup started — discard
  renderPreview(result);
}
```

A **logical AbortController**: we cannot cancel the Office round-trip, but the pane only ever renders the *latest* selection's result. The debounce timer is cleared on each new raw event so superseded timers never fire.

## 5. Data model & algorithm

### 5.1 In-memory defined-terms dictionary

Built **once at task-pane open** by a full paragraph scan (reusing the Paragraph Index), then cached and incrementally invalidated.

```ts
interface DefinedTerm {
  termKey: string;            // normalised lookup key: lowercased, NFC, collapsed ws
  display: string;            // canonical surface form: "Completion Date"
  defRange: Word.Range;       // range of the *definition paragraph/sentence* (for Jump-to)
  previewText: string;        // extracted definition text, trimmed to ~600 chars
  defKind: "inline" | "by-cross-ref" | "external";
  forwardRef?: string;        // if defined by reference: raw cross-ref to follow one hop ("Clause 6.1")
  sourceParagraphIndex: number;
}

interface TermsDictionary {
  byKey: Map<string, DefinedTerm[]>;  // array → handles a term defined more than once
  builtAt: number;                    // timestamp / doc-revision token
  dirty: boolean;
}
```

**Build procedure (at open):**
1. Pull all paragraphs from the Paragraph Index.
2. Identify definition paragraphs by jurisdiction-configured convention (Clause 6): a term in **double-bold quotes** — `“**Completion Date**”` — usually followed by `means` / `has the meaning` / `shall mean`. Definitions may live in a dedicated "1. Definitions and Interpretation" clause *or* inline anywhere (`(the **"Purchaser"**)`).
3. For each, extract `display` (the quoted term), build `termKey`, capture `defRange` and `previewText`.
4. Detect **definition-by-cross-reference**: if the definition body contains a cross-reference pattern (Clause 5.3) and little else, set `defKind = "by-cross-ref"`, store `forwardRef`.
5. Detect **external definitions** ("has the meaning given in the **Facility Agreement**") → `defKind = "external"`.
6. Index into `byKey`. A term defined twice yields a multi-entry array → drives the *ambiguous* state (Clause 7).

**Why pre-built:** classification of "is this a defined term?" must be O(1) at selection time — never a second slow round-trip per selection. The dictionary is the index. (Mirrors the "pre-build dictionaries/indexes at open" rule in [Performance guidelines](../ARCHITECTURE.md#performance-guidelines).)

### 5.2 Section Index query descriptor

Cross-references resolve against the shared [Section Index Builder](../ARCHITECTURE.md#section-index-builder); we do **not** rebuild or re-specify it. We query it with:

```ts
interface RefDescriptor {
  kind: "clause" | "schedule" | "schedule-paragraph" | "recital" | "annexure";
  primary: string;      // "14.2" | "2" | "(B)"
  secondary?: string;   // for schedule-paragraph: the paragraph "3" *within* Schedule "2"
}
// SectionIndex.resolve(descriptor) → SectionHit[]   (0, 1, or many)
```

`SectionHit` (from the shared module) exposes `{ range, label, previewText, paragraphIndex, container }`, where `label` is the human label ("Clause 14.2") and `container` distinguishes body vs a specific Schedule (for the proximity heuristic, Clause 8.6).

### 5.3 Lookup algorithm

```
runLookup():
  mySeq = ++lookupSeq
  windowText, selectionText = readSelectionWindow()      // §4.2, bounded ~240 chars
  if superseded(mySeq): return
  candidate = pickCandidate(selectionText, windowText)   // exact selection wins; else word+sentence
  cls = classify(candidate)
  switch cls.type:
    CROSS_REF:        hits = SectionIndex.resolve(cls.descriptor)
                      result = fromSectionHits(hits)
    DEFINED_TERM:     entries = dictionary.byKey.get(cls.termKey)
                      result = fromTermEntries(entries)   // may follow one cross-ref hop
    EXTERNAL:         result = ExternalRef(cls.docName)
    NEITHER:          result = Idle
  if superseded(mySeq): return
  renderPreview(result)                                   // §7
```

**Classification regexes** (vocabulary from the Config layer, Clause 6; forms below are SG/UK defaults; matching is case-insensitive on the keyword, anchored to capture the *whole* reference even when the cursor sits on one token):

```ts
// Numbering token: 14, 14.2, 14.2(a), 14.2(a)(iii), 4(1)
const NUM = String.raw`\d+(?:\.\d+)*(?:\([a-z0-9ivxlcdm]+\))*`;

// --- Compound Schedule paragraph: MUST be tested FIRST ---
const RX_SCHED_PARA = new RegExp(
  String.raw`\b(?:para(?:graph)?s?\.?)\s+(${NUM})(?:\s+to\s+${NUM})?\s+of\s+(?:Sch(?:edule|\.)?)\s+(${NUM}|[IVXLC]+)\b`,
  "i"
);

// --- External reference (cross-file) — tested before plain cross-ref ---
const RX_EXTERNAL = new RegExp(
  String.raw`\b(?:Clause|Schedule|paragraph|Annexure?|Appendix)\b[^.]*?\bof\s+the\s+([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,4}\s+(?:Agreement|Deed|Letter|Schedule|Lease|Charge|Guarantee))\b`,
  "i"
);

// --- Plain cross-references (internal) ---
const RX_CLAUSE    = new RegExp(String.raw`\bClause\s+(${NUM})\b`, "i");      // SG/UK: "Clause", NOT "Section"
const RX_SCHEDULE  = new RegExp(String.raw`\bSchedule\s+(${NUM}|[IVXLC]+)\b`, "i");
const RX_ANNEXURE  = new RegExp(String.raw`\bAnnexures?\s+(${NUM}|[A-Z])\b`, "i");
const RX_RECITAL   = new RegExp(String.raw`\bRecital\s+\(?([A-Z])\)?\b`, "i"); // "Recital (B)"
const RX_PARA_ONLY = new RegExp(String.raw`\bpara(?:graph)?s?\.?\s+(${NUM})\b`, "i");
```

**Order of precedence (first match wins — the crux of compound handling):**
1. `RX_SCHED_PARA` → `kind: "schedule-paragraph"`, `secondary = para`, `primary = schedule no.` — ensures "paragraph 3 of Schedule 2" resolves to the Schedule-2 paragraph, never a body clause "3".
2. `RX_EXTERNAL` → `EXTERNAL`, capture the proper-noun document name.
3. `RX_CLAUSE` / `RX_SCHEDULE` / `RX_ANNEXURE` / `RX_RECITAL` → respective `kind`.
4. `RX_PARA_ONLY` → bare "paragraph 3": **context-dependent** — if the cursor's containing section is a Schedule, resolve within *that* Schedule; otherwise ambiguous-or-neither.
5. **Defined-term test** (only if no cross-ref matched): look up `dictionary.byKey.get(normalise(phrase))`. **Only triggers if the key exists** — this stops common words ("Business", ordinary "completion") from firing (Clause 8.7).
6. Else `NEITHER`.

**Resolution — compound Schedule reference** (`kind:"schedule-paragraph"`, `primary="2"`, `secondary="3"`):
1. `SectionIndex.resolve({kind:"schedule", primary:"2"})` → the Schedule 2 container hit.
2. Within that container, locate **paragraph 3** using the Schedule's *own* numbering — read `paragraph.listItem.listString` for each container paragraph and match `"3"`, **not** body clause numbering. The Section Index already partitions hits by `container`, so this is a filtered lookup, not a fresh scan.
3. Preview text = that paragraph's text (+ breadcrumb "Schedule 2 — paragraph 3").

This is the explicit reason the compound pattern is tested first and resolved container-scoped: a document will almost always also contain a body "Clause 3"/"paragraph 3", and naïve resolution would preview the wrong one.

**Resolution — defined term defined by cross-reference** (`defKind === "by-cross-ref"`):
1. Render the **definition sentence** as primary preview.
2. **Follow one hop**: classify `forwardRef` ("Clause 6.1") and `SectionIndex.resolve` it.
3. Append the resolved substantive clause as a secondary preview block, with its own Jump-to. We follow **exactly one** hop (no transitive chains) to bound latency and avoid loops.

### 5.4 Cache invalidation

- **Section Index**: invalidated by the shared module's own `onParagraph*` handling — we rely on it.
- **Defined-terms dictionary**: subscribe to the same `onParagraph*` events. On any event set `dictionary.dirty = true` (cheap) and schedule a **debounced rebuild** (~1 s trailing, config `lookup.dictRebuildDebounceMs`, separate from the lookup debounce). The next lookup after a rebuild uses the fresh dictionary; a lookup arriving while `dirty` uses the stale dictionary but is flagged so the pane shows a subtle "updating definitions…" hint rather than a wrong answer for a just-deleted term. Full rebuild cost is bounded (single paragraph scan) and runs off the lookup critical path (Clause 8.5).

## 6. SG/UK config keys used

Inherits the Cross-Reference and Defined-Terms conventions from the [Localisation / Config layer](../ARCHITECTURE.md#localisation--config-layer). Key jurisdiction-correct points (values owned by [Terminology](../CONVENTIONS.md#terminology) and [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions)):

- **"Clause" is the default unit**, *not* "Section"; **"Schedule"** not "Exhibit"; **"Recitals"** lettered "(A)", "(B)" (hence `RX_RECITAL` captures a parenthesised letter); **"paragraph X of Schedule Y"** is canonical (Schedules divide into paragraphs, reflected by the compound resolver Clause 5.3).
- **Defined terms** appear in **double curly quotes, bold** (`“**Completion Date**”`), introduced with *means* / *has the meaning given to it in* / *shall mean*. Both smart and straight quotes accepted via config.
- **No US conventions** (no "§", no "Section" default, no "Exhibit").

Feature-local config keys consumed (defaults shown):

```jsonc
{
  "lookup.debounceMs": 400,
  "lookup.dictRebuildDebounceMs": 1000,
  "lookup.sentenceWindowChars": 240,
  "lookup.previewMaxChars": 600,
  "definition.quoteStyle": "smart|straight",
  "definition.introVerbs": ["means", "shall mean", "has the meaning"]
}
```

(`vocab.clauseKeyword`, `vocab.scheduleKeyword`, jurisdiction profile etc. are shared keys read from the same layer; `NUM` can be profile-tuned for statutory `4(1)(a)` vs decimal `4.1(a)` numbering.)

## 7. UI / task-pane behaviour

Rendered into the [Task Pane UI Shell](../ARCHITECTURE.md#task-pane-ui-shell) preview region. This feature owns a single **Preview panel**.

### 7.1 Preview panel layout (resolved state)

```
┌────────────────────────────────────────────┐
│  ◀ Reference preview                        │  ← panel title (static)
├────────────────────────────────────────────┤
│  Clause 14.2                          [↗]   │  ← TARGET LABEL as heading + Jump-to icon
│  Indemnity                                  │  ← (optional) target heading/marginal note
├────────────────────────────────────────────┤
│  14.2  The Seller shall indemnify the       │  ← SCROLLABLE target text
│        Purchaser against all Losses …       │     (max-height, internal scroll;
│        … (continues)                        │      preserves list numbering via listString)
├────────────────────────────────────────────┤
│                              [ Jump to → ]  │  ← primary Jump-to button → range.select()
└────────────────────────────────────────────┘
```

- **Target label** = `SectionHit.label` / `DefinedTerm.display`, rendered as a heading.
- **Target text** scrollable within the panel (capped height, internal overflow), never expands the pane unbounded. Trimmed to `previewMaxChars` with a "…" affordance.
- **Jump to** calls `targetRange.select()` (on desktop, `select` carries scroll-into-view semantics), moving the *Word* cursor to the target. This is the **only** action that moves the cursor — explicit and user-initiated.
- For **defined-term-by-cross-ref** (Clause 5.3): two stacked blocks — *Definition* then *Resolved clause* — each with its own Jump-to.

### 7.2 Graceful states

| State | Trigger | Pane shows |
|---|---|---|
| **Idle** | selection classified NEITHER | "Select a clause reference or defined term to preview it." Last result optionally retained, dimmed. |
| **No definition found** | term-shaped phrase / cross-ref with **0** hits | "No definition found for *“Foo Bar”* in this document." + (if it *looks* like a term) "It may be defined in an external document or not yet defined." |
| **External reference** | `RX_EXTERNAL` match or `defKind === "external"` | "Points to an **external document**: *Facility Agreement*. This add-in can only preview references within the current file." No Jump-to. |
| **Ambiguous — multiple targets** | >1 Section hit, or multi-entry dictionary key | "**2 possible targets**" + a short pick-list (each with label, container badge — *Body* / *Schedule 2* — and Jump-to). Best guess (proximity, Clause 8.6) expanded first. |
| **Updating** | lookup arrives while `dictionary.dirty` | Subtle inline "updating definitions…" chip on the result, no blocking spinner. |

### 7.3 No automatic navigation

The panel **never** scrolls the document or moves the cursor on its own. Selection only ever updates the *sidebar*. This is the core invariant: read in the sidebar, keep writing.

### 7.4 Performance feel

- **Debounce** 400 ms → a deliberate selection yields one lookup; typing yields none.
- **Latency budget:** debounce (400 ms) + Word round-trip (~200–600 ms Online, less on desktop) ⇒ preview appears **~0.6–1.0 s** after settling Online; snappier on desktop. A lightweight skeleton/shimmer shows between lookup start and render.
- Dictionary/Section lookups themselves are **in-memory and instant**; the only latency is the Office read round-trip and the debounce. No second round-trip is incurred to *find* a definition (that is why Clause 5.1 pre-builds).
- The **cancellation pattern (Clause 4.3)** guarantees no flicker of stale previews when moving quickly between references.

## 8. Edge cases & compatibility

### 8.1 Document with no definitions section
The dictionary builder does **not** assume a Definitions clause — it scans **all** paragraphs for the double-bold-quote convention. Zero definitions ⇒ defined-term lookup is inert (every term → "no definition found"), and cross-reference lookup still works fully. Graceful degradation, never an error.

### 8.2 Cross-file references
Detected via `RX_EXTERNAL` and `defKind === "external"`. We **cannot** read the other file ([Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 10) — surface the external-document message with the captured name. No Jump-to.

### 8.3 Deeply nested numbering
`NUM` captures `14.2(a)(iii)` and statutory `4(1)(a)`. The Section Index resolves to the deepest matching auto-numbered paragraph via `listString`; if only a parent ("14.2") exists, preview the parent and note "deepest available: 14.2".

### 8.4 Auto-numbered targets
Resolution depends on `listString` being populated by Word. Where auto-numbered, `listString` gives "14.2" even though literal text is "The Seller shall…". Where numbering is **typed manually**, the Section Index falls back to text matching. Both paths are handled by the shared index; this feature inherits both (Clause 10 limitation).

### 8.5 Large documents (200+ pages)
Dictionary and Section Index are built **once at open** and cached; selection lookups never re-scan. Rebuilds on `onParagraph*` are debounced (~1 s) and run off the lookup critical path. Selection reads use a **bounded window** (~240 chars), so payload is independent of document length; lookup cost does **not** grow with document size (consistent with [Performance guidelines](../ARCHITECTURE.md#performance-guidelines)).

### 8.6 Ambiguous "Clause 2" / "paragraph 3" in body and multiple schedules
The Section Index returns multiple `SectionHit`s with distinct `container`s. We apply a **proximity heuristic**:
1. Prefer the hit in the **same container** as the cursor (cursor inside Schedule 2 + "paragraph 3" → prefer Schedule 2's paragraph 3).
2. Else prefer the **nearest by paragraph index**.
3. If still tied / cross-container, present the **ambiguous pick-list** (Clause 7.2) with the proximity best-guess expanded first.
The cursor's container comes from the selection read (its paragraph index → Section Index container lookup).

### 8.7 Defined terms that are also common words
Defined-term lookup **only fires when the normalised phrase is a key in the dictionary** — and (for single-word common terms) when the occurrence carries the definitional context, i.e. we match the **defined surface form** (usually multi-word: "Business Day", "Completion Date"). A lone ordinary "completion" with no dictionary key → NEITHER → idle. We never invent a definition for a word not in the dictionary.

### 8.8 Compatibility / degradation
- If `WordApi 1.3` is unavailable (`isSetSupported` false): `listString` auto-numbering is lost; degrade the Section Index to text-matched numbering only and show a one-time banner "Some auto-numbered references may not resolve in this Word version." Defined-term lookup and 1.1 selection handling still work.
- If `addHandlerAsync` registration fails: feature disables itself gracefully with a pane message; the rest of the add-in is unaffected.

## 9. Test cases / acceptance criteria

Samples use SG/UK-correct documents (Clause/Schedule/Recital vocabulary, double-bold-quote definitions).

| # | Action | Expected |
|---|---|---|
| T1 | Click inside "**Clause 14.2**" in body text | Within ~1 s (Online): panel shows heading **"Clause 14.2"**, scrollable body of clause 14.2, working Jump-to. Cursor stays put. |
| T2 | Select "**paragraph 3 of Schedule 2**" | Panel shows **"Schedule 2 — paragraph 3"** and the text of *that Schedule paragraph* — **not** any body "Clause 3" or body "paragraph 3". |
| T3 | Click on "**Completion Date**" (defined term) | Panel shows **"Completion Date"** and the definition sentence. Jump-to navigates to the definition. |
| T4 | Click a term defined by cross-reference ("Completion Date" = "the date determined under Clause 6.1") | Panel shows the definition sentence **and** a second block previewing **Clause 6.1**, each with its own Jump-to. |
| T5 | Click "**Recital (B)**" | Panel previews Recital (B); Jump-to works. |
| T6 | Click "**Clause 4.2 of the Facility Agreement**" | **External-document** state; "Points to an external document: *Facility Agreement*." No Jump-to. |
| T7 | Select a common word "completion" *not* in the dictionary | Idle/NEITHER — **no** false definition preview. |
| T8 | "paragraph 3" referenced while cursor is inside Schedule 2 (also a body "Clause 3" exists) | Proximity heuristic resolves to **Schedule 2 paragraph 3** (same container), expanded first. |
| T9 | "Clause 2" matching body + two Schedules | **Ambiguous** pick-list with container badges; best-guess by proximity expanded. |
| T10 | Type continuously for 10 s without pausing | **Zero** lookups fire (debounce); no preview thrash. |
| T11 | Rapidly click 4 different references in 1 s | Only the **last** reference's preview renders; no stale flicker (cancellation Clause 4.3). |
| T12 | Delete the paragraph defining "Completion Date", then click another "Completion Date" use | After debounced rebuild, lookup returns "no definition found"; during the rebuild window an "updating definitions…" chip shows, never a wrong stale definition asserted as current. |
| T13 | Document with **no** Definitions clause but inline `(the **"Purchaser"**)` | "Purchaser" resolves to the inline definition. |
| T14 | WordApi 1.1-only host (no 1.3) | Banner about auto-numbering; manually-numbered references and defined terms still resolve. |
| T15 | 220-page SPA | Open builds dictionary+index once; subsequent lookups are debounce-latency-bound, not document-size-bound. |

**Acceptance:** T1–T6 and T10–T11 must pass on both Word desktop and Word Online; T7–T9, T12–T15 must pass on desktop and degrade per Clause 8 on Online.

## 10. Known limitations

1. **Sidebar-only — no split-screen.** ❌ No Office JS API to dock or split the document view. The preview lives in the task pane. *Deliberate UX stance*, not a defect: the value is reading **without leaving** the cursor.
2. **Selection-mouse-up latency.** Lookups are driven by `DocumentSelectionChanged` (fires on mouse-up/settle, not mouse-down; ~200–600 ms in Word Online). With the 400 ms debounce, previews appear ~0.6–1.0 s after settling Online. Intrinsic to the platform ([Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 14).
3. **Auto-number resolution depends on `listString` being populated by Word.** Manual hard-typed numbering, or a host where `listString` is empty/unsupported (pre-1.3), falls back to text matching and some deeply-nested auto-numbers may not resolve precisely (Clause 8.4, 8.8).
4. **One hop only** for definition-by-cross-reference; transitive chains ("A means B; B means C") are not followed past the first hop (latency/loop safety).
5. **External documents are never read** — only the reference *pattern* is detected and named ([Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 10).
6. **Classification is convention-driven.** Non-standard reference phrasing or unconventional definition styling may not classify; the feature stays silent (idle) rather than guessing.

## 11. Build milestones

| M | Milestone | Deliverable / exit criterion |
|---|---|---|
| **M0** | Plumbing | `addHandlerAsync(DocumentSelectionChanged)` registered; 400 ms debounce; selection read + bounded sentence window; cancellation/seq guard. Logs the classified candidate. |
| **M1** | Cross-reference preview (internal) | Query reused Section Index for Clause/Schedule/Recital/Annexure; render preview panel + Jump-to (`range.select()`). T1, T5 pass. |
| **M2** | Compound Schedule resolution | `RX_SCHED_PARA` precedence + container-scoped `listString` paragraph resolution. T2, T8 pass. |
| **M3** | Defined-terms dictionary | Build-at-open full scan, double-bold-quote detection, `byKey` index, preview extraction. T3, T13 pass. |
| **M4** | Definition-by-cross-ref (one hop) | `forwardRef` follow + stacked preview blocks. T4 pass. |
| **M5** | Cache invalidation | Subscribe to `onParagraph*`; debounced rebuild; `dirty` flag + "updating" chip. T12 pass. |
| **M6** | Ambiguity + proximity heuristic | Multi-hit pick-list, container badges, cursor-container proximity best-guess. T9 pass. |
| **M7** | External + graceful states | `RX_EXTERNAL`, `defKind:"external"`, no-definition/idle states, common-word suppression. T6, T7 pass. |
| **M8** | Hardening / compat | `isSetSupported` degradation banners, registration-failure fallback, large-doc perf pass, Word Online latency tuning. T10, T11, T14, T15 pass. |
| **M9** | Polish | Skeleton/shimmer, panel scroll caps, copy review (SG/UK wording), config exposure of debounce/window/preview limits. Full T-suite green per Clause 9 acceptance. |
