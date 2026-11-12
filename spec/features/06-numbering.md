# 06 — Numbering (Numbering tab)

> **Status:** ✅ v1. Author/owner of the numbering-scheme JSON (the preset bodies below are the authoritative data; their SG/UK rationale lives in [Numbering Presets](../CONVENTIONS.md#numbering-presets)).
>
> Heading scheme: `#` doc title, `##` numbered top sections, `###` sub-sections.

---

## 1. Purpose & user value

Legal documents live or die by their numbering. A commercial agreement, a deed, or a court bundle relies on a precise, predictable hierarchy — clause `1`, sub-clause `1.1`, sub-sub-clause `1.1.1`, then lettered limbs `(a)`, romanette limbs `(i)`, and occasionally `(A)` — so that cross-references ("subject to clause 4.2(b)(iii)") remain unambiguous. Microsoft Word's built-in multilevel list UI is notoriously fragile for lawyers: it is buried in the ribbon, list templates silently fork, "Restart numbering" misfires, and firm precedents accumulate broken or inconsistent schemes.

This feature gives the drafter a **first-class numbering surface inside the task pane**:

- **Apply a known-good legal numbering scheme** (SG/UK Commercial, Schedule, Recitals, plain outline) to the selection or the whole document in one click, instead of hand-tuning the Define New Multilevel List dialog.
- **Save, name, and re-use firm-specific schemes**, and **share them as JSON** with colleagues (no server, no OneDrive; nothing leaves the machine).
- **Promote / demote** a paragraph one level with a single button, with Word recalculating the surrounding numbers automatically.
- **Be warned, after any numbering change, where cross-references might now be wrong** — the add-in highlights candidate references (e.g. "clause 3.2(a)") for human review rather than silently (and dangerously) rewriting them.

The value proposition: the precision of a properly-defined multilevel list, the speed of a preset library, and the safety of an explicit human-review step — all client-side.

---

## 2. Scope

### 2.1 In scope ✅ v1

1. **Scheme management.** Build, edit, save, load, export, import, and apply named legal numbering schemes. Ship a curated SG/UK preset library (Commercial, Schedule, Recitals, Plain Outline).
2. **Apply scheme.** Apply a `NumberingScheme` to either the current selection's list(s) or the whole document body, configuring each level via `list.setLevelNumbering()` and `list.setLevelIndents()`.
3. **Promote / Demote.** Shift the selected paragraph(s) up or down one list level via `paragraph.listItem.level`, clamped to `0..8`, with surrounding numbering preserved by Word's auto-recalculation. For a non-list paragraph adjacent to a list, `attachToList()` brings it into the list at the chosen level.
4. **Warn + highlight cross-reference review.** After any apply or promote/demote operation, run `body.search()` for legal reference patterns and highlight the matches so the user can manually verify them.

### 2.2 Cut ❌

| Cut capability | Reason | Where it lives instead |
|---|---|---|
| **One-shot static-to-dynamic CONVERSION** ❌ (OOXML `abstractNum`/`num` injection in a single operation) | OOXML injection is fragile and risks corrupting the document, with no reliable undo. | **Still cut.** Detection is the static-numbering flag in [feature 02](02-cross-reference-and-static-numbering.md). **But** the *assisted, tracked-changes, human-driven* removal + re-apply path is now in scope as the **Numbering upgrade** (§11.1, [ADR-0023](../../docs/adr/0023-numbering-upgrade-assisted-static-removal.md), supersedes ADR-0006): the tool deletes typed prefixes as native Word revisions and the **user** selects which cleared paragraphs to re-number — no OOXML injection. |
| **Automatic cross-reference updating** ❌ after renumbering | Rewriting reference text programmatically is error-prone and can silently corrupt meaning. | Replaced by the lightweight **warn + highlight** review (§5.3 / §7.5). |
| **OneDrive / SharePoint sync** ❌ of schemes | Out of the client-side-only design; needs a server/identity. | Replaced by **JSON export / import** (§7.3). |
| **Table-cell numbering** ❌ | List numbering inside table cells behaves inconsistently and interacts badly with cell auto-fit. | Skipped: detect via `paragraph.parentTableCellOrNullObject` and exclude (§8.1). |
| **Style-linked list modification** ❌ | Modifying a list bound to a paragraph style (`numStyleLink`) can break firm-template style inheritance across the document. | Operate on **direct-format lists only**; detect a style-linked abstract numbering and **warn, do not modify** (§8.7). |

See [README → What was cut and why](../README.md#what-was-cut-and-why-) for the consolidated list.

---

## 3. Dependencies

- [Section Index Builder](../ARCHITECTURE.md#section-index-builder) — identifies document sections (Recitals, Operative, Schedules) so "apply to whole document" and preset suggestions can be section-aware (e.g. suggest the Recitals preset inside a recitals block), and supplies section boundaries for the multi-section restart warning (§8.4).
- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — resolves the selection to a stable set of paragraph references and ordinal positions; used to find target list(s) and iterate paragraphs for promote/demote.
- [Findings Model](../ARCHITECTURE.md#findings-model) — read-only consumer here: surfaces the **static-numbering flag** authored in [feature 02](02-cross-reference-and-static-numbering.md) so the apply flow can warn "selection contains static-numbered paragraphs; these will not be converted." This tab never writes that finding.
- [Storage Abstraction](../ARCHITECTURE.md#storage-abstraction) — save/load schemes to `Office.context.document.settings` (document-scoped) and `localStorage` (machine-scoped library), with the JSON export/import path layered on top.
- [Localisation / Config Layer](../ARCHITECTURE.md#localisation--config-layer) — supplies default-scheme selection per jurisdiction, label strings ("clause" vs "paragraph"), indent unit, and the cross-reference regex set; the SG/UK preset rationale and the closed-bracket-form rule live in [Numbering Presets](../CONVENTIONS.md#numbering-presets). Default jurisdiction config = SG/UK.
- **Word Online numbering degradation** — `setLevelNumbering` compound formatStrings and `setLevelIndents` are partially supported on Word Online; see [Cross-Platform Matrix](../ARCHITECTURE.md#cross-platform-matrix). The per-level error-tolerance behaviour specific to this feature is in §3.1.
- The `listItem.listString` (WordApi 1.3) signal and the absence of `document.onChanged` are documented in [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) rows 2 and 7.

### 3.1 Word JavaScript API surface used

| API | Use here |
|---|---|
| `Word.Body.lists` (`body.lists`) | Enumerate existing lists to find the list under the selection / to apply a scheme. |
| `Word.List` | The list object whose levels we configure. |
| `list.setLevelNumbering(level, listNumbering, formatString)` | Core call. `level` is `0..8`; `listNumbering` a `Word.ListNumbering` enum; `formatString` uses `%1`–`%9` tokens (e.g. `"%1.%2.%3"`) plus literal text. Builds compound legal numbers natively. |
| `list.setLevelIndents(level, textIndent, bulletNumberPictureIndent)` | Sets the hanging-indent geometry per level. |
| `list.setLevelAlignment(level, alignment)` *(where supported)* | Left-align number text (legal default). Best-effort. |
| `Word.ListNumbering` enum | `arabic`, `upperRoman`, `lowerRoman`, `upperLetter`, `lowerLetter`, `none` — the only styles we map level `listNumbering` to. |
| `Word.Paragraph.listItem` (`paragraph.listItemOrNullObject`) | Access a paragraph's list membership; null when not in a list. |
| `paragraph.listItem.level` (read/write) | Promote/Demote. Writing recomputes numbering within the list automatically. |
| `paragraph.listItem.listString` *(API 1.3)* | Read the rendered number (`"1.1.1"`, `"(a)"`) for display/preview and as the cross-reference scan baseline. |
| `paragraph.attachToList(listId, level)` | Bring a non-list paragraph into an existing list at a level (promote/demote of an adjacent non-list paragraph; re-homing list paragraphs into the correct list). |
| `paragraph.startNewList()` | Start a brand-new list from a paragraph (only on paragraphs that are NOT static-numbered — static text is never auto-converted, per the cut). |
| `paragraph.parentTableCellOrNullObject` | Detect table-cell paragraphs to skip (§8.1). |
| `Word.Body.search(query, options)` | Cross-reference highlight scan; `matchWildcards: true` for pattern queries. |
| `range.font.highlightColor` | Apply the review highlight to matched references. |

**Per-level error tolerance (Word Online).** On hosts with partial numbering support ([Cross-Platform Matrix](../ARCHITECTURE.md#cross-platform-matrix)), the add-in **still attempts** the apply (best-effort) and shows a non-blocking banner, but **never throws on a partially-applied level**: catch per-level errors, continue applying the remaining levels, and report which levels degraded. Host detection via `Office.context.platform` / `Office.context.requirements`.

---

## 4. Data model

### 4.1 `NumberingScheme` schema

```ts
/** A complete, named legal numbering scheme. Persisted as JSON. */
interface NumberingScheme {
  /** Stable UUID. Generated on create; preserved across export/import. */
  id: string;
  /** Human label shown in the library, e.g. "UK/SG Commercial". */
  name: string;
  /** "SG-UK" | "UK" | "SG" | "US" | "GENERIC" — drives default selection + label strings. */
  jurisdiction: string;
  /** Optional longer description for the library tooltip. */
  description?: string;
  /** True for shipped presets (read-only in UI; user must "duplicate to edit"). */
  builtIn?: boolean;
  /** Unit noun for cross-ref labelling: "clause" (operative) or "paragraph" (schedules). */
  unitLabel?: string;
  /** Schema version for forward-compat on import. */
  schemaVersion: number; // = 1
  /** Exactly the levels defined; index 0..8. Sparse not allowed — define contiguously from 0. */
  levels: NumberingLevel[];
}

interface NumberingLevel {
  /** 0-based Word list level (0..8). */
  levelIndex: number;
  /** Maps to Word.ListNumbering: "arabic" | "upperRoman" | "lowerRoman"
   *  | "upperLetter" | "lowerLetter" | "none". */
  listNumbering: string;
  /** Word formatString with %1..%9 tokens, e.g. "%1.%2." for compound,
   *  "(%4)" for a lettered limb. The token number is (levelIndex+1). */
  formatString: string;
  /** Literal text emitted BEFORE the number token (folded into formatString on apply). */
  prefix: string;
  /** Literal text emitted AFTER the number token (folded into formatString on apply). */
  suffix: string;
  /** Hanging indent geometry, in points. */
  indent: {
    /** Left text indent for the paragraph body. */
    textIndentPt: number;
    /** Indent of the number/bullet itself (hanging position). */
    numberIndentPt: number;
  };
  /** Display hint for the level editor: "decimal-compound" | "paren-lower-alpha"
   *  | "paren-lower-roman" | "paren-upper-alpha" | "plain". Cosmetic only. */
  labelStyle: string;
  /**
   * Restart THIS level's counter whenever the list re-enters a shallower level.
   * restartAfterLevel = N means "reset this level's count after any level <= N".
   * null = inherit Word default (restart under immediate parent). -1 = never restart.
   */
  restartAfterLevel: number | null;
}
```

**Notes on `formatString`.** Word's `setLevelNumbering` `formatString` interleaves literal characters and `%n` placeholders, where `%n` is the *rendered value of level n-1*. For a compound decimal at level index 2 (`1.1.1`), the format is `"%1.%2.%3."`. For a parenthesised lettered limb at level index 3, the format is `"(%4)"`; `prefix`/`suffix` are typically empty because the brackets are baked into the format. The apply algorithm (§5.1) composes the final string as `prefix + formatString + suffix` only when `prefix`/`suffix` are non-empty, otherwise uses `formatString` verbatim.

### 4.2 Default preset — UK/SG Commercial (the DEFAULT)

Hierarchy `1 → 1.1 → 1.1.1 → (a) → (i) → (A)`. Lower-roman and letters in **closed parentheses** — `(a)`, `(i)`, `(A)`, never `a.` or `a)`. Levels 6–8 fall back to deeper parenthesised forms for unusually deep drafting. (SG/UK rationale and the closed-bracket-form rule: [Numbering Presets](../CONVENTIONS.md#numbering-presets).)

```json
{
  "id": "preset-uk-sg-commercial",
  "name": "UK/SG Commercial",
  "jurisdiction": "SG-UK",
  "description": "Standard UK & Singapore commercial hierarchy: 1 / 1.1 / 1.1.1 / (a) / (i) / (A). Closed-bracket form. Default scheme.",
  "builtIn": true,
  "unitLabel": "clause",
  "schemaVersion": 1,
  "levels": [
    {
      "levelIndex": 0, "listNumbering": "arabic", "formatString": "%1.",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 0, "numberIndentPt": 0 },
      "labelStyle": "decimal-compound", "restartAfterLevel": -1
    },
    {
      "levelIndex": 1, "listNumbering": "arabic", "formatString": "%1.%2.",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 36, "numberIndentPt": 18 },
      "labelStyle": "decimal-compound", "restartAfterLevel": 0
    },
    {
      "levelIndex": 2, "listNumbering": "arabic", "formatString": "%1.%2.%3.",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 72, "numberIndentPt": 36 },
      "labelStyle": "decimal-compound", "restartAfterLevel": 1
    },
    {
      "levelIndex": 3, "listNumbering": "lowerLetter", "formatString": "(%4)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 108, "numberIndentPt": 72 },
      "labelStyle": "paren-lower-alpha", "restartAfterLevel": 2
    },
    {
      "levelIndex": 4, "listNumbering": "lowerRoman", "formatString": "(%5)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 144, "numberIndentPt": 108 },
      "labelStyle": "paren-lower-roman", "restartAfterLevel": 3
    },
    {
      "levelIndex": 5, "listNumbering": "upperLetter", "formatString": "(%6)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 180, "numberIndentPt": 144 },
      "labelStyle": "paren-upper-alpha", "restartAfterLevel": 4
    },
    {
      "levelIndex": 6, "listNumbering": "arabic", "formatString": "(%7)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 216, "numberIndentPt": 180 },
      "labelStyle": "plain", "restartAfterLevel": 5
    },
    {
      "levelIndex": 7, "listNumbering": "lowerLetter", "formatString": "(%8)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 252, "numberIndentPt": 216 },
      "labelStyle": "paren-lower-alpha", "restartAfterLevel": 6
    },
    {
      "levelIndex": 8, "listNumbering": "lowerRoman", "formatString": "(%9)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 288, "numberIndentPt": 252 },
      "labelStyle": "paren-lower-roman", "restartAfterLevel": 7
    }
  ]
}
```

> The top level uses `restartAfterLevel: -1` (clause numbers run continuously `1, 2, 3 …` through the document and never reset). Each deeper level resets under its immediate parent.

### 4.3 Preset — Schedule

Schedules use `1 / 1.1 / (a)`, unit **"paragraph"** (not "clause"), top level frequently `Part 1 / Part 2`. We model the common case where the body of a Schedule numbers `1 / 1.1 / (a)`; "Part" headings are typically a separate heading style, so level 0 is `1.` (paragraph) with a configurable variant that prefixes `Part`.

```json
{
  "id": "preset-schedule",
  "name": "Schedule (paragraph)",
  "jurisdiction": "SG-UK",
  "description": "Schedule numbering: 1 / 1.1 / (a). Unit = 'paragraph'. Top level often grouped under 'Part 1 / Part 2' headings.",
  "builtIn": true,
  "unitLabel": "paragraph",
  "schemaVersion": 1,
  "levels": [
    {
      "levelIndex": 0, "listNumbering": "arabic", "formatString": "%1.",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 0, "numberIndentPt": 0 },
      "labelStyle": "decimal-compound", "restartAfterLevel": -1
    },
    {
      "levelIndex": 1, "listNumbering": "arabic", "formatString": "%1.%2.",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 36, "numberIndentPt": 18 },
      "labelStyle": "decimal-compound", "restartAfterLevel": 0
    },
    {
      "levelIndex": 2, "listNumbering": "lowerLetter", "formatString": "(%3)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 72, "numberIndentPt": 36 },
      "labelStyle": "paren-lower-alpha", "restartAfterLevel": 1
    },
    {
      "levelIndex": 3, "listNumbering": "lowerRoman", "formatString": "(%4)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 108, "numberIndentPt": 72 },
      "labelStyle": "paren-lower-roman", "restartAfterLevel": 2
    }
  ]
}
```

> A shipped **variant** `preset-schedule-part` sets level 0 to `formatString: "Part %1"`, `listNumbering: "arabic"` — for Schedules whose top grouping is literally "Part 1 / Part 2" as numbered list items rather than headings.

### 4.4 Preset — Recitals

Recitals are lettered in capitals in parentheses: `(A) (B) (C)`. Single-level.

```json
{
  "id": "preset-recitals",
  "name": "Recitals (A)(B)(C)",
  "jurisdiction": "SG-UK",
  "description": "Recital lettering: (A) (B) (C). Single level, parenthesised upper-letter.",
  "builtIn": true,
  "unitLabel": "recital",
  "schemaVersion": 1,
  "levels": [
    {
      "levelIndex": 0, "listNumbering": "upperLetter", "formatString": "(%1)",
      "prefix": "", "suffix": "",
      "indent": { "textIndentPt": 36, "numberIndentPt": 0 },
      "labelStyle": "paren-upper-alpha", "restartAfterLevel": -1
    }
  ]
}
```

### 4.5 Preset — Plain Outline (memo)

Plain decimal outline for memos / advisory notes: `1 / 1.1 / 1.1.1 / 1.1.1.1`, no parenthesised limbs. (Schema identical in shape; all levels `arabic`, compound `%1.%2…` formats.)

### 4.6 US litigation style — non-default, optional ❌-by-default

A **single optional** `preset-us-litigation` (`I. / A. / 1. / a.`) MAY be shipped but is **never** the default and is clearly labelled "US style — not for SG/UK drafting." It is suppressed entirely in SG/UK config unless the user enables "show non-SG/UK presets" — see [Numbering Presets](../CONVENTIONS.md#numbering-presets).

---

## 5. Algorithm

### 5.1 Apply a scheme to a list

**Input:** a resolved target `Word.List` (or a set of paragraphs to be brought into one list) + a `NumberingScheme`.

```
applyScheme(list, scheme, context):
  for each level in scheme.levels (sorted by levelIndex 0..8):
     enumValue   = mapToWordListNumbering(level.listNumbering)
     format      = composeFormat(level.prefix, level.formatString, level.suffix)
     try:
        list.setLevelNumbering(level.levelIndex, enumValue, format)
     catch perLevelError:
        record degraded(level.levelIndex); continue            // Word Online tolerance (§3.1)
     try:
        list.setLevelIndents(level.levelIndex,
                             level.indent.textIndentPt,
                             level.indent.numberIndentPt)
     catch: record degraded; continue
     if level.restartAfterLevel != null:
        applyRestart(list, level.levelIndex, level.restartAfterLevel)  // best-effort*
  context.sync()
  return { degradedLevels }
```

- `composeFormat`: when `prefix`/`suffix` are empty, return `formatString` verbatim; otherwise return `prefix + formatString + suffix`. Token integrity (`%1…%9`) must be preserved.
- `mapToWordListNumbering`: string → `Word.ListNumbering` member; `"none"` produces an unnumbered level (not in our default presets).
- **`applyRestart` caveat:** Office JS does not expose a fully general "restart this level after level N" control on `Word.List` in all builds. Where unavailable, restart behaviour falls to Word's default (restart under immediate parent), which already matches our presets for levels ≥ 1. `restartAfterLevel: -1` (continuous top-level clauses) is the only case requiring explicit handling; where the API cannot enforce continuity we **warn** that top-level clauses may restart per Word section (§8.4). Recorded as a known limitation (§10).

**Resolving the target list:**

1. From the selection, get paragraphs via the Document Model / Paragraph Index.
2. **Exclude** any paragraph where `parentTableCellOrNullObject` is non-null (§8.1) and any footnote-body paragraph (§8.2).
3. For each remaining paragraph, read `listItemOrNullObject`:
   - If already in a list → collect its `list`.
   - If **not** in a list and flagged static-numbered ([feature 02](02-cross-reference-and-static-numbering.md) finding) → **do not convert**; add to a "skipped (static)" report.
   - If not in a list and not static → it may be brought in via `startNewList()` (new list) or `attachToList(existingListId, level)` (merge into an adjacent scheme'd list).
4. If the selection spans multiple distinct lists, apply the scheme to each, and warn that multiple independent lists were affected.

### 5.2 Promote / Demote

```
changeLevel(paragraphs, delta /* -1 promote, +1 demote */, context):
  for each p in paragraphs:
     if p.parentTableCellOrNullObject is non-null: skip (table)        // §8.1
     li = p.listItemOrNullObject
     if li is null:
        // non-list paragraph adjacent to a list → attach at target level
        targetList = findAdjacentListOrNull(p)
        if targetList:
           newLevel = clamp((inferLevelFromNeighbour(p) ?? 0) + delta, 0, 8)
           p.attachToList(targetList.id, newLevel)
        else:
           skip (not in a list; nothing to promote/demote)
     else:
        newLevel = clamp(li.level + delta, 0, 8)
        li.level = newLevel        // Word auto-recalculates numbering in the list
  context.sync()
```

- **Clamping:** target level always clamped to `0..8`. Promote at level 0 and demote at level 8 are **no-ops** (button disabled in UI when all selected items are at the boundary).
- **Tracked changes:** writing `li.level` and `attachToList` generate tracked revisions when Track Changes is on. The op proceeds but the UI **warns** before running (§8.6).
- **Auto-recalculation:** changing `listItem.level` makes Word recompute every affected number in the list — no manual renumber pass needed.

### 5.3 Warn + highlight cross-reference review (post-op)

After **any** apply or promote/demote `context.sync()`:

```
reviewCrossReferences(body, config, context):
  clearPriorReviewHighlights(body)           // remove our own previous highlight colour
  patterns = config.crossRefPatterns         // SG/UK reference regexes (wildcards)
  for each pattern in patterns:
     results = body.search(pattern, { matchWildcards: true })
     context.load(results, 'items')
     await context.sync()
     for each r in results.items:
        r.font.highlightColor = config.reviewHighlightColour   // e.g. "#FFF2A8"
  context.sync()
  showReviewPanel(count, "Numbering changed — review highlighted references")
```

- We **do not** edit reference text; we only highlight candidates.
- The review panel lists each highlighted reference (with surrounding text), with a **"Clear highlights"** button that removes only our review colour (tracked by a sentinel so we never clear a user's own highlighting).
- **SG/UK reference patterns** (illustrative; supplied by Config — see [Numbering Presets](../CONVENTIONS.md#numbering-presets) for the cross-reference regex vocabulary):
  - `clause [0-9]@.[0-9]@` (compound clause refs, e.g. "clause 4.2", "clause 12.3.1")
  - `clause [0-9]@\([a-z]@\)` ("clause 4(a)")
  - `paragraph [0-9]@` (Schedule refs)
  - `\([a-z]\)` / `\([ivx]@\)` standalone limb refs ("see (b) above")
  - `Schedule [0-9]` / `Part [0-9]`
  - `recital \([A-Z]\)`

  (`@` = "1 or more of the preceding" in Word wildcard syntax.)

---

## 6. SG/UK config keys used

The feature reads these from the central config; the **values and rationale** live in [Numbering Presets](../CONVENTIONS.md#numbering-presets) and [Terminology](../CONVENTIONS.md#terminology):

- **Default scheme** = `preset-uk-sg-commercial` (§4.2); pre-selected on first run under SG/UK config.
- **Unit labels:** operative = "clause", Schedule bodies = "paragraph", recitals = "recital" — feed both the UI and the cross-reference patterns.
- **Closed-bracket form** `(a)/(i)/(A)` is mandatory in presets — never `a.`, `a)`, `i)`.
- **Cross-reference regex set** (§5.3), indent unit (pt/cm), and label strings.
- **US litigation outline** suppressed by default in SG/UK config (§4.6).

Singapore == England & Wales for these purposes (`jurisdiction: "SG-UK"`); no separate SG-only hierarchy ships. See [Jurisdiction Config Layer](../CONVENTIONS.md#jurisdiction-config-layer).

---

## 7. UI / task-pane behaviour

### 7.1 Layout (Numbering tab)

```
┌─ Numbering ─────────────────────────────┐
│ [ Scheme library ▾ ]   (search box)     │
│  • UK/SG Commercial   ★default  (builtIn)│
│  • Schedule (paragraph)         (builtIn)│
│  • Recitals (A)(B)(C)           (builtIn)│
│  • Plain Outline                (builtIn)│
│  • <firm custom schemes…>                │
│  [ New ] [ Duplicate ] [ Import ] [Export]│
├─────────────────────────────────────────┤
│ Level editor (selected scheme)           │
│  L0  arabic   %1.        ▢ preview: 1.    │
│  L1  arabic   %1.%2.     ▢ preview: 1.1   │
│  L2  arabic   %1.%2.%3.  ▢ preview: 1.1.1 │
│  L3  l.letter (%4)       ▢ preview: (a)   │
│  L4  l.roman  (%5)       ▢ preview: (i)   │
│  L5  u.letter (%6)       ▢ preview: (A)   │
│  [ + add level ] [ save ]                 │
├─────────────────────────────────────────┤
│ Apply:  ( ) selection  ( ) whole document │
│         [ Apply scheme ]                   │
├─────────────────────────────────────────┤
│ This paragraph:   [ ◄ Promote ] [ Demote ► ]│
│   current level: 1.1  →                    │
├─────────────────────────────────────────┤
│ ⚠ Review references (after apply)          │
│   12 references highlighted [ View ] [Clear]│
└─────────────────────────────────────────┘
```

### 7.2 Build / edit a scheme — the level editor

- A scheme is edited as a list of up to 9 **level rows**. Built-in presets are read-only; the user clicks **"Duplicate to edit"** to create an editable copy (new `id`, `builtIn: false`, name suffixed " (copy)").
- Each level row exposes: **Number style** dropdown (`Word.ListNumbering`); **Format** field with a **token helper** (insert `%1…%9`, brackets, dots) and live validation (tokens must reference defined levels; brackets must balance); **Prefix / Suffix** fields; **Indents** (text / number) in points (or cm per Config locale); **Restart after level** selector; and a **live preview** column rendering what level N looks like in context (`1.1.1`, `(a)`, `(i)`), computed without touching the document.
- Validation gates Save: contiguous levels from 0, valid tokens, non-empty name. Save enabled only when valid and editable.

### 7.3 Save + JSON export / import (firm sharing)

- **Save** writes the scheme via the Storage abstraction: **machine library** → `localStorage`; **document-pinned** (optional) → `Office.context.document.settings` (≤ ~2 MB budget shared with other features; schemes are small JSON). Library keyed by `id`; saving an existing `id` updates in place. Built-in preset `id`s are reserved (a duplicate is forced).
- **Export:** "Export scheme" (or "Export all") serialises to a downloadable `.json` (`schemaVersion` + scheme objects). No document content is included — only numbering definitions.
- **Import:** "Import scheme" reads a `.json` (file picker, client-side `FileReader`), validates `schemaVersion` and structure, and: assigns a fresh `id` if the incoming `id` collides with a built-in, or offers "replace / keep both" on a user-library collision; rejects malformed JSON with a clear error and never partially imports.
- This is the **only** sharing mechanism in v1 (replaces OneDrive sync): a firm distributes its house style as a `.json` over its normal channels; the add-in never transmits anything.

### 7.4 Apply scope
- **Apply to selection** — affects the list(s) intersecting the current selection (§5.1 resolution).
- **Apply to whole document** — iterates `body.lists`, optionally section-scoped via the Section Index (e.g. "apply Recitals preset only within the Recitals section").
- Apply is followed automatically by the cross-reference review (§5.3). Result toast, e.g. "Applied UK/SG Commercial to 1 list; 2 static paragraphs skipped; 0 degraded levels".
- **Promote / Demote** act on the current selection; buttons disabled at level boundaries; current rendered level (`listItem.listString`) shown live.

### 7.5 Post-apply cross-reference highlight review
- After any apply/promote/demote, the **Review references** panel populates with the count and a **View** list (each entry: highlighted text + a "go to" that selects the range). **Clear** removes only our review highlight colour. Non-blocking; persists until cleared or the next op refreshes it.

### 7.6 Banners / warnings (non-blocking)
- Word Online limited-support banner ([Cross-Platform Matrix](../ARCHITECTURE.md#cross-platform-matrix)).
- Track Changes on → "Level changes will be recorded as revisions."
- Style-linked list detected → "This list is linked to a paragraph style; it was not modified to avoid breaking your template." (§8.7)
- Static-numbered paragraphs in selection → "N paragraphs look like manually typed numbers; they were not converted (see the [feature 02](02-cross-reference-and-static-numbering.md) static-numbering flag)."

---

## 8. Edge cases

### 8.1 Tables (skip)
For every candidate paragraph, load `parentTableCellOrNullObject`; if non-null, **exclude** from apply and promote/demote. List numbering in table cells is unreliable and interacts badly with cell auto-fit. The result toast reports the count of skipped table paragraphs.

### 8.2 Footnotes (out of scope)
Footnote content lives in a separate body (`Word.NoteItem` body). v1 operates only on the main document body and does **not** number footnote content. If the selection is inside a footnote, apply/promote actions are disabled with a tooltip.

### 8.3 Mixed manual/automatic numbering
A selection may mix true list paragraphs with hand-typed "numbers." Per the cut, we **never auto-convert static text**. For already-list paragraphs belonging to a foreign list that should join the scheme'd list, use `attachToList(targetListId, level)` to **merge** them (re-homing list paragraphs, not converting static text). Static paragraphs are flagged ([feature 02](02-cross-reference-and-static-numbering.md)) and reported as skipped.

### 8.4 Multi-section documents
Word lists can restart at section boundaries, and per-section list restart is not exposed by Office JS. For a document split into Word sections (front matter / operative / schedules), top-level continuity (`restartAfterLevel: -1`) cannot be guaranteed across sections. We **warn**: "This document has multiple sections; top-level clause numbering may restart at a section break — verify manually." The [Section Index Builder](../ARCHITECTURE.md#section-index-builder) supplies the boundaries used to detect this.

### 8.5 Restart-after-level limits
Where `applyRestart` (§5.1) cannot enforce the requested restart/continuity directly, Word's default (restart under immediate parent) is used and, for the continuous-top-level case, the §8.4 warning covers the residual risk.

### 8.6 Tracked changes
Changing `listItem.level` or calling `attachToList` while Track Changes is on produces revisions. We detect the track-changes state and **warn before running** (§7.6); we never silently disable Track Changes.

### 8.7 Style-linked lists (detect + warn, do NOT modify)
If the target list's abstract numbering is **linked to a paragraph style** (`numStyleLink` / list-style-linked), modifying its levels would propagate to every paragraph using that style and can break firm-template inheritance. v1 **detects** this and **warns**, and **does not modify** the list. We operate on **direct-format lists only**. (Detection is best-effort given Office JS exposure; where the API cannot confirm direct-vs-linked, we apply conservatively and surface the warning.)

### 8.8 Word Online partial application
Per §3.1, per-level failures are caught and reported; the scheme is applied as far as the host allows, and the banner explains the limitation.

---

## 9. Test cases / acceptance criteria

| # | Scenario | Steps | Expected |
|---|---|---|---|
| T1 | **Apply UK/SG Commercial** | New doc; type 6 nested list paragraphs; select; Apply `preset-uk-sg-commercial` to selection. | Rendered numbers: `1`, `1.1`, `1.1.1`, `(a)`, `(i)`, `(A)`. Closed brackets, lower-roman/letters parenthesised. `listItem.listString` confirms each. |
| T2 | **Compound + restart** | Apply Commercial to `1 / 1.1 / 1.2 / 2 / 2.1`. | Level 1 restarts under each level-0 (`2.1` not `2.3`); top level continuous (`1`, `2`). |
| T3 | **Demote (a) to 1.1 level** | Cursor on a `(a)` item; promote it (level 3→2): a `(a)` promoted to the `1.1` style level renders as a compound decimal and the list **renumbers**. | Item changes to `1.1.x` style; following items recompute; no manual renumber needed. |
| T4 | **Promote/Demote clamping** | Item at level 0 → Promote disabled / no-op; item at level 8 → Demote disabled / no-op. | Buttons disabled at boundaries; no error thrown. |
| T5 | **Export scheme JSON** | Duplicate Commercial, rename "Firm House Style", edit L3 suffix; Export. | Downloaded `.json` round-trips: valid `schemaVersion`, levels intact, no document content present. |
| T6 | **Import scheme JSON** | On a clean machine/profile, Import the T5 `.json`. | Scheme appears in library with a fresh/clean `id`; applying reproduces the edited L3. Malformed file → rejected with clear error, nothing imported. |
| T7 | **Word Online degradation** | Open same doc in Word Online; apply Commercial. | Best-effort applied; limited-support banner shown; any degraded levels reported; no unhandled exception. |
| T8 | **Table skip** | Selection includes list paragraphs inside a table cell. | Table-cell paragraphs excluded; toast reports skipped count; non-table paragraphs applied. |
| T9 | **Static not converted** | Selection includes hand-typed "1." / "(a)" paragraphs (feature-02 flagged). | Not converted; reported as skipped (static); no OOXML injection attempted. |
| T10 | **Cross-ref review** | Doc contains "clause 4.2(a)"; apply scheme. | Reference highlighted; review panel lists it; **Clear** removes only the review highlight. |
| T11 | **Style-linked warn** | Apply to a list whose numbering is style-linked. | List not modified; warning shown. |
| T12 | **Tracked changes warn** | Track Changes on; Demote an item. | Warning shown before op; revision created; numbering recomputed. |
| T13 | **Recitals preset** | Apply `preset-recitals` to a recitals block. | Renders `(A) (B) (C)`. |
| T14 | **Schedule preset** | Apply `preset-schedule`. | Renders `1 / 1.1 / (a)`; unit label "paragraph" used in cross-ref patterns. |

**Acceptance:** T1, T3, T5–T7, T10 are **must-pass** for v1 release.

---

## 10. Known limitations

1. **Word Online** — `setLevelNumbering` formatString tokens and `setLevelIndents` are incompletely supported; compound formats and per-level indents may degrade. Desktop is the supported surface; Online is best-effort with a banner ([Cross-Platform Matrix](../ARCHITECTURE.md#cross-platform-matrix)).
2. **No automatic cross-reference updating** — after renumbering, references are only **highlighted for review**, never rewritten. Updating reference text remains a manual step.
3. **No *one-shot* static-to-dynamic conversion** — OOXML `abstractNum`/`num` injection stays out of scope. Manually typed numbering is detected/flagged ([feature 02](02-cross-reference-and-static-numbering.md)); the **Numbering upgrade** (§11.1, [ADR-0023](../../docs/adr/0023-numbering-upgrade-assisted-static-removal.md)) adds an *assisted, tracked-changes, human-driven* removal + re-apply path instead — typed-API only, no injection.
4. **Style-linked lists are not modified** — to protect firm-template inheritance; detected and warned, not changed. Only direct-format lists are modified.
5. **Per-section list restart not exposed** — top-level continuity across Word section breaks cannot be guaranteed via the API; surfaced as a warning.
6. **`restartAfterLevel` enforcement is best-effort** — where Office JS lacks a direct restart control, Word's default restart-under-parent behaviour is used.
7. **No table-cell or footnote numbering** in v1.
8. **Storage budget** — schemes share the ~2 MB `Office.context.document.settings` budget with other features; the machine library uses `localStorage`. Large libraries should be shared via JSON export, not document-pinned en masse.

---

## 11. Build milestones

| Milestone | Deliverable |
|---|---|
| **M1 — Data model + presets** | `NumberingScheme` / `NumberingLevel` types; ship Commercial (default), Schedule, Recitals, Plain Outline presets in Config; JSON schema validation. |
| **M2 — Apply engine** | §5.1 apply over `Word.List` with `setLevelNumbering` + `setLevelIndents`; target-list resolution; table/footnote/static exclusion; per-level error tolerance for Word Online. |
| **M3 — Promote/Demote** | §5.2 `listItem.level` change with clamping + `attachToList` for adjacent non-list paragraphs; boundary disabling; tracked-changes warning. |
| **M4 — Scheme management UX** | Library list, level editor with live preview + token helper + validation, save to Storage abstraction (localStorage + document settings). |
| **M5 — JSON export/import** | Export one/all; import with validation, collision handling, schemaVersion gate. |
| **M6 — Cross-reference review** | §5.3 `body.search()` highlight scan with SG/UK patterns from Config; review panel with go-to + scoped Clear. |
| **M7 — Edge cases + warnings** | Section-break warning, style-linked detect+warn, Word Online banner, static-numbering (feature 02) integration. |
| **M8 — Test pass + hardening** | T1–T14 automated/manual; desktop (Win/Mac) + Word Online verification; must-pass gate (T1, T3, T5–T7, T10). |

### 11.1 Numbering upgrade — assisted static-number removal (ADR-0023)

> **Added after v1 (ADR-0023, supersedes ADR-0006).**
>
> A two-step, human-gated power-up that lets a user clear hand-typed clause numbers and re-impose proper numbering. The tool's automation is bounded to **non-destructive, natively-reversible** operations; the structural judgment (which paragraphs are a list, which are hanging) stays with the human. **One-shot OOXML `abstractNum`/`num` conversion remains cut** — removal is a typed-API text delete recorded as tracked changes; re-apply reuses the existing M2 engine. See [ADR-0023](../../docs/adr/0023-numbering-upgrade-assisted-static-removal.md) for the six binding design constraints and the rejected alternatives.

| Milestone | Deliverable |
|---|---|
| **M9 — Assisted removal (Step 1)** | Sidebar fix-up queue over the feature-02 static-numbering findings; batch-select (queue + live Word selection). Delete the leading number token + delimiter + trailing/leading whitespace **as tracked changes** (`changeTrackingMode = trackAll`, prior mode restored; step disabled if change tracking unavailable — never an untracked delete). **Batch guard:** only delete where a leading static-number token is actually detected (hanging paragraphs swept into a selection are skipped). **Whitespace auto-rule** per paragraph with a **single global override** (collapse-to-one-space / strip-to-flush / keep-one-tab) in the scheme/config store. Token selection anchored to paragraph start, spanning run-split prefixes, never matching a later in-body occurrence. |
| **M10 — Sequence anomaly surfacing** | Surface duplicate/gap typed-number anomalies with **both readings** (typo *or* deliberately reserved); never auto-resequence. User retypes, ignores, or converts (Step 2 dissolves the typo). |
| **M11 — Human-driven re-apply (Step 2)** | After the user accepts the tracked deletions and **selects** the cleared paragraphs they judge should carry numbering (excluding hanging paragraphs), run the existing **M2 apply engine** + the **M6 §5.3 cross-reference review**. No new structural decisions made by the tool. |
| **M12 — Upgrade test pass** | Tracked-change round-trip (accept-all-except-one); batch-guard skip of an un-numbered paragraph; each whitespace variant (double-space, tab, NBSP) and the global override; run-split prefix; leading-occurrence-only selection; sequence-anomaly dual-reading; Step 2 hanging-paragraph exclusion via selection. Desktop + Word Online (change-tracking availability gate). |
