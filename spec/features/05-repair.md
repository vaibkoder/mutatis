# 05 — Document Repair (Repair tab)

> **Status:** ✅ v1. Primary consumer of the [OOXML Service](../ARCHITECTURE.md#the-ooxml-service).
>
> Heading scheme: `#` doc title, `##` numbered top sections, `###` sub-sections.

---

## 1. Purpose & user value

Legal documents are passed between firms, counterparties, courts and document management systems (DMS) repeatedly over the life of a matter. Each hop adds damage: conflicting style overrides pasted from precedents, blank-paragraph cruft, leftover comments, accepted/rejected-but-not-cleaned tracked-change colouring, `rsid` revision bloat that balloons file size, stray section breaks that cause blank pages, and — most sensitively — **metadata** (authors, last-saved-by, DMS custom XML, custom properties) that can leak privilege, client identity, or internal commentary.

The **Repair tab** gives a lawyer a single, controlled place to **stabilise a "problem document"** received from another firm and to **scrub confidential metadata** before sending one out. Two headline user values:

1. **Stabilisation** — turn an unstable, inconsistently-formatted, bloated document into a clean, predictable one that behaves when edited, paginated, or filed.
2. **Confidentiality scrubbing** — remove metadata, comments, and DMS-injected artefacts that could breach privilege or confidentiality obligations (SRA Code of Conduct in England & Wales; PDPA and professional-conduct duties in Singapore).

Repair runs entirely on the user's machine (see [Client-Side Model](../ARCHITECTURE.md#client-side-model)) — the differentiator versus cloud "document inspector" tools, and the reason metadata scrubbing can be trusted on privileged material. The UI states this **on-device-only** posture explicitly and repeatedly (§7).

**Non-destructive by default.** Repair never silently mutates a document. Every action is opt-in via checkbox, every destructive action requires explicit confirmation, and the user always sees a **repair report** of what was found before anything is changed.

---

## 2. Scope

### 2.1 In scope ✅ v1

1. **Strip metadata** — clear document properties (built-in + custom); remove all comments; delete DMS/iManage/NetDocuments custom XML parts; strip `rsid` revision IDs (via the OOXML loop).
2. **Clear conflicting style overrides / standardise styles** — reset direct (inline) formatting by manual iteration, then conform the document to a chosen **style profile** (typography, margins, spacing, indents), propagating into footnotes/endnotes.
3. **Normalise spacing + delete blank paragraphs** — normalise paragraph `spaceBefore`/`spaceAfter`/`lineSpacing` to the profile; delete empty paragraphs. **Line-spacing and margin normalisation are opt-in**, never silent (§7).
4. **Accept all tracked changes** — bulk accept, **gated behind an explicit, separately-confirmed user action** (never part of a "repair everything" sweep).
5. **Section-break detection + removal** and **rsid-bloat detection + stripping**, both via the OOXML loop. Section-break removal is **flag-and-confirm** because orientation breaks are legitimate (§7).

### 2.2 Cut ❌

- **Text box detection/removal** ❌ — removed from scope. Not in v1.
- **Per-author selective tracked-change management** ❌ — the typed API only supports bulk `acceptAll()`/`rejectAll()`; per-author accept/reject is **not available** (see [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 4) and **not** in scope.
- **PDF-to-Word structural repair** ❌ — not part of this product at all.

See [README → What was cut and why](../README.md#what-was-cut-and-why-) for the consolidated list.

---

## 3. Dependencies

- [The OOXML Service](../ARCHITECTURE.md#the-ooxml-service) — Repair is its **primary v1 consumer**; the read-modify-write loop, slice handling, base64, and the `insertFileFromBase64` / `createDocument().open()` fallback are all specified there. This doc keeps only the **feature-specific surgery** (§5.6 rsid strip, §5.7 section-break edits).
- [Findings Model](../ARCHITECTURE.md#findings-model) — the Repair Report's findings render through the shared Findings/UI shell.
- [Task Pane UI Shell](../ARCHITECTURE.md#task-pane-ui-shell) — hosts the Repair tab.
- [Localisation / Config Layer](../ARCHITECTURE.md#localisation--config-layer) carrying [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles) — the SG/UK court typography, spacing, margin and orientation-protection rules live there. This doc keeps the JSON profile **fixtures** but defers the **rules** to that anchor.
- Office JS facts used and **not** restated here — see [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference): `clearOrResetDirectFormat` does not exist (row 6), `SectionCollection` read-only for structure (row 9), tracked-changes accept/reject-all in WordApi 1.6 / per-author not exposed (row 4), `insertFileFromBase64` artifacts (row 11), batch `ctx.sync()` (row 13).

---

## 4. Data model

### 4.1 Style Profile schema

```jsonc
{
  "$schema": "mutatis/style-profile@1",
  "id": "sg-court-filing",
  "name": "SG Court Filing (PD 2021)",
  "jurisdiction": "SG",          // "SG" | "UK" | "GENERIC"
  "kind": "court" ,              // "court" | "commercial"
  "description": "Singapore Supreme Court / State Courts Practice Directions 2021.",
  "version": "1.0.0",
  "locked": true,                // shipped profiles are read-only; user clones to edit

  // Per-section page geometry. Values in points (Office JS unit). 1 mm = 2.83464567 pt.
  "page": {
    "margins": {                 // points
      "top": 99.21, "bottom": 99.21, "left": 99.21, "right": 99.21  // 35mm
    },
    "size": "A4",                // informational; we set pageHeight/pageWidth if enforced
    "orientation": "portrait",
    "pageNumber": { "position": "top-right", "enforced": false }     // we flag, do not auto-insert
  },

  // Default body paragraph formatting.
  "body": {
    "font": { "name": "Times New Roman", "size": 12, "color": "#000000",
              "bold": false, "italic": false, "underline": "None" },
    "spacing": {
      "lineSpacingRule": "multiple",   // "single"|"onePtFive"|"double"|"multiple"|"exactly"|"atLeast"
      "lineSpacing": 2.0,              // 2.0 == double when rule=multiple
      "spaceBefore": 0, "spaceAfter": 0   // points
    },
    "indent": { "firstLine": 0, "left": 0, "right": 0 }  // points
  },

  // Named-style overrides applied via getStyles(). Keyed by Word style name.
  "styles": {
    "Normal":   { "font": { "name": "Times New Roman", "size": 12 },
                  "paragraphFormat": { "lineSpacingRule": "multiple", "lineSpacing": 2.0,
                                       "spaceBefore": 0, "spaceAfter": 0 } },
    "Heading 1":{ "font": { "name": "Times New Roman", "size": 12, "bold": true } },
    "Quotation":{ "font": { "name": "Times New Roman", "size": 10 },     // TNR10 for quoted authorities
                  "paragraphFormat": { "lineSpacingRule": "onePtFive", "lineSpacing": 1.5,
                                       "left": 36, "right": 36 } },
    "Footnote Text": { "font": { "name": "Times New Roman", "size": 10 } }
  },

  // Special instructions.
  "rules": {
    "citationBracketStyle": "square",   // SG: paragraph cross-citation in [xx]
    "quotedAuthoritySize": 10,
    "normalisationOptIn": true,         // line-spacing/margin changes require explicit opt-in
    "protectOrientationSectionBreaks": true
  }
}
```

Field notes:
- All linear measurements are in **points** to match Office JS (`section.pageSetup.topMargin` etc. and `paragraphFormat` are points). Conversions from mm are pre-computed in shipped profiles; the loader also accepts a `"mm"` unit hint and converts.
- `lineSpacing: 2.0` with `lineSpacingRule: "multiple"` is the canonical encoding of **double** spacing.
- `styles` keys must be the exact Word style **display names** as returned by `getStyles()`.

The concrete shipped profiles (SG Court Filing, UK Court Filing, Firm Commercial) are JSON **fixtures** under §4.3; their **rules** (when normalisation defaults unchecked, orientation protection, etc.) are governed by [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles).

### 4.2 Repair Report model

Produced in two phases — **scan** (what was found) and **apply** (what was changed); same shape, different `phase`. Findings map onto the shared [Findings Model](../ARCHITECTURE.md#findings-model).

```jsonc
{
  "documentName": "Deed of Assignment (3rd markup).docx",
  "generatedAt": "2026-06-06T09:30:00+08:00",
  "profile": { "id": "sg-court-filing", "name": "SG Court Filing (PD 2021)" },
  "phase": "scan",                 // "scan" | "apply"
  "findings": [
    {
      "id": "metadata.author",
      "category": "metadata",      // metadata|comments|customXml|styles|spacing|blankParas|tracked|sectionBreaks|rsid
      "severity": "confidentiality", // info|stabilisation|confidentiality
      "label": "Document author / last-saved-by present",
      "detail": "author='Jane Tan', lastAuthor='DMS Service Account'",
      "count": 2,
      "destructive": false,
      "requiresConfirm": false,
      "selectedByDefault": true,
      "action": "clear",
      "result": null              // populated in apply phase: "done"|"skipped"|"error"
    },
    {
      "id": "tracked.changes",
      "category": "tracked",
      "severity": "stabilisation",
      "label": "Tracked changes present",
      "detail": "Document contains tracked insertions/deletions.",
      "count": 47,
      "destructive": true,
      "requiresConfirm": true,
      "selectedByDefault": false,  // never auto-selected
      "action": "acceptAll"
    },
    {
      "id": "sectionBreaks.orientation",
      "category": "sectionBreaks",
      "severity": "stabilisation",
      "label": "Section break changing orientation to Landscape",
      "detail": "Before paragraph ~120 (Schedule 3 table). Likely intentional.",
      "count": 1,
      "destructive": true,
      "requiresConfirm": true,
      "selectedByDefault": false,
      "protected": true,           // orientation change → human confirmation mandatory
      "action": "removeSectionBreak"
    }
  ],
  "metrics": {
    "fileSizeBeforeBytes": null,
    "fileSizeAfterBytes": null,
    "rsidCountBefore": null,
    "rsidCountAfter": null
  },
  "errors": []                     // { findingId, code: "GeneralException"|"AccessDenied"|..., message }
}
```

---

## 5. Algorithm — detection then action per capability

Each capability follows **detect → present in report → (user opts in / confirms) → apply → record result**.

### 5.1 Metadata stripping

**Detect**
- `context.document.properties.load("author,lastAuthor,subject,title,keywords,company,manager,comments,category,revisionNumber")` → any non-empty → finding.
- `properties.customProperties.load("items")` → each item → finding (list keys, redact values in UI unless expanded).
- `context.document.customXmlParts.load("items")` → for each part, `getXml()` and inspect namespace; tag known DMS namespaces (iManage `…/imanage…`, NetDocuments, generic `customXml/itemN.xml`) → finding `customXml`.
- `body.getComments()` → `.load("items")` → count → finding `comments`.

**Action**
- Built-in props: set each to empty string / appropriate empty value:
  ```js
  const p = context.document.properties;
  p.author = ""; p.subject = ""; p.title = ""; p.keywords = "";
  p.company = ""; p.manager = ""; p.comments = ""; p.category = "";
  // lastAuthor / revisionNumber are managed by Word; we clear what is writable and
  // note in the report any field Word re-populates on next save.
  ```
- Custom props: iterate and `item.delete()`.
- Custom XML parts: iterate `customXmlParts.items` and `part.delete()`. **Note (§8):** some DMS add-ins re-inject content controls / XML on the next save; we delete what is present and warn.
- Comments: iterate `comments.items` and `comment.delete()`.
- **`rsid` removal is NOT done here** — it is part of the OOXML loop (§5.6) because rsids live in `document.xml`/`settings.xml` attributes the typed API cannot reach.

### 5.2 Detect conflicting style overrides

**Detect**
- Iterate `body.paragraphs`; for each, compare its effective `font`/`paragraphFormat` to its named style's definition. Where direct formatting deviates from the named style (e.g. paragraph in `Normal` but font is Arial 11, or hard-coded `spaceAfter: 10`), flag a `styles` finding with a count of affected paragraphs/runs.
- Detect "mixed font within paragraph" by loading `paragraph.getTextRanges([" "], false)` or splitting runs and comparing `font.name`/`size`.

**Action** — see §5.3 (direct-format clearing) and §5.4 (profile application).

### 5.3 Clearing direct formatting (manual)

There is **no** API to clear direct formatting in one call (`clearOrResetDirectFormat` does not exist — see [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 6). We do it explicitly:

```text
For the selected scope (whole body, plus footnotes/endnotes — see 5.5):

A. Reset run-level (character) direct formatting:
   for each run/range r in scope:
     r.font.name = profile.body.font.name           // set explicit profile value (not "")
     r.font.size = profile.body.font.size
     r.font.bold = false; r.font.italic = false
     r.font.underline = Word.UnderlineType.none
     r.font.color = profile.body.font.color
     r.font.highlightColor = null                    // clear highlight

B. Reset paragraph-level direct formatting:
   for each paragraph p in scope:
     p.font.* = profile.body.font.*                  // paragraph-level font defaults
     p.paragraphFormat.lineSpacingRule = profile.body.spacing.lineSpacingRule
     p.paragraphFormat.lineSpacing     = profile.body.spacing.lineSpacing
     p.paragraphFormat.spaceBefore     = profile.body.spacing.spaceBefore
     p.paragraphFormat.spaceAfter      = profile.body.spacing.spaceAfter
     p.paragraphFormat.leftIndent      = profile.body.indent.left
     p.paragraphFormat.rightIndent     = profile.body.indent.right
     p.paragraphFormat.firstLineIndent = profile.body.indent.firstLine

C. Re-apply the named style so the paragraph re-binds to the style definition:
   p.style = "Normal"   (or the mapped target style)
   -- We apply STYLE FIRST, then explicitly set the profile values in B,
      so the explicit values win and there is no ambiguous "inherit".
```

**Batching:** chunk these writes across multiple `context.sync()` calls (windows of ~200 paragraphs) to avoid proxy-object and request-size limits (§8; and [Performance Guidelines](../ARCHITECTURE.md#performance-guidelines)).

**Emphasis preservation note.** A blunt reset destroys legitimate bold/italic (defined terms, case names in italics — important in SG/UK drafting). v1 policy: **direct-format clearing is its own opt-in checkbox separate from profile typography.** When the user only wants typography conformance (font/size/spacing) we apply §5.4 **without** the §5.3-A run-level bold/italic reset, preserving inline emphasis. The aggressive run-level reset is a distinct, clearly-labelled "Strip all inline formatting" option.

### 5.4 Apply style profile (typography, margins, spacing)

**Named styles (`getStyles`):**
```text
const styles = context.document.getStyles(); styles.load("items/nameLocal");
for each (styleName, def) in profile.styles:
   find style with nameLocal === styleName
   if found:
     style.font.name/size/bold/... = def.font.*
     style.paragraphFormat.lineSpacingRule/lineSpacing/spaceBefore/spaceAfter/... = def.paragraphFormat.*
   // built-in styles cannot be deleted but CAN be reset — that is exactly what we do.
```

**Margins (per section):**
```text
context.document.sections.load("items");
for each section s:
   s.pageSetup.topMargin    = profile.page.margins.top
   s.pageSetup.bottomMargin = profile.page.margins.bottom
   s.pageSetup.leftMargin   = profile.page.margins.left
   s.pageSetup.rightMargin  = profile.page.margins.right
   // orientation: DO NOT force portrait on a section flagged as an
   // intentional landscape Schedule (§7). Margin enforcement is OPT-IN.
```

**Body spacing/typography:** apply §5.3-B values across body paragraphs (chunked).

### 5.5 Footnote / endnote propagation (styles do NOT auto-propagate)

A style or direct-format change applied to `body` does **not** reach footnote/endnote text. We iterate them explicitly:

```text
body.footnotes.load("items"); body.endnotes.load("items");
for each note in [...footnotes, ...endnotes]:
   const nb = note.body;
   nb.paragraphs.load(...)
   apply the SAME profile typography/spacing (and, if "Footnote Text" style is in the
   profile, set size to profile.styles["Footnote Text"].font.size, e.g. TNR 10)
   run-level + paragraph-level resets exactly as 5.3, chunked.
```

Mandatory for SG/UK court documents where quoted authorities and footnoted citations have their own size (TNR 10) and must not be left at body size.

### 5.6 rsid bloat detection + stripping (OOXML loop — feature-specific surgery)

Driven through the [OOXML Service](../ARCHITECTURE.md#the-ooxml-service); only the rsid-specific edits are described here.

**Detect**
- `readDocx()`, parse `word/document.xml`, count attributes matching `^w:rsid(R|RDefault|P|Del|Tr|Sect)?$`, and count `<w:rsid>` entries in `word/settings.xml`.
- Heuristic threshold for "bloat": `rsidCount > 200` OR `(rsidCount / paragraphCount) > 3`. Below threshold we still offer stripping but mark severity `info`.
- Record `fileSizeBeforeBytes` and `rsidCountBefore` in `report.metrics`.

**Action** — within the OOXML loop:
- Strip all matching `w:rsid*` attributes from every element in `document.xml` (use the service's `stripAttributes(part, ['w:rsid*'])`; prefer targeted regex over a full XML re-serialise to preserve byte fidelity).
- Remove the `<w:rsids>` block (and child `<w:rsid>`) from `settings.xml`.
- Write back, re-measure `fileSizeAfterBytes` / `rsidCountAfter` for the report.

rsids are purely revision-tracking salt with no semantic meaning; stripping is safe and is the single biggest file-size win on a heavily-edited document.

### 5.7 Section-break detection + removal (OOXML loop, flag-and-confirm)

Section breaks cannot be deleted via the typed API ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 9), so this is OOXML surgery.

**Detect**
- Typed pass first: `context.document.sections.load("items")` plus each `section.pageSetup` (orientation, pageWidth/Height, margins) → enumerate section count and per-section geometry to **describe** each break to the user.
- OOXML pass: locate each `<w:sectPr>` in `document.xml`. Classify each break:
  - **Orientation-changing** (this section's orientation ≠ neighbour's) → `protected: true`. Almost always an intentional landscape Schedule/wide table in a deed/agreement.
  - **Geometry-changing** (different page size/margins) → `requiresConfirm: true`.
  - **Cosmetic / continuous with no geometry change** → still `requiresConfirm: true` but lower severity.

**Action (only on confirmed, non-protected breaks, or protected breaks the user explicitly confirms):**
- In `document.xml`, for each confirmed paragraph-embedded `<w:sectPr>` at a paragraph boundary: remove the `<w:p>` that wraps it OR merge its properties into the final body-level `<w:sectPr>`, so the artificial break collapses and following content flows into the previous section.
- **Never** touch a `<w:sectPr>` whose orientation differs from the body `sectPr` unless explicitly confirmed.
- Write back. Because in-place replace can add a stray paragraph/blank page, prefer the OOXML Service's `createDocument().open()` fallback if a stray-paragraph regression is detected (§8).

**Hard rule:** never strip an orientation break without explicit, per-break confirmation — per [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles), legitimate landscape Schedules must be protected.

### 5.8 Normalise spacing + delete blank paragraphs

**Detect**
- Blank paragraphs: iterate `body.paragraphs`; `paragraph.text === ""` (also treat whitespace-only as blank if option enabled). Count consecutive runs of blanks.
- Spacing anomalies: paragraphs whose `spaceBefore/spaceAfter/lineSpacing` differ from the profile.

**Action**
- Delete blank paragraphs (`paragraph.delete()`), chunked. **Preserve** a single intentional blank where it is the only separator before a signature block / Schedule heading if the option "collapse multiple blanks to one" is chosen rather than "delete all blanks".
- Spacing normalisation: apply profile `spaceBefore/spaceAfter/lineSpacing/lineSpacingRule`. **Opt-in** — never on by default, and disabled/warned when a court profile is active and the document already matches court spacing (§7).

### 5.9 Accept all tracked changes (gated)

**Detect**
- `body.getTrackedChanges()` → `.load("items")` → count → finding `tracked` with `selectedByDefault: false`, `requiresConfirm: true`. Gated on WordApi 1.6, feature-detected ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 4).

**Action (explicit, separate confirmation only):**
- `trackedChanges.acceptAll()` then `context.sync()`.
- Per-author selection is **not** offered (API limitation). The confirmation dialog states clearly: "This accepts ALL tracked changes from ALL authors. This cannot be selectively undone."
- **Orphaned-colour cleanup (§8):** documents processed by some comparison tools sometimes leave coloured runs (manual colour, not real tracked changes) after acceptance. After `acceptAll`, if the user also selected "strip inline formatting", the §5.3-A run reset clears residual colour; otherwise we flag remaining coloured runs in the report but do not auto-recolour (could destroy intentional emphasis).

---

## 6. SG/UK config keys used

Repair reads its court-typography rules from the central config; the **values** live in [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles). Feature-local profile-rule keys (carried in each Style Profile, §4.1):

- `rules.normalisationOptIn` — line-spacing/margin normalisation requires explicit opt-in.
- `rules.protectOrientationSectionBreaks` — orientation section breaks require per-break confirmation.
- `rules.citationBracketStyle` (`square` for SG PD 2021), `rules.quotedAuthoritySize` (TNR 10).

Mandatory behaviours (never single-space under a court profile; warn if a commercial profile is selected on a court-looking document; protect orientation breaks; shipped profiles `locked`) are defined in [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles) and are not restated here; §7 gives the Repair-specific UI guards.

---

## 7. UI / task-pane behaviour

### 7.1 Layout

```
┌──────────────────────────────────────────────┐
│ Repair                                        │
│ Profile: [ SG Court Filing (PD 2021) ▼ ]      │  ← profile selector (drives standardisation)
│ "All repairs run on your device. Nothing is   │
│  uploaded."                                    │  ← persistent trust banner
├──────────────────────────────────────────────┤
│ [ Scan document ]                             │  ← runs detection, builds report
├──────────────────────────────────────────────┤
│ Findings (opt-in checklist):                  │
│  CONFIDENTIALITY                              │
│   [✓] Clear document properties (2)           │
│   [✓] Remove comments (5)                      │
│   [✓] Delete DMS custom XML parts (3)          │
│  STABILISATION                                │
│   [✓] Strip rsid bloat  (12,540 → file -41%)   │
│   [ ] Reset direct formatting                  │
│   [ ] Conform to profile typography/margins    │
│   [ ] Normalise spacing      ⚠ court profile   │
│   [✓] Delete blank paragraphs (18)             │
│  DESTRUCTIVE (explicit confirm required)      │
│   [ ] Accept ALL tracked changes (47) 🔒       │
│   [ ] Remove section break → Landscape (1) 🔒  │
├──────────────────────────────────────────────┤
│ [ Preview changes ]   [ Apply selected ]      │
└──────────────────────────────────────────────┘
```

### 7.2 Behaviour rules

- **Scan first.** Nothing is changed by scanning. Confidentiality items default **checked**; stabilisation typography/spacing default **unchecked** (except blank-para deletion and rsid strip, which are safe and default checked); destructive items always **unchecked**.
- **Preview/report.** "Preview changes" shows the repair report (§4.2) describing exactly what each selected action will do, including metrics estimates (file-size reduction for rsid strip, paragraph counts). No mutation yet.
- **Explicit confirmation for destructive ops.** Accepting tracked changes and removing any section break each pop a modal confirm with the specific warning text. Orientation-break removal additionally warns that the Schedule reverts to portrait. These 🔒 items cannot be batch-applied without their individual confirm.
- **Profile selector** changes which standardisation values apply and re-evaluates warnings. Concrete court-profile guards (normalisation defaults unchecked when the doc already matches court spacing; warn if a commercial profile is selected on a court-looking document — double-spaced + ≥35 mm + TNR12 heuristic) follow [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles).
- **Apply** runs the selected actions, then re-renders the report in `phase: "apply"` with per-finding `result` and final metrics, plus a summary line, e.g. "Repaired: cleared 2 properties, removed 5 comments, deleted 3 DMS parts, stripped 12,540 rsids (file 41% smaller), deleted 18 blank paragraphs."
- **Undo guidance.** Typed-API changes are covered by Word's undo stack; OOXML-loop replacements (rsid/section break) replace the body and may not be cleanly undoable — the UI warns to keep a copy before OOXML operations.
- **Metadata-consciousness (SRA / PDPA).** The trust banner persists; custom XML (DMS) deletion is surfaced prominently because that is where client/matter identifiers and privilege markers most often hide.

---

## 8. Edge cases & failure handling

1. **DMS re-injection.** iManage/NetDocuments add-ins may re-inject content controls / custom XML on the **next save**. We delete what is present at repair time and warn: *"Your DMS may re-add metadata when you save. Repair again before sending, or detach the document from the DMS."* We cannot prevent re-injection from in-process.
2. **Orphaned coloured runs from some comparison tools.** After `acceptAll`, comparison tools sometimes leave runs with manual font colour (not true revisions). We detect remaining coloured runs and report them; we only auto-clear colour if the user selected "strip inline formatting" (§5.3-A), to avoid destroying intentional emphasis.
3. **Protected/locked documents.** Office JS cannot reliably read document protection state. Write attempts on a protected document throw. We **catch `GeneralException` / access errors gracefully**, mark the affected finding `result: "error"` with a friendly message (*"This document appears to be protected. Remove protection (Review → Restrict Editing) and try again."*), and continue with the remaining actions rather than aborting the whole run.
4. **Compatibility-mode documents.** We cannot detect compatibility mode via the API. We note in the report that results may vary and recommend converting (File → Info → Convert). No hard failure.
5. **`insertFileFromBase64` fidelity edge cases.** In-place `replace` can add a **stray empty paragraph / blank page** in some builds. Mitigation: after write-back, check for a trailing empty paragraph and delete it, or re-run via the OOXML Service `createDocument().open()` fallback; Repair chooses the fallback automatically on a detected regression.
6. **Large documents.** Slice the file read; chunk typed-API writes in windows (~200 paragraphs per `sync()`); show a progress indicator; warn above a size threshold (e.g. > 8–10 MB) that OOXML operations may take time. See [Performance Guidelines](../ARCHITECTURE.md#performance-guidelines).
7. **Empty / single-paragraph documents.** Guard against `paragraphs.items.length === 0`; section-break and rsid passes are no-ops; report says "nothing to repair".
8. **Document with no `settings.xml` or unusual package layout.** OOXML Service part lookup returns undefined → skip rsid-table removal gracefully (still strip `document.xml` rsid attributes); report notes partial action.
9. **Mixed-author tracked changes.** We cannot offer per-author accept (API limit). The confirm dialog states this explicitly.

---

## 9. Test cases / acceptance criteria

| # | Scenario | Steps | Pass criteria |
|---|---|---|---|
| T1 | Scrub metadata | Open sample with author/lastAuthor, 3 custom props, 2 DMS custom XML parts, 5 comments. Scan → select all confidentiality → Apply. | All built-in props empty; all custom props deleted; both custom XML parts gone (`customXmlParts.items.length === 0`); 0 comments. Report lists each. No network egress (verify via dev tools — zero outbound requests). |
| T2 | Apply SG Court Filing profile | Open mixed-format sample. Select `sg-court-filing`, check "conform to profile" + "footnote propagation". Apply. | Body font = TNR 12; line spacing = double (`lineSpacingRule` multiple / 2.0); all four margins = 99.21 pt (35 mm) per section; footnote text = TNR 10; Quotation style = TNR 10. |
| T3 | Detect + confirm + remove stray section break | Open doc with one **continuous/cosmetic** section break (same geometry) plus one **landscape orientation** break. Scan. | Both detected and described. Cosmetic break is `requiresConfirm`; orientation break is `protected: true`. Remove the cosmetic break only after confirm → section count drops by 1, content reflows, landscape Schedule untouched. Orientation break NOT removed without explicit per-break confirm. |
| T4 | Strip rsid bloat | Open heavily-edited doc with >10,000 rsid attributes. Scan (record `fileSizeBeforeBytes`, `rsidCountBefore`). Apply rsid strip. | `rsidCountAfter === 0`; `<w:rsids>` gone from settings.xml; **file size measurably reduced** (report shows e.g. −40%); document opens and renders identically. |
| T5 | Accept tracked changes (gated) | Open doc with 47 tracked changes. Confirm the action is **unchecked by default**; selecting it triggers the explicit modal stating "all authors / cannot selectively undo". Confirm → Apply. | `getTrackedChanges().items.length === 0` after; modal was required; action never ran as part of a general "apply all". |
| T6 | Court-profile normalisation guard | Open an already-double-spaced 35 mm SG court doc under `sg-court-filing`. | "Normalise spacing" and "Normalise margins" checkboxes are **unchecked** with the warning shown; applying other repairs leaves spacing/margins intact. |
| T7 | Direct-format clearing preserves emphasis (default) | Doc with italic case names + a paragraph with Arial 14 override. Apply "conform to profile" WITHOUT "strip inline formatting". | Arial 14 paragraph conforms to profile font/size; italic case names **preserved**. With "strip inline formatting" checked, italics removed too. |
| T8 | Protected document | Open a document with editing restrictions. Apply any write action. | `GeneralException` caught; finding marked error with friendly message; other actions either skipped cleanly or also report error; **no crash**, task pane remains usable. |
| T9 | Stray-paragraph fallback | Force the in-place replace path to add a trailing blank paragraph (test fixture). | Repair detects and removes the stray paragraph, or completes via `createDocument().open()` fallback; final document has no extra blank page. |

---

## 10. Known limitations

- **No per-author tracked-change control.** Bulk `acceptAll`/`rejectAll` only ([Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 4). Stated in-product.
- **Document protection state is undetectable.** Discovered only by a thrown exception and handled gracefully.
- **Compatibility mode is undetectable.** Some formatting may behave differently; we warn rather than guarantee.
- **Some package-level XML is out of reach in-browser even via the OOXML loop.** Font tables (`fontTable.xml`), theme parts (`theme/theme1.xml`), and certain styles internals are not safely editable client-side in v1 (see [The OOXML Service](../ARCHITECTURE.md#the-ooxml-service) fidelity caveats). Profile typography is applied through `getStyles()` + body iteration, not by hand-editing the styles part.
- **Orientation section breaks require human confirmation.** By design.
- **Page-number positioning (PD 2021 top-right) is flagged, not auto-inserted in v1.** Header/footer field insertion deferred; the report tells the user to add it manually.
- **`lastAuthor` / `revisionNumber` may be re-populated by Word** on the next save.
- **DMS re-injection** cannot be prevented from inside the add-in (§8.1).
- **OOXML-loop operations are not cleanly undoable** via Word's undo stack; users are advised to keep a copy.

---

## 11. Build milestones

| Milestone | Deliverable |
|---|---|
| **M1 — Detection & report skeleton** | Scan engine for metadata, comments, custom XML, blank paragraphs, tracked-change count, section enumeration. Repair-report model (§4.2) rendered read-only via the Findings shell. No mutation. |
| **M2 — Metadata & comments scrubbing** | Apply-path for built-in/custom properties, custom XML deletion, comment deletion. Trust banner. T1 passing. |
| **M3 — Style profiles + standardisation** | Profile schema + loader; ship SG Court Filing, UK Court Filing, Firm Commercial. Direct-format clearing (manual), profile application (getStyles + margins + body), footnote/endnote propagation. Court-profile normalisation guards. T2, T6, T7 passing. |
| **M4 — Spacing & blank-paragraph normalisation** | Opt-in spacing normalisation, blank-paragraph deletion with collapse-vs-delete option, chunked writes. |
| **M5 — OOXML loop integration** | Consume the OOXML Service for rsid detection/stripping and section-break detection/removal (flag-and-confirm, orientation protection). Metrics (file-size, rsid count). T3, T4 passing. |
| **M6 — Destructive ops & gating** | Accept-all-tracked-changes with explicit modal; orphaned-colour handling. T5 passing. |
| **M7 — Robustness & edge cases** | Protected-doc GeneralException handling, stray-paragraph fallback, large-document chunking/progress, compat-mode note. T8, T9 passing. |
| **M8 — Polish & ship** | Preview UX, summary lines, undo warnings, profile cloning/editing, accessibility, final SG/UK domain review sign-off. |

### Appendix — Shipped Style Profiles (fixtures)

JSON bodies of the three shipped profiles. The **rules** they encode (court typography, spacing, margins, orientation protection) are governed by [Court Formatting Profiles](../CONVENTIONS.md#court-formatting-profiles); these are the loadable fixtures.

**`sg-court-filing` — SG Court Filing (PD 2021):** TNR 12 body; TNR 10 quoted authorities/footnotes; double spacing; ≥35 mm (99.21 pt) all margins; A4 portrait; page numbers top-right (flagged); square-bracket `[xx]` paragraph cross-citation.

```jsonc
{
  "id": "sg-court-filing", "name": "SG Court Filing (PD 2021)",
  "jurisdiction": "SG", "kind": "court", "locked": true,
  "page": { "margins": { "top": 99.21, "bottom": 99.21, "left": 99.21, "right": 99.21 },
            "size": "A4", "orientation": "portrait",
            "pageNumber": { "position": "top-right", "enforced": false } },
  "body": { "font": { "name": "Times New Roman", "size": 12 },
            "spacing": { "lineSpacingRule": "multiple", "lineSpacing": 2.0,
                         "spaceBefore": 0, "spaceAfter": 0 },
            "indent": { "firstLine": 0, "left": 0, "right": 0 } },
  "styles": {
    "Normal":    { "font": { "name": "Times New Roman", "size": 12 },
                   "paragraphFormat": { "lineSpacingRule": "multiple", "lineSpacing": 2.0 } },
    "Quotation": { "font": { "name": "Times New Roman", "size": 10 },
                   "paragraphFormat": { "lineSpacingRule": "onePtFive", "lineSpacing": 1.5,
                                        "left": 36, "right": 36 } },
    "Footnote Text": { "font": { "name": "Times New Roman", "size": 10 } }
  },
  "rules": { "citationBracketStyle": "square", "quotedAuthoritySize": 10,
             "normalisationOptIn": true, "protectOrientationSectionBreaks": true }
}
```

**`uk-court-filing` — UK Court Filing:** TNR 12; 1.5 line spacing (double also accepted); ~25 mm (70.87 pt) margins; A4 portrait; page numbers bottom-centre or top-right (flagged); neutral-citation conventions. Court-bundle requirements vary by court/division — a sensible default; firms clone.

```jsonc
{
  "id": "uk-court-filing", "name": "UK Court Filing", "jurisdiction": "UK", "kind": "court",
  "locked": true,
  "page": { "margins": { "top": 70.87, "bottom": 70.87, "left": 70.87, "right": 70.87 },
            "size": "A4", "orientation": "portrait" },
  "body": { "font": { "name": "Times New Roman", "size": 12 },
            "spacing": { "lineSpacingRule": "onePtFive", "lineSpacing": 1.5,
                         "spaceBefore": 0, "spaceAfter": 0 } },
  "styles": { "Normal": { "font": { "name": "Times New Roman", "size": 12 },
                          "paragraphFormat": { "lineSpacingRule": "onePtFive", "lineSpacing": 1.5 } },
              "Footnote Text": { "font": { "name": "Times New Roman", "size": 10 } } },
  "rules": { "normalisationOptIn": true, "protectOrientationSectionBreaks": true }
}
```

**`firm-commercial` — Firm Commercial:** generic non-court default; TNR 11; single spacing with modest `spaceAfter`; 25.4 mm (72 pt) margins; portrait, landscape Schedules preserved; normalisation most useful here.

```jsonc
{
  "id": "firm-commercial", "name": "Firm Commercial", "jurisdiction": "GENERIC", "kind": "commercial",
  "locked": true,
  "page": { "margins": { "top": 72, "bottom": 72, "left": 72, "right": 72 },
            "size": "A4", "orientation": "portrait" },
  "body": { "font": { "name": "Times New Roman", "size": 11 },
            "spacing": { "lineSpacingRule": "single", "lineSpacing": 1.0,
                         "spaceBefore": 0, "spaceAfter": 6 } },
  "styles": { "Normal": { "font": { "name": "Times New Roman", "size": 11 },
                          "paragraphFormat": { "lineSpacingRule": "single", "spaceAfter": 6 } } },
  "rules": { "normalisationOptIn": false, "protectOrientationSectionBreaks": true }
}
```
