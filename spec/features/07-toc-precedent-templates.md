# 07 — Content: Table of Contents, Precedent / Clause Library, Templates (Content tab)

> **Status:** ✅ v1 (three sub-features). Covers **A — Table of Contents**, **B — Precedent / Clause Library**, **C — Templates (Variable Automation + Shared Data Sets)** in one file.
>
> Heading scheme for this multi-feature file: `#` file title, `##` per sub-feature (A/B/C) and shared sections, `###` template sections within each sub-feature. The standard feature template (Purpose / Scope / Dependencies / Data model / Algorithm / Config keys / UI / Edge cases / Tests / Limitations / Milestones) is applied **per sub-feature** at `###`.

These three "content" features are grouped because they all *generate, retrieve, or transform document body content* and share the same plumbing and jurisdictional posture.

---

## Shared scope, dependencies & conventions (all three sub-features)

### Dependencies

- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — style resolution, heading/Schedule paragraph enumeration, content-control registry, selection helpers.
- [Storage Abstraction](../ARCHITECTURE.md#storage-abstraction) — uniform `get/put/list` over `localStorage`, `document.settings`, JSON files, and the pluggable **GitHub** and **SharePoint / Microsoft Graph** adapters. No sub-feature talks to a backend directly.
- [Localisation / Config Layer](../ARCHITECTURE.md#localisation--config-layer) — jurisdiction profile (`SG` | `UK`/E&W), date format, and the [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields) values (entity-suffix lists, registration-number validators, execution-block templates). Firm TOC config is feature-local (§A).
- [Task Pane UI Shell](../ARCHITECTURE.md#task-pane-ui-shell) — tab host, ribbon buttons, toasts, modal confirms, settings.

### Office JS facts (not restated — see the shared reference)

These are ground truth from [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference); implementers must not assume APIs beyond them:
- TOC field insertion + switches via `field.code`, refresh via `field.updateResult()` (row 5; WordApi 1.5).
- Content controls (`.tag`/`.title`/`.text`/`.appearance`); **Repeating Section Content Control is NOT exposed** (row 15) → fixed-slot model, no OOXML cloning.
- Task panes **cannot initiate drag-and-drop** (row 16) → click-to-insert only.
- Cross-document reading impossible except base64 workarounds (row 10).

### SG/UK posture

- **Target Singapore and England & Wales only. No US conventions** — see [Jurisdiction Config Layer](../CONVENTIONS.md#jurisdiction-config-layer); US-isms are a bug.
- **SharePoint / Microsoft Graph is the FAVOURED default backend** for shared library/template/data-set content (M365 dominance + data-residency: privileged text stays in the firm tenant/region, not on US-hosted SaaS). GitHub is offered with a US-hosting caveat surfaced in the setup wizard. This is the per-feature UX rule layered over [Storage Abstraction](../ARCHITECTURE.md#storage-abstraction).
- Dates rendered **DD Month YYYY** ([Dates](../CONVENTIONS.md#dates)); entity fields per [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields).

### Global cut list ❌ (applies across the group)

Removed from this product — must not be implemented, stubbed, or referenced in UI:

- **Table of Authorities, citation parsing, Bluebook, OSCOLA** ❌ — entirely cut (US/citation-tooling territory).
- **KMS integrations (iManage / NetDocuments / OpenText)** ❌ — not in the open-source core; only a closed enterprise plugin adapter against the Storage interface.
- **True repeating-section cloning / OOXML surgery for repeating data** ❌ — impossible in Office JS; replaced by fixed slots.
- **Any US conventions** ❌ ("Inc.", "LLC", "EIN", "State of incorporation", month-first dates, Bluebook).
- **Drag-and-drop insertion** ❌ — impossible in task panes; click-to-insert only.

Consolidated rationale: [README → What was cut and why](../README.md#what-was-cut-and-why-).

---

# Sub-feature A — Table of Contents (TOC)

### A.1 Purpose & user value

Legal documents (contracts, deeds, prospectuses, court bundles) are navigation-heavy. Lawyers need to: (1) insert a **native, updatable Word TOC** that reflects firm house style without hand-tuning field codes; (2) **refresh** all TOCs after heavy editing in one click; (3) generate a separate **"Index of Schedules"** — a recurring SG/UK pain point because Schedules are often styled differently from body headings and are silently omitted from a heading-based TOC. User value: zero field-code literacy; consistent firm-branded TOCs; no missed Schedules.

### A.2 Scope

**In scope ✅:** insert one or more TOCs via the native Word TOC field; apply **firm formatting** from a JSON config mapping heading styles → TOC levels and switch sets; **multiple TOCs** in one document via bookmarked ranges and `{ TOC \b BookmarkName }`; **Refresh all** / refresh selected; **Index of Schedules** via a custom-style TOC (`\t`) or collected Schedule-styled paragraphs.

**Cut ❌:** Table of Authorities and any citation/case-law indexing (removed entirely); cross-document / master-document TOCs; TOC of figures/tables/equations (possible later, not this group).

### A.3 Dependencies (feature-specific)

| Capability | API / module |
|---|---|
| Insert TOC field | `Range.insertField(Word.InsertLocation.before/replace, Word.FieldType.toc)` |
| Set/modify switches | `field.code` |
| Refresh | `field.updateResult()` |
| Locate existing TOCs | `context.document.fields.getByTypes([Word.FieldType.toc])` |
| Bookmark a range | `range.insertBookmark(name)` / Document Model bookmark helper |
| Enumerate Schedule paragraphs | Document Model `getParagraphsByStyle(styleName)` |
| Config | Localisation/Config layer → `tocProfile` (feature-local) |

WordApi version handling and the F9 fallback are per [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 5: feature-detect `isSetSupported('WordApi','1.5')`; if absent, insert via `range.insertOoxml` of a pre-built `w:fldSimple` TOC field, then `updateResult` is unavailable — the UI tells the user to press **F9** / right-click → Update Field. Document this fallback in the UI string table.

### A.4 Data model — TOC config schema

Stored in the Config layer (firm-level), overridable per document via `document.settings`.

```jsonc
{
  "tocProfiles": [
    {
      "id": "house-default",
      "label": "Firm Standard TOC",
      "switches": {
        "levels": "1-3",          // -> \o "1-3"
        "useHyperlinks": true,     // -> \h
        "hidePageNumbers": false,  // -> \n omitted; if true \n "1-3"
        "rightAlignPageNo": true,  // tab leader via style
        "preserveTabAndLeader": true
      },
      // Optional explicit style->level mapping. Emitted as \t "Style,Level; ...".
      // When present, this SUPERSEDES \o for those styles.
      "styleLevelMap": [
        { "style": "Heading 1", "level": 1 },
        { "style": "Heading 2", "level": 2 },
        { "style": "Heading 3", "level": 3 },
        { "style": "Clause Heading", "level": 1 },
        { "style": "Sub-Clause", "level": 2 }
      ],
      "tabLeader": "dots"          // dots | dashes | none  (applied via TOC styles, not field switch)
    }
  ],

  "scheduleIndex": {
    "enabled": true,
    "label": "Index of Schedules",
    "scheduleStyles": ["Schedule Heading", "Schedule Title"],
    "level": 1,
    "separateField": true,        // own TOC field with \t over scheduleStyles only
    "placement": "afterMainToc"   // afterMainToc | atCursor | beforeFirstSchedule
  }
}
```

#### A.4.1 Switch derivation rules

- `\o "1-3"` emitted from `switches.levels` **only when** `styleLevelMap` is empty.
- `\h` emitted when `useHyperlinks`.
- `\b BookmarkName` emitted for **scoped** TOCs (multiple-TOC case).
- `\t "Style,Level;Style,Level"` emitted from `styleLevelMap` and for the Schedule index from `scheduleIndex.scheduleStyles`.
- `\n "1-3"` emitted when `hidePageNumbers`.
- Tab leaders are **not** field switches — they are a property of the `TOC 1..9` paragraph styles; applied via the Document Model style API, not the field code.

### A.5 Algorithm / workflows

#### A.5.1 Insert a TOC
```
1. Resolve active tocProfile (document override → firm default).
2. Determine insertion target:
     - "atCursor" → Word.getSelection().getRange('Start')
     - "topOfDocument" → body.getRange('Start')
3. range.insertField(InsertLocation.before, FieldType.toc) → field.
4. Build switch string from profile (A.4.1). Set field.code = `TOC ` + switches.
5. await context.sync(); field.updateResult(); await context.sync().
6. If tabLeader/styles configured, apply TOC-style updates.
7. Register the field id in document.settings.tocRegistry[] for "refresh all".
```

#### A.5.2 Multiple TOCs (scoped)
```
1. User selects a range → "Make scoped TOC here".
2. Generate a stable bookmark name: TOC_SCOPE_<slug>_<counter>.
3. range.insertBookmark(name).
4. Insert TOC field as A.5.1 but append `\b <name>` to the switch string.
5. Persist {fieldId, bookmark, scopeLabel} to tocRegistry.
```

#### A.5.3 Refresh
```
- "Refresh this TOC": locate field by registry id → field.updateResult().
- "Refresh all":
    fields = document.fields.getByTypes([FieldType.toc]); load("code");
    for each f: f.updateResult();
    one context.sync() (batched).
- After refresh, reconcile registry: drop entries whose field no longer exists.
```

#### A.5.4 Index of Schedules
Two strategies; the config picks one.

**Strategy 1 — custom-style TOC field (preferred, stays live):**
```
1. Build a TOC field whose ONLY switch is
     \t "Schedule Heading,1;Schedule Title,1" \h     (no \o so body Heading styles are excluded).
2. Insert at scheduleIndex.placement.
3. updateResult().
```
A self-updating list limited to Schedule-styled paragraphs.

**Strategy 2 — collected snapshot (fallback when Schedules are not consistently styled):**
```
1. Document Model: getParagraphsByStyle(scheduleStyles[]) → ordered list.
2. Build an HTML list (each item links to a bookmark inserted at the paragraph).
3. insertHtml at placement.
   NOTE: SNAPSHOT, not a live field — label as such; offer "Regenerate Index of Schedules".
```

> SG/UK rationale: Schedules frequently use a bespoke "Schedule Heading" style and restart numbering, so a plain `\o "1-3"` TOC misses them. Strategy 1 is the correct default; Strategy 2 covers legacy precedents with inconsistent styling. Terminology per [Terminology](../CONVENTIONS.md#terminology).

### A.6 SG/UK config keys used
- `tocProfiles[].switches.rightAlignPageNo` + `tabLeader: "dots"` — dot leaders + right-aligned page numbers are the SG/UK house norm.
- `scheduleIndex.label` = "Index of Schedules" — never "Table of Authorities".
- Heading style names default to UK/SG firm styles (`Clause Heading`, `Schedule Heading`) and are fully config-driven so firms remap. Label text is localised (en-GB). Schedule terminology from [Terminology](../CONVENTIONS.md#terminology).

### A.7 UI / task-pane behaviour
TOC panel with: **Profile selector** (firm profiles from config); **Insert TOC** (at cursor or top, per profile); **Make scoped TOC from selection** (enabled only when a non-empty range is selected); **Generate / Regenerate Index of Schedules** (visible when `scheduleIndex.enabled`); **Refresh this TOC** and **Refresh all TOCs**; a list of registered TOCs with scope label + "go to" (`range.select()`). If WordApi 1.5 is unavailable, the panel shows the F9-update banner (A.3).

### A.8 Edge cases
- **No headings yet:** inserting a TOC over a doc with no mapped styles yields "No table of contents entries found." Detect (empty `field.result.text`) and toast a hint to apply heading styles.
- **`.dotx` templates:** TOC fields insert fine but page numbers are meaningless; suppress page numbers (`\n`) when the document is detected as a template (see C.8 for detection).
- **Manual deletion of a TOC** outside the registry → reconcile on refresh (A.5.3).
- **Bookmark collision** for scoped TOCs → counter-suffixed names, verify uniqueness against `document.getBookmarks`.
- **Schedules without a style** → Strategy 1 returns empty; UI offers Strategy 2 with a snapshot warning.

### A.9 Test cases / acceptance criteria

| # | Scenario | Expected |
|---|---|---|
| A-T1 | Insert TOC into a doc with Heading 1–3 | Native TOC field appears, levels 1–3, hyperlinked, dot leaders. |
| A-T2 | Edit headings, click **Refresh all** | Entries + page numbers update; no duplicate TOC. |
| A-T3 | Select a Part, **Make scoped TOC** | Field code contains `\b TOC_SCOPE_…`; lists only that range. |
| A-T4 | **Generate Index of Schedules** with two "Schedule Heading" paragraphs | Separate field listing exactly those two Schedules, excluding body headings. |
| A-T5 | WordApi 1.5 absent | Fallback insert + F9 banner shown; no crash. |

### A.10 Known limitations
- TOC entry styling beyond switch+style level (e.g. per-entry colour) is not exposed.
- Strategy 2 Index of Schedules is a snapshot, not live.
- No figure/table TOCs in this release.

### A.11 Build milestones
1. **A-M1** Single native TOC insert + refresh with default profile.
2. **A-M2** Config-driven profiles (style/level map, switches, leaders).
3. **A-M3** Scoped multiple TOCs via bookmarks + registry + refresh-all.
4. **A-M4** Index of Schedules (Strategy 1, then Strategy 2 fallback).
5. **A-M5** WordApi version fallback + edge-case hardening.

---

# Sub-feature B — Precedent / Clause Library

### B.1 Purpose & user value

Lawyers reuse vetted clauses constantly. This feature lets a user: (1) **save** a selected clause to a personal or shared library with **risk-management metadata**; (2) **search** the library and **click-to-insert** a vetted clause at the cursor; (3) see **jurisdiction** and **last-reviewed** information **at insert time**, preventing the classic malpractice error of dropping a NY-law clause into an English-law contract.

### B.2 Scope

**In scope ✅:** save selected clause → library entry with mandatory metadata; backends through the Storage abstraction — **GitHub-backed** (zero infra, Fuse.js client index) or **SharePoint / Microsoft Graph** (tenant-controlled); **search** (Fuse.js fuzzy index for client-side backends); **click-to-insert** with a **jurisdiction warning** when clause jurisdiction ≠ document jurisdiction; insert as **HTML** (portability default) or **OOXML blob** (table-fidelity option).

**Cut ❌:** KMS integrations (enterprise plugin adapter only, not core); drag-and-drop insertion (impossible — click-to-insert only); server-side full-text search service in core (large-library guidance in B.8, but no bundled server); AI clause generation / suggestion (out of this group).

### B.3 Dependencies (feature-specific)

| Capability | API / module |
|---|---|
| Read selected clause | `Word.getSelection()` → `range.getHtml()` and/or `range.getOoxml()` |
| Insert clause | `range.insertHtml(html,'Replace')` (default) or `range.insertOoxml(ooxml,'Replace')` |
| Wrap inserted clause for tracking | `range.insertContentControl(RichText)`, set `.tag = "CLAUSE:<id>"`, `.title` |
| Storage | Storage abstraction (GitHub / SharePoint adapters) |
| Search | **Fuse.js** over the loaded index (client backends) |
| Jurisdiction of doc | Config layer document profile (`SG`/`UK`/E&W) |

### B.4 Data model — clause-library entry schema

Metadata marked **REQUIRED** is enforced at save time; Save is blocked until provided.

```jsonc
{
  "id": "clz_2026_0142",                 // stable id (uuid or content hash)
  "title": "Limitation of Liability (cap at fees)",     // REQUIRED
  "body": {
    "html": "<p>…</p>",                  // REQUIRED (portable default)
    "ooxml": "<?xml …?>",                // OPTIONAL (table-fidelity insert)
    "plainText": "…"                     // derived, for indexing
  },
  "metadata": {
    "jurisdiction": "EW",                // REQUIRED enum: "SG" | "EW" | "OTHER"
    "jurisdictionOther": "NY",           // REQUIRED iff jurisdiction == "OTHER"
    "practiceArea": "Commercial",        // REQUIRED (config-controlled taxonomy)
    "lastReviewed": "2026-02-14",        // REQUIRED, ISO date (displayed DD Month YYYY)
    "approvedBy": "J. Tan (Partner)",    // REQUIRED
    "visibility": "shared",              // REQUIRED enum: "personal" | "shared"
    "tags": ["liability", "cap", "indemnity"],   // optional
    "language": "en-GB",                 // en-GB | en-SG
    "sourcePrecedent": "MSA v3.2",       // optional provenance
    "notes": "Use only with mutual cap clause."   // optional drafting note
  },
  "audit": {
    "createdBy": "alice@firm",
    "createdAt": "2026-02-14T09:11:00Z",
    "updatedAt": "2026-02-14T09:11:00Z",
    "version": 3                          // increments on edit; drives drift warning (B.8)
  }
}
```

#### B.4.1 Validation at save
- `jurisdiction` ∈ {SG, EW, OTHER}; if OTHER, `jurisdictionOther` required (free text, flagged non-target — B.6).
- `lastReviewed` not in the future; warn if older than a configurable staleness window (default 24 months).
- `approvedBy` non-empty (risk governance).
- `visibility=shared` requires a configured shared backend; otherwise the option is disabled.

#### B.4.2 Search index (Fuse.js)
```jsonc
{
  "keys": [
    { "name": "title", "weight": 0.5 },
    { "name": "metadata.tags", "weight": 0.2 },
    { "name": "metadata.practiceArea", "weight": 0.15 },
    { "name": "body.plainText", "weight": 0.15 }
  ],
  "threshold": 0.34,
  "ignoreLocation": true,
  "minMatchCharLength": 3
}
```
Filters (faceted, applied **before** Fuse on the candidate set): `jurisdiction`, `practiceArea`, `visibility`, staleness.

### B.5 Algorithm / workflows

#### B.5.1 Save a clause
```
1. range = getSelection(); load isEmpty.
2. html = range.getHtml(); ooxml = range.getOoxml() (optional capture).
3. plainText = range.text (for indexing).
4. Open Save dialog → collect REQUIRED metadata (B.4.1). Block save until valid.
5. Build entry; compute id (uuid).
6. Storage.put(entry):
     - GitHub adapter: write JSON to /clauses/<id>.json, update /index.json, commit via Git Data API.
     - SharePoint adapter: create list item + attach body to a document library; Graph PATCH.
7. Update in-memory Fuse index incrementally.
8. Toast success; offer "wrap original selection in a CLAUSE content control" for later provenance (optional).
```

#### B.5.2 Search
```
1. Load index (GitHub: fetch /index.json; SharePoint: Graph query top N).
2. Apply facet filters → candidate set.
3. fuse.search(query) → ranked results.
4. Render result cards (title, jurisdiction chip, practiceArea, lastReviewed, approvedBy, visibility).
```

#### B.5.3 Click-to-insert with jurisdiction warning
```
1. User clicks "Insert" on a result.
2. docJurisdiction = Config.document.jurisdiction (SG|EW).
3. If entry.metadata.jurisdiction != docJurisdiction OR entry.metadata.jurisdiction == "OTHER":
        Show BLOCKING confirm modal:
          "This clause is governed by <X>. The current document is <Y>.
           Inserting cross-jurisdiction clauses is a known risk. Insert anyway?"
        Require explicit "Insert anyway".
4. Choose insert mode:
      - default: range.insertHtml(entry.body.html, 'Replace')
      - if entry has ooxml AND user picked "preserve tables":
            range.insertOoxml(entry.body.ooxml, 'Replace')
5. Wrap inserted range in a RichText content control:
      cc.tag = "CLAUSE:" + id + ":v" + audit.version
      cc.title = entry.title + " (rev " + lastReviewed + ")"
   → enables later drift detection (B.8).
6. context.sync().
```

> The jurisdiction warning is **mandatory and non-suppressible per-insert** for cross-jurisdiction inserts. A firm may, via config, downgrade *same-region but different-sub-jurisdiction* warnings, but SG↔EW↔OTHER always warns.

### B.6 SG/UK config keys used
- **Jurisdiction enum centred on SG and E&W**; anything else is `OTHER` and elevated risk (warning copy names it).
- **SharePoint / Microsoft Graph is the favoured shared backend** (per the shared posture above): data residency keeps clause text in the firm tenant/region; GitHub allowed with a US-hosting caveat surfaced in setup. Onboarding wizard: if jurisdiction ∈ {SG, EW} and no backend chosen → recommend SharePoint first.
- **`lastReviewed` display** is **DD Month YYYY** per [Dates](../CONVENTIONS.md#dates).
- **Language** facet `en-GB` / `en-SG` ([en-GB Spelling](../CONVENTIONS.md#en-gb-spelling)); no `en-US` variants promoted.

### B.7 UI / task-pane behaviour
- **Browse / Search panel:** search box (debounced Fuse query) + facet chips (jurisdiction, practice area, visibility, "stale only"); result cards show title, **jurisdiction chip** (SG/E&W/OTHER, colour-coded), practice area, **last-reviewed (DD Month YYYY)**, approved-by, visibility icon. Per card: **Insert** (primary), **Insert (preserve tables)** if OOXML present, **Preview**, **Edit metadata** (if permitted).
- **Save panel:** appears on "Save selection as clause"; required-field form with inline validation; visibility toggle disabled if no shared backend.
- **Insert confirm modal** for cross-jurisdiction (B.5.3).
- **Backend status** indicator (which adapter, last sync, item count).

### B.8 Edge cases
- **Formatting / style conflicts on insert:** `insertHtml` maps to the **document's** named styles; if the clause references a style absent in the target doc, Word substitutes → orphan/mismatched styles. Mitigation: prefer HTML (re-styled to host); document that OOXML carries its own style defs and can introduce duplicate/renamed styles (`Style1`, `Heading 10`). Smart-quote / encoding: `getHtml` may emit curly quotes / `&#8217;`; normalise on save, ensure UTF-8, round-trip-test that `insertHtml` renders identical glyphs. Escape special XML chars (`& < > " '`) in OOXML; reject control characters.
- **Clause version drift:** an inserted clause is a **copy** and goes stale when the library entry is updated. The wrapping content control stores `:v<version>`; a "Check clauses in this document" action compares against the library and flags outdated inserts, surfacing **lastReviewed** and the newer version. (No auto-update — the lawyer decides.)
- **Large libraries (>2000 clauses):** client-side Fuse.js index load + search degrades. Above a configurable threshold (default 2000) the GitHub `/index.json` approach is **not** recommended; UI warns and recommends the **SharePoint/Graph backend** with server-side query/paging. Document the Fuse.js client-side ceiling honestly.
- **Content controls inside tables:** an inserted clause containing content controls landing inside a host table cell can corrupt the cell when `insertText` carries embedded newlines. Guidance: for table-bearing clauses use **OOXML insert**, and never push multi-paragraph values into a *PlainText* control inside a table (see C.8).
- **`.dotx` vs `.docx`:** clause body insertion behaves the same; page-dependent fields differ but are not generally an issue here.
- **Empty selection on save** → block with a clear message.

### B.9 Test cases / acceptance criteria

| # | Scenario | Expected |
|---|---|---|
| B-T1 | Save a selected clause with **jurisdiction = SG**, practice area, last-reviewed, approved-by, visibility | Entry persisted; required-field validation enforced; appears in search. |
| B-T2 | Search "liability cap" | Fuzzy match returns the limitation clause ranked by title weight. |
| B-T3 | Document jurisdiction = E&W; insert the **SG** clause | **Blocking jurisdiction warning** naming SG vs E&W; insert only after explicit confirm. |
| B-T4 | Insert same-jurisdiction clause | No warning; HTML insert; wrapped in `CLAUSE:` content control with version tag. |
| B-T5 | Insert a table clause with "preserve tables" | OOXML path used; table fidelity retained. |
| B-T6 | Update a library entry, run "Check clauses" | Inserted copy flagged as drifted, showing newer last-reviewed. |
| B-T7 | Library exceeds 2000 entries on GitHub backend | UI warns and recommends SharePoint backend. |

### B.10 Known limitations
- **No drag-and-drop** — click-to-insert only.
- **Client-side search scale limit** (~2000 entries) for GitHub/Fuse.js; large libraries need SharePoint/Graph server-side query.
- Inserted clauses are copies; **no live linking / auto-update** (drift is flagged, not auto-resolved).
- **No KMS** in core.

### B.11 Build milestones
1. **B-M1** Save-clause flow with mandatory metadata + GitHub adapter + JSON index.
2. **B-M2** Fuse.js search + faceted browse UI + result cards (DD Month YYYY).
3. **B-M3** Click-to-insert (HTML), tracking content control, jurisdiction warning.
4. **B-M4** OOXML "preserve tables" insert + encoding/escaping hardening.
5. **B-M5** SharePoint/Graph adapter (favoured backend) + residency onboarding.
6. **B-M6** Drift detection ("Check clauses") + large-library guidance.

---

# Sub-feature C — Templates (Variable Automation + Shared Data Sets)

### C.1 Purpose & user value

Convert a precedent into a **fill-in template**: (1) an author **marks** selected text as a **tagged Plain Text content control** (e.g. `CLIENT_NAME`); (2) a drafter runs the **fill workflow** — the add-in scans tags, renders a **task-pane questionnaire**, and writes answers back via `insertText`; (3) **conditional clauses** are wrapped in tagged controls and hidden/shown based on answers (prefer **hide** over delete); (4) **Shared Data Sets** (e.g. a roster of directors) are saved and **pushed** into templates; (5) **repeating data** uses a **fixed-slot** model with the cap stated honestly in the UI.

### C.2 Scope

**In scope ✅:** **Mark variable** (wrap selection in a tagged PlainText content control); **questionnaire fill** (scan tags → form → write values); **conditional inclusion** (hide/show tagged sections via `control.appearance`; confirm before any destructive delete); **Shared Data Sets** (save/load groups of values; push into matching tags); **fixed-slot repeating** (N pre-placed slots, unused hidden); **SG/UK variable library** (party name, registration numbers, registered office, jurisdiction of incorporation, Business Day location, entity-type pickers, day-first dates, execution/attestation blocks).

**Cut ❌:** Repeating Section Content Control / OOXML cloning (not exposed; fixed slots instead); any US-isms ("Inc./LLC", "EIN", "State of incorporation", month-first dates — banned); calculated/derived fields beyond simple formatting (no expression engine this release); cross-document template propagation.

### C.3 Dependencies (feature-specific)

| Capability | API / module |
|---|---|
| Mark variable | `range.insertContentControl(Word.ContentControlType.plainText)`, set `.tag`, `.title` |
| Scan tags | `context.document.contentControls.load("tag,title,text,appearance")` |
| Write value | `control.insertText(value, Word.InsertLocation.replace)` |
| Conditional hide/show | `control.appearance` + Document Model wrapper |
| Destructive delete (confirmed) | `control.delete(keepContent=false)` |
| Data sets / templates persistence | Storage abstraction (`document.settings` in-doc; JSON/SharePoint shared) |
| Date / entity formatting | Localisation/Config layer → [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields), [Dates](../CONVENTIONS.md#dates) |

> Conditional **show/hide**: Office JS `contentControl.appearance` does not expose a true "visibility" flag. The **non-destructive default** keeps the control and clears/restores its stored content (cached in `document.settings`), so a hidden section can be re-shown. Destructive delete (`control.delete(keepContent=false)`) is opt-in and confirmed.

### C.4 Data model

#### C.4.1 Template variable schema
```jsonc
{
  "templateId": "tmpl_msa_ew_v2",
  "jurisdictionProfile": "EW",            // EW | SG  (drives defaults & validators)
  "variables": [
    {
      "tag": "CLIENT_NAME",               // machine name == content-control .tag
      "label": "Client name",             // == .title, shown in questionnaire
      "type": "text",                     // text|longtext|number|date|select|entityType|regNumber|address|boolean
      "required": true,
      "validation": { "maxLen": 120 },
      "default": ""
    },
    {
      "tag": "EFFECTIVE_DATE",
      "label": "Effective date",
      "type": "date",
      "required": true,
      "format": "DD MMMM YYYY"            // DAY-FIRST, never US month-first
    },
    {
      "tag": "COMPANY_NUMBER",
      "label": "Companies House number",
      "type": "regNumber",
      "regScheme": "UK_CH",               // UK_CH (8 chars) | SG_UEN
      "required": true,
      "validation": { "pattern": "^[A-Z0-9]{8}$" }
    },
    {
      "tag": "ENTITY_TYPE",
      "label": "Entity type",
      "type": "entityType",
      "options": ["Limited", "Ltd", "plc"],   // EW set; SG set differs (C.6)
      "required": true
    },
    {
      "tag": "GOVERNING_LAW",
      "label": "Governing law",
      "type": "select",
      "options": ["England & Wales", "Singapore"],
      "required": true
    }
  ],

  "conditionals": [
    {
      "tag": "SECTION_TUPE",              // a tagged section content control
      "label": "Include TUPE clause",
      "showWhen": { "var": "INCLUDES_STAFF", "equals": true },
      "onHide": "non-destructive"          // non-destructive | destructive-confirm
    }
  ],

  "repeating": [
    {
      "group": "DIRECTORS",
      "slotCount": 6,                      // FIXED-SLOT cap, stated in UI
      "slotTags": [                        // one set of tags per slot index 1..N
        { "field": "name",   "tagPattern": "DIRECTOR_{i}_NAME" },
        { "field": "uenOrCh","tagPattern": "DIRECTOR_{i}_ID" }
      ],
      "hideUnusedSlots": true
    }
  ]
}
```

#### C.4.2 Shared Data Set schema
```jsonc
{
  "dataSetId": "ds_acme_directors",
  "label": "Acme Pte Ltd — board roster",
  "scope": "shared",                       // personal | shared
  "jurisdiction": "SG",
  "kind": "repeating",                     // scalar | repeating
  "groupRef": "DIRECTORS",
  "rows": [
    { "name": "Lim Wei",   "uenOrCh": "201912345A" },
    { "name": "S. Patel",  "uenOrCh": "198700456K" }
  ],
  // For scalar data sets:
  "values": { "CLIENT_NAME": "Acme Pte. Ltd.", "REGISTERED_OFFICE": "…" }
}
```

#### C.4.3 Fixed-slot repeating model (explicit)
- A repeating group declares **N fixed slots** (`slotCount`, default 4, max 6).
- The template document **physically contains** N pre-placed slot blocks, each a set of tagged controls (`DIRECTOR_1_NAME … DIRECTOR_6_NAME`).
- Filling K rows (K ≤ N): write rows 1..K; **hide** slots K+1..N (non-destructive).
- K > N → UI error: "This template supports up to N directors; reduce the list or use a template with more slots." **No cloning** (Repeating Section CC not exposed — [Office JS Constraints Reference](../ARCHITECTURE.md#office-js-constraints-reference) row 15).

### C.5 Algorithm / workflows

#### C.5.1 Author marks a variable
```
1. User selects text → "Mark as variable".
2. Prompt for tag (UPPER_SNAKE, validated unique in doc) + label + type.
3. range.insertContentControl(PlainText).
4. cc.tag = tag; cc.title = label; cc.appearance = boundingBox (authoring view).
5. Persist field meta into document.settings.templateSchema.variables[].
6. context.sync().
```

#### C.5.2 Questionnaire fill
```
1. Scan: contentControls.load("tag,title,text,appearance").
2. Reconcile against templateSchema (warn on tags in doc but not schema and vice versa).
3. Render form: one control per variable, widget by type:
      text/longtext → input/textarea; number → numeric
      date → date picker, formatted DD MMMM YYYY on write
      select/entityType → dropdown (jurisdiction-aware options)
      regNumber → masked input + scheme validator (UK_CH / SG_UEN)
      address → multi-line; boolean → checkbox (also feeds conditionals)
4. Validate required + patterns.
5. On submit, for each tag:
      cc = controls.getByTag(tag).getFirst();
      cc.insertText(formatValue(type, value), 'Replace');
   Batch context.sync() (chunked, see C.8 perf).
6. Evaluate conditionals (C.5.3) and repeating slots (C.5.4).
```

#### C.5.3 Conditional show/hide
```
For each conditional:
  visible = evaluate(showWhen, answers).
  if visible:
      ensure section control present; restore cached content if previously hidden.
  else:
      if onHide == "non-destructive":
          cache control.getOoxml() into document.settings.hiddenSections[tag];
          clear content (insertText('', Replace)) keeping the control as a placeholder.
      if onHide == "destructive-confirm":
          show CONFIRM modal; only on confirm → control.delete(false).
Default is non-destructive; destructive requires explicit user confirmation.
```

#### C.5.4 Fixed-slot fill + data-set push
```
1. rows = dataSet.rows (or questionnaire repeating entries).
2. K = rows.length; N = group.slotCount.
3. if K > N → error (C.4.3). Stop.
4. for i in 1..K:
       for slotTag in group.slotTags:
           tag = slotTag.tagPattern.replace("{i}", i)
           controls.getByTag(tag).insertText(rows[i-1][slotTag.field], 'Replace')
5. for i in K+1..N: hide slot i (non-destructive).
6. context.sync().
```

**Data-set field-name vs tag fuzzy-match:** when pushing a **scalar** data set, match `values` keys to schema tags. Exact match preferred; on miss, run a **fuzzy match** (normalised, e.g. `client name` ↔ `CLIENT_NAME`) and **warn** before applying — never silently map. Unmatched keys are listed for manual mapping.

### C.6 SG/UK config keys used

Validators, entity-suffix sets, and execution blocks come from [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields); dates from [Dates](../CONVENTIONS.md#dates). Feature-local schema keys: `regScheme` (`UK_CH` | `SG_UEN`), `type: entityType` `options`, `format: "DD MMMM YYYY"`, `jurisdictionProfile`.

- **Registration numbers:** `UK_CH` (Companies House, 8 chars, `SC`/`NI`/`OC`/`SO` prefixes allowed) and `SG_UEN` (`nnnnnnnnX` / `yyyynnnnnX` / `Tyypqnnnnx`) — pattern sets held in config, not hardcoded.
- **Entity-type pickers (jurisdiction-aware):** EW = `Limited`, `Ltd`, `plc`, `LLP`; SG = `Pte. Ltd.`, `Ltd.`, `LLP`. **No** `Inc.`/`LLC`/`EIN`/"State of incorporation" in option lists or templates.
- **Dates:** all date content controls render **DD Month YYYY**, never month-first (ISO accepted as input, rendered day-first).
- **Registered office address**, **jurisdiction of incorporation** (E&W / Singapore / other), **Business Day location** ship as standard variables.
- **Execution / attestation blocks** (shipped templates): UK deed per Companies Act 2006 s.44 (two authorised signatories, **or** single director with witness; "EXECUTED as a DEED" + witness attestation — ship both s.44 variants); Singapore execution block (common seal / per constitution; director + secretary or two directors). Authoritative wording: [Entity & Template Fields](../CONVENTIONS.md#entity--template-fields).
- **Storage residency:** shared Data Sets and shared templates **favour SharePoint/Graph** (tenant residency, same SRA/PDPA reasons as B); GitHub allowed with the US-hosting caveat surfaced in setup.

### C.7 UI / task-pane behaviour
- **Author mode:** **Mark as variable** (selection → tag/label/type dialog); **Mark conditional section** (wrap selection, set `showWhen`); **Define repeating group** (declare group, slot count with the **honest cap** label "Up to 6 directors — unused slots are hidden"); schema panel listing all variables/conditionals/slots with edit/delete.
- **Fill mode:** auto-generated **questionnaire form** (grouped sections, required markers, live validation — reg-number masks, date pickers day-first); **conditional toggles** that immediately show/hide sections (non-destructive); destructive removal behind a confirm; **repeating editor** (add/remove rows up to the cap; cap reached → "Add" disabled with explanation).
- **Data-set manager:** save current answers as a data set (personal/shared); browse/load and **push** into template (scalar fuzzy-match warnings; repeating slot fill); backend/residency indicator.

### C.8 Edge cases
- **Content controls inside tables (corruption risk):** writing multi-line values via `insertText` into a **PlainText** control in a table cell can corrupt cell structure. Mitigations: (a) signatory/director slots in execution tables use **single-line** values only; (b) addresses needing line breaks use a **RichText** wrapper or separate single-line slots, never one `insertText` with embedded `\n` into a PlainText-in-table control.
- **Special XML / encoding:** values escaped for XML; smart quotes normalised; control characters rejected (mirrors B.8).
- **`.dotx` vs `.docx`:** authoring typically in **.dotx**; detect template via `Office.context.document.url`/settings and adjust (suppress page-number-dependent behaviour, keep content controls). On "New from template", Word produces a `.docx` instance carrying the controls — fill operates on the instance.
- **Fixed-slot blank handling:** unused slots hidden **cleanly** (non-destructive) so the printed deed shows no empty director lines; re-running fill with more rows must **re-show** previously hidden slots from cache.
- **Data-set field-name vs tag mismatch:** fuzzy match **warns**, never auto-applies silently (C.5.4).
- **Tag drift between doc and schema:** controls deleted in the body but present in schema (and vice versa) reconciled with a warning at fill start.
- **Conditional re-show after destructive delete:** a *destructively* deleted section **cannot** be auto-restored — UI warns at delete time that this is irreversible within the fill session.
- **Duplicate tags:** `getByTag` may return multiple; for scalar variables enforce uniqueness at author time; for slots, tags are unique by construction (`{i}`).

### C.9 Test cases / acceptance criteria

| # | Scenario | Expected |
|---|---|---|
| C-T1 | Author marks `CLIENT_NAME` and `EFFECTIVE_DATE` | Two PlainText controls created with correct `.tag`/`.title`; schema persisted. |
| C-T2 | Fill via questionnaire | Values written via `insertText`; **EFFECTIVE_DATE renders DD Month YYYY** (e.g. "06 June 2026"), never month-first. |
| C-T3 | Enter UK Companies House number `00012345`-style | Validated as 8-char `UK_CH`; SG_UEN rejected under EW profile and vice versa. |
| C-T4 | Toggle a conditional off (non-destructive) then on | Section hidden then restored from cache; no content loss. |
| C-T5 | Toggle conditional with `destructive-confirm` | Confirm modal shown; only deletes on confirm; irreversibility warned. |
| C-T6 | Push a **directors data set** (2 rows) into a 6-slot group | Slots 1–2 filled; slots 3–6 hidden cleanly; re-push with 4 rows re-shows slots 3–4. |
| C-T7 | Push data set with 7 directors into 6 slots | Blocked with honest cap error; no cloning attempted. |
| C-T8 | Render **UK-deed execution block** (s.44, director + witness) | "EXECUTED as a DEED" header + correct witness attestation; SG block available under SG profile. |
| C-T9 | Scalar data-set key `client name` vs tag `CLIENT_NAME` | Fuzzy-match **warning** shown before apply; no silent mapping. |

### C.10 Known limitations
- **No true repeating sections** — fixed-slot cap (default 4, max 6); stated honestly in UI; large rosters need a larger-slot template.
- **No drag-and-drop**; all fills are click/submit driven.
- **No expression/derived-field engine** (only type-based formatting like day-first dates).
- Conditional show/hide is content-control based, not a Word "visibility" property; destructive delete is irreversible mid-session.
- Shared data residency depends on the chosen backend; **SharePoint favoured** for SG/UK.

### C.11 Build milestones
1. **C-M1** Mark-variable authoring + schema persistence in `document.settings`.
2. **C-M2** Questionnaire scan + render + write-back (type widgets, validation).
3. **C-M3** SG/UK variable library: reg-number validators (UK_CH, SG_UEN), entity-type pickers, day-first date formatter.
4. **C-M4** Conditional show/hide (non-destructive default + confirmed destructive).
5. **C-M5** Fixed-slot repeating model + honest cap UI.
6. **C-M6** Shared Data Sets (save/load/push, fuzzy-match warnings) on SharePoint-favoured storage.
7. **C-M7** Ship UK-deed (CA 2006 s.44 variants) + Singapore execution-block templates; table-in-cell hardening.

---

## Appendix — Cross-feature dependency summary

| Concern | A (TOC) | B (Clauses) | C (Templates) |
|---|---|---|---|
| Office JS surface | TOC field + `getByTypes` | get/insertHtml/insertOoxml + content control | content controls + `insertText` + `appearance` |
| Storage backend | `document.settings` (registry/config) | GitHub or **SharePoint (favoured)** | `document.settings` + **SharePoint (favoured)** shared sets |
| Jurisdiction surfacing | labels ("Index of Schedules") | **mandatory** jurisdiction warning at insert | profile-driven validators/defaults/execution blocks |
| Date format | n/a | last-reviewed **DD Month YYYY** | variables **DD Month YYYY** |
| Hard "no" ❌ | Table of Authorities | drag-drop, KMS, >2000 client search | repeating-section cloning, US-isms |
