# 11 — AI-Readable Export ("Copy for AI")

**Status:** ✅ v1 — Tier-1 per-panel exports + composite briefing pack ship. Established by **ADR-0021**. the project scope (row 13).

> A pure serialisation layer over in-memory data already computed by the checkers. Adds "Copy for AI" buttons to five feature panels and a one-click "Copy Briefing Pack" composite to the Proof tab header. Exports structured JSON or Markdown — resolved, human-cited, before/after, with blind-spot disclosures — so a downstream contract-analyser LLM has clean priming context without the user having to reconstruct structure Mutatis has already computed precisely. Zero network calls; client-side only; one-time awareness notice.

---

## 1. Purpose & user value

Mutatis computes rich structured data on every analysis run: a definitions map, a clause-tree index, a classified redline change set, normalised deal dates and amounts, and a party/role cast list. Each of these is high-value priming context for a contract-analyser AI — but today all of it is trapped in the task pane.

A user who wants to "ask an AI about this contract" must copy raw document text and ask the AI to reconstruct structure that Mutatis has already computed precisely. The AI then has to re-parse definitions, re-infer party roles, and interpret strikethrough formatting to understand what changed.

**This feature makes those data sets portable** — one button per panel copies the relevant structured data; one composite button copies a complete briefing pack for the whole document. The AI gets a clean glossary, a deal-facts table, and plain before/after change pairs rather than raw redline markup.

**User value in one line:** "Give an AI the document structure Mutatis already computed, without the AI having to re-parse the document from scratch."

---

## 2. Scope

### 2.1 In scope (v1) ✅

**Tier-1 per-panel exports:**

| # | Export | Source data | Feature panel |
|---|--------|-------------|---------------|
| 1 | **Defined-terms glossary** | `DefinitionsMap` (feat. 01) | Defined Terms |
| 2 | **Section outline / clause tree** | `SectionIndex` (feat. 02) | Cross-Reference |
| 3 | **Change set with before/after** | `DiffOp.wordDiffs` (feat. 08/10) | Compare |
| 4 | **Deal-facts sheet** (dates, amounts, placeholders) | `DateRecord / AmountPairRecord / PlaceholderHit` (feat. 03) | Factual Consistency |
| 5 | **Party/role cast list** | `PartyAnchor / RoleRecord` (feat. 09) | Party-Role |

**Composite:**

| # | Export | Contents |
|---|--------|---------|
| 14 | **AI briefing pack** | Exports #1, #2, #4, #5 + open-items list (unfilled placeholders + undefined terms) + provenance header. One button, Proof tab header. |

**Cross-cutting capabilities:**
- JSON format (canonical) and Markdown format (paste-into-chat). Same data, two serialisers. User-selectable toggle persisted in `localStorage`.
- One-time awareness notice (toast, non-blocking) on first copy action per browser profile.
- Clipboard copy (primary) and file download (secondary — `Blob` + object URL → `<a download>`).
- Blind-spot disclosures on every export (see §4 Format rules / rule F6).

### 2.2 Deferred / cut 🔜❌

These are acknowledged, not silently omitted:

- 🔜 **#6 Cross-reference integrity report** — every prose ref + its resolution status. Deferred v2 (adds `ReferenceMatch[]` serialisation).
- 🔜 **#7 Open-items checklist (composite)** — merged unfilled placeholders + broken cross-refs + undefined-but-used terms. The placeholder + undefined-term slice is in the v1 briefing pack; the full merged list (including broken xrefs) is v2.
- 🔜 **#8 Change summary by category** — `ChangeSummary` counts + cascade annotations. v2; pairs with the full change-set export.
- 🔜 **#9 Extracted comments / negotiation markup** — `{author, clause, text}` per tracked comment. v2.
- 🔜 **#10 Filled template variables** — `AnswersMap` from the Templates feature. v2 (Templates not yet built).
- 🔜 **#11 Term-usage consistency map** — term → all usage sites with capitalisation drift flags. v2.
- 🔜 **#12 Clause-vs-precedent pair** — draft clause + matching approved library clause + drift. v2 (Clause Library not yet built).
- 🔜 **#13 Document hygiene / metadata report** — `RepairReport` metadata findings. v2; export carries an explicit "contains embedded metadata" warning.
- ❌ **App-initiated AI query** (the app sends data to an AI endpoint) — violates the confidentiality invariant. Non-negotiable. ADR-0021, ADR-0001/0003.
- ❌ **"AI summary" branding** — the app exports *data*; the AI generates the summary. Calling it a summary misrepresents what Mutatis does.
- ❌ **Exporting the on-screen filtered view** — a filtered export can silently omit a substantive item (ADR-0008). Full data set always; a future filtered-export option must print an explicit "N items hidden" banner.

---

## 3. Dependencies (consume; do not re-specify)

This feature issues **no** document-wide scan of its own. It is a pure serialisation layer over data that the source features have already computed and hold in memory. Build those first.

- **Feature 01** `src/features/defined-terms/` — `DefinitionsMap`, `UsageIndex`, `normaliseTerm` (export #1).
- **Feature 02** `src/features/cross-reference/` + `src/shared/section-index/` — `SectionIndex` with `SectionNode.displayLabel` (export #2).
- **Feature 03** `src/features/factual-consistency/` — `DateRecord[]`, `AmountPairRecord[]`, `PlaceholderHit[]` (export #4).
- **Feature 08/10** `src/features/compare/` — `DiffOp[]` with `wordDiffs`, `ChangeSummary` (export #3).
- **Feature 09** `src/features/party-role/` — `RoleRecord[]`, `PartyAnchor[]` (export #5).
- [Config layer](../ARCHITECTURE.md#localisation--config-layer) — `jurisdiction` value for the provenance header.
- Browser APIs: `navigator.clipboard.writeText`, `URL.createObjectURL`, `document.createElement('a').click()` — client-side only; no Office JS required for the copy/download action itself.
- `localStorage` — persistence for `ai-export.noticeShown` (boolean) and `ai-export.format` (`'json'|'markdown'`).

---

## 4. Data model

### 4.1 Top-level envelope

Every export — whether a single-panel export or the full briefing pack — is wrapped in this envelope.

```
AiExportEnvelope {
  _mutatis:  "1.0"                     // format version sentinel
  generated:    string                    // ISO-8601, e.g. "2026-06-09T14:32:00Z"
  document:     string                    // filename, or "Untitled" if unavailable
  jurisdiction: "SG" | "EW" | "unknown"  // from Config layer
  checks_run:   CheckId[]                 // which checkers contributed data this run
                                          // e.g. ["defined-terms", "factual", "compare"]
  blind_spots:  string[]                  // what was NOT covered — see F6
  [section]:    ExportSection             // one key per included export, named below
}

CheckId = "defined-terms" | "cross-reference" | "factual-consistency"
        | "compare" | "party-role"
```

**Blind spots (always populated):**
The `blind_spots` array is assembled from the `checks_run` list and a fixed lookup table:

| Not in `checks_run` | Blind-spot string |
|---|---|
| `"defined-terms"` | `"Defined-terms checker was not run; glossary and usage findings absent."` |
| `"cross-reference"` | `"Cross-reference / section-index checker was not run; clause-tree absent."` |
| `"factual-consistency"` | `"Factual-consistency checker was not run; deal dates, amounts, and placeholders absent."` |
| `"compare"` | `"Compare was not run; change set absent."` |
| `"party-role"` | `"Party-role checker was not run; cast list absent."` |

Additionally, regardless of what ran, always append:
- `"Headers, footers, and text boxes are not scanned."`
- `"Image-only content (scanned pages, embedded pictures) is not text-compared."`
- `"Cross-document and external references are detected but not resolved (Office JS cannot read other files)."`

### 4.2 Defined-terms glossary (export #1, section key: `"defined_terms"`)

```
DefinedTermsExport {
  glossary: TermEntry[]
}

TermEntry {
  term:          string          // canonical surface form, e.g. "Completion Date"
  definition:    string          // full resolved definition text
                                 // if isCrossRef: the referent clause text is appended inline:
                                 // "has the meaning given in Clause 14.2 [→ "means the date…"]"
  defined_at:    string          // human citation, e.g. "Clause 1.1" or "Schedule 1, row 3"
  source:        "prose"         // DefinitionRecord.source
               | "table"
               | "footnote"
  defined_multiple_times?: true  // omit key when false
  variants?:     string[]        // observed surface variants that differ from `term`
                                 // (lowercase uses, spelling drift) — omit key when empty
}
```

**Inline pointer resolution (rule F2):** if `DefinitionRecord.isCrossRef === true`, the export
fetches the referent `SectionNode` from the `SectionIndex` and appends its body text to `definition`,
formatted as `"<verb phrase> [→ "<resolved text>"]"`. If the referent cannot be resolved (node not
found, or body text empty), append `"[→ target not found in this document]"` instead of silently
omitting.

### 4.3 Section outline / clause tree (export #2, section key: `"section_outline"`)

```
SectionOutlineExport {
  clauses: ClauseNode[]
}

ClauseNode {
  number:    string         // SectionNode.displayLabel, e.g. "1.2.3" or "Schedule 2"
  heading:   string         // heading text, or "" if unnumbered
  level:     number         // 1-based nesting depth
  kind:      string         // SectionNode.kind, e.g. "clause" | "schedule" | "recital"
  children?: ClauseNode[]   // nested; omit key when empty
}
```

Full-text export of clause bodies is **not** included in v1 — the outline is the navigational skeleton. An LLM using the outline to direct its own reading of raw document text (via its own PDF/DOCX parser, or via the user providing raw text separately) benefits from the structured index without this feature needing to re-extract all paragraph text.

### 4.4 Change set with before/after (export #3, section key: `"change_set"`)

```
ChangeSetExport {
  baseline: string          // "original" document label (filename or user-supplied label)
  revised:  string          // "revised" document label
  changes:  ChangeEntry[]
}

ChangeEntry {
  id:             string    // DiffOp.id (stable)
  type:           "insert" | "delete" | "replace" | "move"
  classification: string    // DiffOp.classification — primary resolved class
                            // e.g. "obligationLanguage" | "numericalValue" | "formatting"
  clause:         string    // human citation from partId/SectionIndex, e.g. "Clause 4.2"
                            // or "Header" / "Footnote 3" for non-body parts
  original?:      string    // full original text of the changed span
                            // present for type: "delete" | "replace" | "move"
  revised?:       string    // full revised text of the changed span
                            // present for type: "insert" | "replace" | "move"
  move_target?:   string    // for type:"move" — target clause citation
  low_confidence?: true     // omit key when false
}
```

**Derivation of `original` / `revised` from `wordDiffs`:**
- `original` = concatenation of all `WordDiff` tokens where `op === 'del'` or `op === 'eq'` (for the context span of a replace).
- `revised`  = concatenation of all `WordDiff` tokens where `op === 'ins'` or `op === 'eq'`.
- For `type === 'delete'`: `original` = the full deleted paragraph text; `revised` is omitted.
- For `type === 'insert'`: `revised` = the full inserted paragraph text; `original` is omitted.
- For `type === 'replace'`: both present, each being the full reconstructed text for their side.
- For `type === 'move'`: `original` = the paragraph as it appeared at the source location; `revised` = the paragraph as it appears at the target. `move_target` carries the target citation.

This guarantees that no `ChangeEntry` requires the LLM to interpret formatting conventions to understand what changed.

### 4.5 Deal-facts sheet (export #4, section key: `"deal_facts"`)

```
DealFactsExport {
  dates:        DateEntry[]
  amounts:      AmountEntry[]
  placeholders: PlaceholderEntry[]
}

DateEntry {
  raw:             string          // "6 June 2026"
  iso:             string          // "2026-06-06"
  day_of_week:     string          // "Saturday"
  role:            string          // DateRecord.roleHint: "execution" | "completion" | "effective" | "asof" | "other"
  flagged_weekend: boolean
  clause:          string          // human citation
}

AmountEntry {
  raw:            string           // "S$150,000"
  value:          number
  currency:       string           // ISO 4217, e.g. "SGD" | "GBP"
  words_raw?:     string           // "One Hundred and Fifty Thousand Singapore Dollars"
  words_value?:   number
  mismatch?:      "value" | "currency" | "both"  // omit key when status === "match"
  clause:         string
}

PlaceholderEntry {
  raw:        string               // "[●]", "[TBC]", "[NTD: …]"
  category:   string               // PlaceholderHit.category
  clause:     string
}
```

### 4.6 Party/role cast list (export #5, section key: `"cast_list"`)

```
CastListExport {
  parties: PartyEntry[]
}

PartyEntry {
  role:           string           // canonical surface form, e.g. "Purchaser"
  entities:       string[]         // observed entity names associated with this role
                                   // e.g. ["ACME Pte. Ltd.", "ACME"]
  introduced_at?: string           // clause citation for the introduction parenthetical
                                   // omit if not formally introduced
  body_uses:      number           // number of body occurrences (excluding introduction)
  finding?:       "used-not-introduced" | "inconsistent-label"
                                   // present if the party-role checker flagged this entry
}
```

### 4.7 AI briefing pack (export #14, section key: `"briefing_pack"`)

The briefing pack is not a new data structure — it is the five sections above assembled under one envelope, plus a preamble.

```
BriefingPackExport extends AiExportEnvelope {
  preamble: {
    summary: string          // one-sentence auto-generated summary:
                             // "A {jurisdiction} agreement between {party roles joined},
                             //  dated {execution date if found}, containing {N} defined terms,
                             //  {N} clauses, {N} deal dates, and {N} open items."
    open_items: OpenItem[]   // unfilled placeholders + undefined-but-used terms
  }
  defined_terms:    DefinedTermsExport
  section_outline:  SectionOutlineExport
  deal_facts:       DealFactsExport
  cast_list:        CastListExport
}

OpenItem {
  type:    "placeholder" | "undefined-term"
  raw:     string
  clause:  string
}
```

The `preamble.summary` is assembled mechanically from the available data — it is not AI-generated and makes no claims beyond what the checkers found.

---

## 5. Format rules (mandatory)

These nine rules govern every serialiser. No serialiser may violate them. They derive from ADR-0021.

| # | Rule | Enforcement |
|---|------|-------------|
| **F1** | **Before/after, never markup.** Change exports emit `{original, revised}` string pairs. No strikethrough, no `<w:del>`, no caret notation. | `ChangeEntry` schema has no markup field. |
| **F2** | **Resolve pointers inline.** A definition that says "has the meaning given in Clause X" is exported with Clause X's body text appended. A cross-ref carries its target heading/text, not just the citation. | `TermEntry.definition` + inline resolver in defined-terms serialiser. |
| **F3** | **Cite in human terms, not GUIDs.** Every location exports as `"Clause 4.2"` / `"Schedule 2, para 3"` (from `SectionNode.displayLabel`). Never the internal `paraId`. | `clause` field on every entry uses `displayLabel` only. |
| **F4** | **Normalise and preserve raw.** Amounts: `{raw, value, currency}`. Dates: `{raw, iso, day_of_week}`. The AI gets clean machine values without losing how the text appears in the document. | `DateEntry` and `AmountEntry` schemas carry both forms. |
| **F5** | **Self-describe canonical vs surface variants.** Terms export their canonical form and any observed lowercase or drift variants in `variants[]`. | `TermEntry.variants` — see §4.2. |
| **F6** | **Blind-spot disclosures on every export.** `AiExportEnvelope.blind_spots[]` states what was not covered. | Assembler populates `blind_spots` from `checks_run` diff against the full checker set + the three fixed always-append entries. |
| **F7** | **Two serialisations, one schema.** The same export object is serialised by two formatters (JSON + Markdown). No data in one format absent from the other. | `src/features/ai-export/formats/json.ts` + `markdown.ts` both consume `AiExportEnvelope`. |
| **F8** | **Provenance header on every export.** `generated`, `document`, `jurisdiction`, `checks_run` always present. | `AiExportEnvelope` schema — all four fields required. |
| **F9** | **Lossless on substance.** Strip rsids, internal IDs, style names, and other structural plumbing. Never strip text that is semantically meaningful. | Serialisers pick named fields from source types; they do not JSON.stringify raw TypeScript objects. |

---

## 6. UI specification

### 6.1 Per-panel copy buttons

A **"Copy for AI"** button is added to each of the five feature panels, visible only when the panel has results to export (the source feature's last run returned data).

| Panel | Button placement | Data exported |
|-------|-----------------|---------------|
| Defined Terms | Below the findings list | Export #1 (glossary) |
| Cross-Reference | Below the findings list | Export #2 (section outline) |
| Factual Consistency | Below the findings list | Export #4 (deal facts) |
| Party-Role | Below the findings list | Export #5 (cast list) |
| Compare | Below the change-set summary | Export #3 (change set) |

> **Implementation note (Proof tab).** Factual Consistency and Party-Role have no distinct visual panel — they render in the shared findings list of the Proof tab — so "below the findings list" has no literal per-feature target there. The four Proof-tab exports (Defined Terms, Cross-Reference, Factual, Party-Role) are therefore gathered into one **"Copy for AI"** block at the foot of the Proof tab (`ProofExports`), each row still gated on its own source's results (the §6.1 visibility rule is preserved per-row). Compare keeps its own in-panel button (literal placement). The briefing-pack button sits in that block's header (§6.2).

Each button:
- Label: `Copy for AI` (or `Download for AI` for the secondary download action).
- Secondary action: a small download icon (`⬇`) beside the copy icon opens the file-download path. Clicking the main button copies to clipboard; clicking the icon downloads.
- Format toggle: a `JSON | MD` segmented control beside the button, persisted per user.
- Disabled state: greyed out with tooltip "Run [Feature Name] first" when the panel has no current results.

### 6.2 Briefing-pack composite button

A **"Copy Briefing Pack"** button lives in the Proof tab header (top-right of the Proof panel, beside the "Run All" button if one exists).

- Enabled only when at least one of the five source checkers has current results.
- Exports whichever Tier-1 sections have data; blind_spots covers the rest.
- Same clipboard/download + JSON/Markdown affordance as per-panel buttons.
- The one-time awareness notice fires on the first use of *any* copy action (briefing pack or per-panel), not on every click.

### 6.3 Awareness notice

A non-blocking toast (bottom of the task pane, auto-dismisses after 8 s, also has an ✕):

> **Copying document content** — This copies document content to your clipboard for use outside Mutatis. Mutatis makes no network calls. Your firm's policies on AI and data handling apply.

Fires once per browser profile (`localStorage` key `ai-export.noticeShown = true`). Not a modal. Does not block the copy action. No "I understand" button required.

---

## 7. Implementation approach

### 7.1 Module layout

```
src/features/ai-export/
  index.ts                  // public API: buildExport(), buildBriefingPack(), copy(), download()
  types.ts                  // AiExportEnvelope, ChangeEntry, TermEntry, … (all schemas from §4)
  serialise/
    defined-terms.ts        // DefinitionsMap → DefinedTermsExport
    section-outline.ts      // SectionIndex → SectionOutlineExport
    change-set.ts           // DiffOp[] → ChangeSetExport
    deal-facts.ts           // DateRecord[] + … → DealFactsExport
    cast-list.ts            // RoleRecord[] + PartyAnchor[] → CastListExport
  formats/
    json.ts                 // AiExportEnvelope → string (JSON.stringify + 2-space indent)
    markdown.ts             // AiExportEnvelope → string (Markdown tables + headings)
  assemble/
    briefing-pack.ts        // calls all serialisers, builds envelope + preamble
    blind-spots.ts          // CheckId[] → string[] (lookup table + three fixed entries)
  ui/
    CopyButton.tsx          // generic copy-to-clipboard + download button pair
    FormatToggle.tsx        // JSON | MD segmented control
    AwarenessNotice.tsx     // toast, fires-once logic
    BriefingPackButton.tsx  // composite button for the Proof tab header
```

### 7.2 Integration points (minimal changes to existing feature panels)

Each feature panel receives a `exportData` prop (the in-memory data structure it already holds). It renders a `<CopyButton data={exportData} serialise={serialiseDefinedTerms} />` below its results. No scanning code is added to any feature; the panel already has the data.

The `buildExport(checkId, data, config)` function in `index.ts` is the single entry point: it calls the correct serialiser, wraps in the envelope, and populates `checks_run` and `blind_spots`.

### 7.3 Clipboard + download

```
// clipboard (primary)
await navigator.clipboard.writeText(formatted)

// download (secondary)
const blob = new Blob([formatted], { type: format === 'json' ? 'application/json' : 'text/markdown' })
const url  = URL.createObjectURL(blob)
const a    = document.createElement('a')
a.href     = url
a.download = `mutatis-export-${checkId}-${isoDate}.${ext}`
a.click()
URL.revokeObjectURL(url)
```

Both are client-side only. No Office JS. No network.

---

## 8. Config keys

No new config keys are required for v1. The feature reads:
- `config.jurisdiction` — for `AiExportEnvelope.jurisdiction`.
- `localStorage.getItem('ai-export.noticeShown')` — awareness-notice flag.
- `localStorage.getItem('ai-export.format')` — `'json' | 'markdown'`, default `'markdown'`.

---

## 9. Tests

All tests are **unit tests** (Vitest + jsdom). No Office JS required (the feature makes no Office calls).

| Test file | What it covers |
|-----------|---------------|
| `tests/unit/ai-export/serialise/defined-terms.test.ts` | Given a mock `DefinitionsMap`: canonical term, cross-ref inline resolution, `defined_multiple_times` flag, `variants[]` from `UsageIndex`. |
| `tests/unit/ai-export/serialise/section-outline.test.ts` | Given a mock `SectionIndex`: nesting preserved, `displayLabel` used (not `paraId`), `kind` forwarded. |
| `tests/unit/ai-export/serialise/change-set.test.ts` | Given mock `DiffOp[]`: `replace` emits both `original` and `revised`; `delete` emits only `original`; `insert` emits only `revised`; `move` carries `move_target`; no markup chars in output. |
| `tests/unit/ai-export/serialise/deal-facts.test.ts` | Weekend date flagged; amount mismatch present; placeholder category forwarded; `iso` normalised correctly. |
| `tests/unit/ai-export/serialise/cast-list.test.ts` | Role emits `entities`, `body_uses`; `introduced_at` omitted when not introduced; `finding` present when party-role checker flagged. |
| `tests/unit/ai-export/formats/json.test.ts` | Valid JSON output; `_mutatis` sentinel present; `blind_spots` non-empty. |
| `tests/unit/ai-export/formats/markdown.test.ts` | Contains `##` headings per section; `blind_spots` in a `> ⚠` blockquote; no raw paraId strings; amounts formatted as table. |
| `tests/unit/ai-export/assemble/blind-spots.test.ts` | Missing checker → correct blind-spot string; three fixed always-present entries regardless of `checks_run`. |
| `tests/unit/ai-export/assemble/briefing-pack.test.ts` | Envelope `checks_run` matches contributed sections; `preamble.summary` assembles from available data; `open_items` includes both placeholder and undefined-term types. |
| `tests/unit/ai-export/ui/AwarenessNotice.test.tsx` | Shows on first copy; does not show on second copy (localStorage persisted); auto-dismisses; ✕ closes early. |
| `tests/unit/ai-export/ui/CopyButton.test.tsx` | `navigator.clipboard.writeText` called with formatted string; download `<a>` created with correct MIME type; button disabled when no data; format toggle switches serialisation. |

---

## 10. Cut / deferred items (cross-reference)

See §2.2 for the full deferred list. The canonical source of what is ❌/🔜 is **ADR-0021** + **the project scope (row 13)**. Do not treat the deferred items as out-of-scope for the product — only for this v1 build. The serialisation layer in `src/features/ai-export/serialise/` is designed to be extended one file at a time as v2 items are promoted.
