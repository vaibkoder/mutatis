# 01 — Defined Terms (Proof tab)

**Status:** ✅ v1 — all four sub-checks ship.

> Statically analyses the open document and surfaces four classes of defined-term defect (used-but-not-defined, defined-but-not-used, defined-multiple-times, capitalisation inconsistency) plus their locations, tuned for SG/UK commercial-contract drafting and deliberately conservative about false positives. Detect-and-navigate only; no auto-fix.

## 1. Purpose & user value

Defined terms are the backbone of a commercial contract, and a mis-managed definitions section is one of the most common and most legally consequential drafting defects: a term used but never defined creates ambiguity; a term defined twice creates internal contradiction; a defined term written in lowercase silently strips it of its bespoke meaning and reverts it to ordinary language.

This checker surfaces those four defects so a SG or UK (England & Wales) lawyer can fix them in seconds rather than by manual cross-referencing. It is tuned for commercial-contract conventions (double, often bold, quotation marks; a Clause 1 "Definitions and Interpretation"; interpretation schedules; footnote definitions) and is conservative because a checker that cries wolf on "the Buyer" or "Companies Act 2006" will be switched off.

**User value in one line:** "Show me every place my definitions are inconsistent, missing, redundant, or silently broken by lowercase — without flagging the things every contract legitimately leaves capitalised."

## 2. Scope

### 2.1 In scope (v1) — all four sub-checks ship ✅

| # | Sub-check | Default severity | Description |
|---|-----------|------------------|-------------|
| 1 | **Used-but-not-defined** | Critical | A capitalised term used in the body that has no definition anywhere (prose, table, footnote, or cross-reference), and is not on the exclusion allowlist. |
| 2 | **Defined-but-not-used** | Informational | A term that is defined but never referenced elsewhere in the document. |
| 3 | **Defined-multiple-times** | Critical | A term with more than one *definition* (not merely more than one use). |
| 4 | **Capitalisation inconsistency** | Informational | A defined term appearing in lowercase. Ships with a **smart/strict toggle** (smart = default). |

Plus the cross-cutting capabilities (all ✅ v1):

- **Footnote scanning** for definitions (footnote bodies — see [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index)).
- **Table-format definition** parsing (two-column term/definition tables) — see Clause 6.
- **Cross-reference definitions** treated as defined (`"Group" has the meaning given in Clause 14.2`).
- **Cross-document definitions** detected and *suppressed* (`"X" as defined in the [Other Agreement]`) — counted and reported as informational context, never resolved.

Nothing from the four sub-checks is descoped.

> **Recognition broadened (ADR-0013).** Beyond the quoted-term + default-verb model in Clause 5.1, v1 also recognises **colon-as-verb** definitions (`"Term": …`), **unquoted verb-anchored** definitions (`**Affiliate** means …`, `AFFILIATE means …` — the mandatory definitional verb is the false-positive control, since `ParaNode` exposes no run-level bold), and an **expanded verb set** (`shall have the meaning given/set out/specified/ascribed to it in`, `is defined in`, `shall mean and include`). Gated by `definedTerms.recogniseUnquoted` / `recogniseColon` (default on). Accepts a bounded false-positive increase by product-owner decision. See ADR-0013 and Clause 5.1.

### 2.2 Deferred / cut 🔜❌

These are acknowledged, not silently dropped:

- 🔜 **Multi-column (3+) definition tables** with separate numbering/term/definition columns. v1 handles the 2-logical-column case (a leading numbering column is tolerated — Clause 6.3); arbitrary layouts deferred.
- ❌ **Definitions embedded in images/OLE objects** (un-extractable text) — undetectable.
- ❌ **Resolving** a cross-document definition (opening the other agreement) — never in scope; detect-and-suppress only.
- 🔜 **Grammatical-correctness style nudges** beyond the single optional `shall mean` nudge (itself off by default).
- 🔜 **Auto-fix / one-click remediation** — v1 is detect-and-navigate only; suggested actions are textual.
- 🔜 **Acronym expansion validation** (e.g. checking that "SPA" was introduced as `(the "SPA")`) beyond treating the acronym as a defined term.

For the consolidated v1/v2/cut lists see [README.md](../README.md#feature-list).

## 3. Dependencies

Read these first; this feature consumes them and does not re-specify them.

- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — source of all paragraph text, footnote-body text, `listString`, paragraph style, and per-run bold/italic. Read-only here; this feature issues **no** document-wide scan of its own and reads the in-memory index. If the index already surfaces table-cell paragraphs tagged with table/row/cell coordinates, use that; Clause 6 assumes it may not and specifies a fallback walk.
- [Findings model](../ARCHITECTURE.md#findings-model) — every defect is emitted as a `Finding` (mapping in Clause 9.4); this feature owns no results UI.
- [Localisation / Config layer](../ARCHITECTURE.md#localisation--config-layer) — runtime carrier for the config keys in Clause 6.
- [Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) — footnote bodies (row 1), `listItem.listString` (row 2), the 255-char `body.search()` limit (row 8). This feature does **not** restate these; `body.search()` is used only for short jump-to-navigation re-location, never bulk scanning.
- Conventions: [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions) and [en-GB spelling](../CONVENTIONS.md#en-gb-spelling) supply quote style, definitional verbs, definition markers, the exclusion allowlist, and en-GB tolerance.

## 4. Data model

Two in-memory structures are built per analysis run and then discarded (no persistence).

### 4.1 Normalisation rules — `normaliseTerm(raw): string`

The canonical normaliser, applied to every candidate term before it is used as a key or looked up:

1. **Smart-quote / straight-quote fold.** Replace curly double `“ ”` → `"`, curly single `‘ ’` → `'`, and `« »` / `″ ′` similarly. Strip surrounding quote characters entirely from the *key* (the key stores the bare term, never the quotes).
2. **Trim & collapse whitespace.** Collapse internal runs of whitespace (incl. non-breaking space) to a single ASCII space; trim ends.
3. **Strip possessive.** Remove a trailing `'s` / `’s` (e.g. `"Buyer's"` → `Buyer`).
4. **Strip plural.** Remove a trailing `s` **only** when the singular form (or the term itself) is a known key — pluralisation is resolved at *lookup time* against the map, not destructively at storage time. Lookup tries exact, then possessive-stripped, then naive-singular `s`/`es`/`ies`→`y`. Never invent a plural that collides with a different real key.
5. **Case key.** The **map key** is the `toLowerCase()` of the result. The **display/canonical surface form** (original capitalisation, e.g. `Completion Date`) is stored alongside as `canonical`.
6. **Do NOT en-GB↔en-US fold the key.** `Authorised` and `Authorized` are *different keys* (per [en-GB spelling](../CONVENTIONS.md#en-gb-spelling)). en-GB tolerance governs spell-flagging in other checkers, not term identity here — but the capitalisation check must not treat `Organisation` vs `organisation` as a spelling inconsistency (it is purely a case comparison; the letters are identical).

**Longest-match-first for multi-word terms.** When scanning body text, multi-word defined terms must be matched greedily longest-first so that `"Completion Date"` is matched as one term and not double-counted as `"Completion"` + `"Date"`. The usage scanner builds an alternation of all defined surface forms sorted by descending length and consumes the longest match at each position. A use of "Closing Date" counts only toward `closing date`; a bare "Closing" counts toward `closing` (if defined) — see Clause 8.

### 4.2 `DefinitionsMap`

```
DefinitionsMap : Map<NormKey, DefinitionEntry>

DefinitionEntry {
  key:            string            // normalised lowercase key
  canonical:      string            // surface form as first/authoritatively defined, e.g. "Completion Date"
  definitions:    DefinitionRecord[]  // >1 ⇒ defined-multiple-times candidate
}

DefinitionRecord {
  source:         "prose" | "table" | "footnote"
  verb:           DefinitionalVerb | "cross-reference" | "cross-document" | null
  isCrossRef:     boolean           // "has the meaning given in Clause X" (same document)
  isCrossDoc:     boolean           // "as defined in the [Other Agreement]" — suppress, never resolve
  crossRefTarget: string | null     // e.g. "Clause 14.2" or the named other agreement
  location:       Range | LocationRef // for jump-to-navigation; Range when available
  styling:        { bold: boolean, italic: boolean }
  rawQuote:       "double" | "single" | "none"   // delimiter actually used
  contextScope:   string | null     // e.g. "for the purposes of Clause 7" (local redefinition) — see Clause 8
}
```

Multiple `DefinitionRecord`s under one key is the signal for sub-check 3. **Exception:** a record with `isCrossRef === true` plus a single substantive prose/table/footnote record is *not* a multiple-definition defect — the cross-reference and the substance are the same definition expressed twice. Two substantive records *are* a defect. A `contextScope`-bearing record (local redefinition) is excluded from the multiple-definition count when its scope differs (Clause 8).

### 4.3 `UsageIndex`

```
UsageIndex : Map<NormKey, UsageRecord[]>

UsageRecord {
  surface:    string          // exact surface text matched, e.g. "Completion Date" / "completion date"
  isLowercase: boolean        // surface form's first letter (and the term as a whole) is lowercase
  precededByDeterminer: boolean  // immediately preceded by the/a/an/this/that/such/its/their (for smart mode)
  isSentenceInitial: boolean  // first token of a sentence (capitalisation false-positive guard)
  location:   Range | LocationRef
  inFootnote: boolean
  isDefinitionOccurrence: boolean // TRUE for the occurrence inside its own definition — EXCLUDED from "used" counts
}
```

**Critical rule:** the occurrence of a term *inside its own definition* (the quoted term in `**"Completion Date"** means…`) is recorded with `isDefinitionOccurrence = true` and **excluded** from the "is this term used elsewhere?" computation. Otherwise every defined term would appear "used" and sub-check 2 would never fire.

### 4.4 How table & footnote definitions merge in

- **Footnote** definitions are produced by the same prose-pattern pass (Clause 5.2) run over footnote-body text; each `DefinitionRecord.source = "footnote"`.
- **Table** definitions are produced by the separate table path (Clause 6); `source = "table"`, `verb` may be `null` (the table structure *is* the verb).
- All three sources write into the **same** `DefinitionsMap` keyed by `normaliseTerm`. A term defined in a table *and* in prose is therefore correctly detected as defined-multiple-times (unless one is a cross-reference to the other).

## 5. Algorithm

```
1. Load config (Clause 6)      → quoteStyle, verbs, allowlist, capMode, enGBMode, styleNudges
2. buildDefinitionsMap()       → prose patterns + table cells + footnotes
3. buildUsageIndex()           → all capitalised/defined-term occurrences, excluding definition occurrences
4. runSubCheck1_usedNotDefined()
5. runSubCheck2_definedNotUsed()
6. runSubCheck3_definedMultipleTimes()
7. runSubCheck4_capitalisation()   (smart|strict)
8. emit Findings (grouped, Clause 7)
```

### 5.1 Build definitions map — definitional-verb + quote regex strategy

Definitions are recognised by the following structural shapes. Shapes A & B are anchored on a **quoted** candidate term; Shapes A′ & C broaden this per **ADR-0013** (colon-as-verb, and unquoted-but-verb-anchored) — the mandatory definitional verb is the false-positive control for the unquoted form.

**Shape A — quote-then-verb (the dominant form):**

```
("|'|“|‘) <TERM> ("|'|”|’)  [optional , or (]  <DEFINITIONAL VERB>
```

**Shape B — inline parenthetical:**

```
( [the|each|a|an]?  ("|'|“|‘) <TERM> ("|'|”|’) )
```

e.g. `(the "Disclosing Party")`, `(each a "Shareholder")`.

**Shape A′ — quote-then-colon (ADR-0013):**

```
("|'|“|‘) <TERM> ("|'|”|’)  :   <DEFINITION TEXT>
```

e.g. `"Business Day": any day on which banks in Singapore are open`. `verb = null` (the colon is the marker, as in tables). Gated by `definedTerms.recogniseColon` (default on).

**Shape C — unquoted, verb-anchored (ADR-0013):**

```
(sentence/paragraph start) <UNQUOTED TERM CANDIDATE> <DEFINITIONAL VERB>
```

e.g. `**Affiliate** means …`, `AFFILIATE means …`, `Affiliate has the meaning given to it in Clause 3`. The candidate is a 1–6-word Title-Case run or an ALL-CAPS acronym (≥2 chars); the **definitional verb is mandatory** and is the only false-positive control (`ParaNode` exposes no run-level bold). Guards (mandatory): the candidate must be **sentence-/paragraph-initial**; bare `means`/`includes`/`refers to` fire **only** sentence-initial, while the `…meaning…`-family and `shall mean` may fire either position; self-reference/allowlisted structural words are excluded (`This Clause means …`, `The Company means business …` must **not** flag). A candidate whose span is already covered by a quoted match (Shapes A/A′/B) is skipped. Gated by `definedTerms.recogniseUnquoted` (default on).

Regex construction notes (the verbs come from [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions), so the verb alternation is built at runtime, not literal):

- Quote class: `[\"'“”‘’]`. Opening and closing need not be the same glyph (Word autocorrect may pair `“…”`). Capture is non-greedy up to the next quote of the closing class.
- `<TERM>`: `[^\"'“”‘’]{1,120}` (cap length to avoid runaway). Must contain at least one letter. May be all-caps (acronyms), Title Case, or contain digits/`&`/`-` (`"S$"`, `"R&D Costs"`, `"Long-Stop Date"`).
- Verb alternation (default set from config): `means`, `shall mean`, `means and includes`, `includes`, `refers to`, `has the meaning given (to it )?in`, `has the meaning (set out|specified) in`. Case-insensitive, but require a word boundary so `meanwhile` ≠ `mean`.
- **Cross-reference detection:** if the matched verb is a `has the meaning given/set out in …` form, set `isCrossRef = true` and capture the target (`Clause \d+(\.\d+)*`, `Schedule \w+`, `paragraph …`). The definition is still recorded (the term IS defined) but with no local substance.
- **Cross-document detection:** if the trailing target text matches `(as |has the meaning |as defined )?in the <ProperNoun (Agreement|Deed|Document|SPA|…)>` referring to a *named external instrument* rather than a Clause/Schedule of *this* document, set `isCrossDoc = true`. These are suppressed from "used-but-not-defined" but reported in the informational banner (Clause 7.6). Boundary heuristic: a target containing `Clause`/`Schedule`/`paragraph`/`Annex` referencing this doc ⇒ cross-ref; a capitalised *named agreement* ⇒ cross-doc.
- **Styling capture:** read bold/italic for the term run from the Document Model. Styling is recorded but a definition is valid regardless of styling — bold and/or italic on a quoted term is a definition marker, never a defect (per [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions)).
- `shall mean` nudge: if config `styleNudges.shallMean === true` (off by default), additionally emit an Informational finding suggesting `means`. Recognition of `shall mean` as a valid verb is unconditional.

Each successful match produces a `DefinitionRecord` written into `DefinitionsMap` under `normaliseTerm(TERM)`.

### 5.2 Source passes feeding the map

1. **Prose pass** — run Shapes A & B over every body paragraph's text (from the Paragraph Index). Records `source="prose"`, `location` resolved to a Range, and `contextScope` if the sentence opens with `for the purposes of (this )?(Clause|Schedule|paragraph) <ref>` (Clause 8).
2. **Table pass** — Clause 6. Records `source="table"`.
3. **Footnote pass** — run Shapes A & B over each footnote body. Records `source="footnote"`, `inFootnote=true`.

Run order does not affect correctness (the map is order-independent); for stable `canonical`/first-location reporting, process prose → table → footnote and keep the earliest by document position.

### 5.3 Build usage index

Once `DefinitionsMap` is known, build the alternation of all `canonical` surface forms (and their config-tolerated variants) sorted longest-first (Clause 4.1). Scan **all** body paragraphs and footnote bodies:

- At each match, record a `UsageRecord`. Compute `precededByDeterminer` from the token immediately to the left (`the|a|an|this|that|these|those|such|its|their|any|each|no` — config-extensible). Compute `isSentenceInitial` (match starts a sentence: preceded only by `^`, `. `, `! `, `? `, `: `, newline, or bullet). Compute `isLowercase` (the matched surface begins lowercase).
- **Exclude definition occurrences:** pre-mark the character span of each definition's quoted term; any usage whose span intersects is the definition occurrence (`isDefinitionOccurrence = true`).

Lowercase occurrences for sub-check 4 come from the same scan; the alternation match is case-insensitive and `isLowercase` distinguishes them.

### 5.4 Sub-check 1 — Used-but-not-defined

A candidate is "used-but-not-defined" when:

- It is **capitalised** in the body (a token or run of Title-Case tokens; all-caps acronyms of length ≥2), AND
- `normaliseTerm(candidate)` is **absent** from `DefinitionsMap` (after plural/possessive lookup), AND
- it is **not** on the **exclusion allowlist** (party-roles, statutory/proper nouns, conventionally-capitalised terms — from [Defined-term conventions](../CONVENTIONS.md#defined-term-conventions)), AND
- it is **not sentence-initial-only** (Clause 8), AND
- it is **not** covered by a cross-document suppression (Clause 5.1).

Candidate harvesting uses a **multi-word greedy** rule: consecutive Title-Case tokens (optionally joined by `of`, `and`, `&`, `the`) form one candidate (`"Board of Directors"`, `"Memorandum and Articles"`). Statutory citations (`<Name> Act <Year>`, `<Name> Regulations <Year>`, `<Name> (Amendment) Act`) are recognised by pattern and routed to the allowlist logic, not flagged. Default severity **Critical**.

> Sub-check 1 is the highest false-positive risk; the allowlist and sentence-initial guard are mandatory, and a candidate must appear capitalised at least once in a non-sentence-initial position before flagging.

### 5.5 Sub-check 2 — Defined-but-not-used

For each key: count `UsageRecord`s with `isDefinitionOccurrence === false`. If zero, emit a **Defined-but-not-used** finding (Informational). A term imported by cross-reference but never used is still dead weight, so cross-reference and cross-document definitions are checked normally.

### 5.6 Sub-check 3 — Defined-multiple-times

For each key, count *substantive* `DefinitionRecord`s (exclude pure cross-reference records that duplicate a substantive one; exclude differently-scoped `contextScope` records). If ≥2 remain, emit **Defined-multiple-times** (Critical), listing every definition location. A term defined in both a Clause 1 table and restated in prose is the classic hit.

### 5.7 Sub-check 4 — Capitalisation inconsistency (smart | strict)

For each defined key, examine its lowercase `UsageRecord`s (`isLowercase === true`, `isDefinitionOccurrence === false`):

- **strict mode:** flag **every** lowercase occurrence of a defined term.
- **smart mode (DEFAULT):** flag a lowercase occurrence **only when** `precededByDeterminer === true` (e.g. "the completion date" where `"Completion Date"` is defined). This cuts false positives on defined terms that are also common nouns (a defined `"Property"` legitimately appears as "property" mid-sentence; "the property" is the strong signal it was meant as the defined term).

Guards (both modes inherit them):

- **Styling guard:** never flag a lowercase occurrence merely because the definition was bold/italic — styling is a definition marker, not a case signal.
- **en-GB guard:** the comparison is purely letter-case. `Organisation`/`organisation` is a valid hit; it is NOT a spelling inconsistency and must not be routed to any spell-flag (per [en-GB spelling](../CONVENTIONS.md#en-gb-spelling)).
- **Sentence-initial / heading guard:** a lowercase form cannot occur sentence-initially; a *capitalised* form at sentence start is not evidence of the canonical term (handled in Clause 5.4).

Default severity **Informational** (known false-positive ceiling — Clause 10).

## 6. Table-format definitions — ✅ included in v1

Many SG/UK contracts (especially schedules and term sheets) lay definitions out as a **two-column table**: left cell = the defined term (often bold, often quoted), right cell = the definition text. There is frequently no definitional verb — the table structure carries the meaning. These MUST be captured or sub-check 1 produces a flood of false "used-but-not-defined" errors.

### 6.1 Parsing path

```
for each table in body.tables (load: rows.cells.body.text, and font for col-0 runs)
  if isDefinitionTable(table):              // 6.2
     for each row:
        termCell = row.cells[firstContentCol]
        defnCell = row.cells[firstContentCol + 1]
        term = extractTerm(termCell)         // 6.3
        if term:
           write DefinitionRecord{ source:"table", verb:null,
                                   location: termCell range, styling, rawQuote }
```

If the Document Model surfaces table cells in the Paragraph Index (tagged with table/row/cell indices), read from there and skip the direct `body.tables` load. Otherwise issue a single batched load of `tables/rows/cells/body.text` + col-0 `font.bold`.

### 6.2 Detecting a definition table — `isDefinitionTable(table)`

Treated as a definition table when **all** hold:

1. It has **exactly 2 logical content columns** (a leading pure-numbering column — every col-0 cell matches `^\s*\d+(\.\d+)*\.?\s*$` — is tolerated and the content columns shift right by one; see 6.3).
2. The **left content column** cells are short (≤ 12 words, ≤ 80 chars) and predominantly Title-Case or quoted; the **right content column** cells are longer (sentence-like).
3. At least **2 rows** satisfy the term/definition shape (a 2-col payment schedule won't qualify).
4. Optional positive signal (raises confidence, not required): a header row containing `Term`/`Defined Term`/`Definition`/`Meaning`/`Interpretation`, or the table sitting under a heading whose text contains `Definitions`/`Interpretation` (the heading paragraph immediately preceding the table).

If ambiguous (fails 2 or 3 but the preceding heading strongly says "Definitions"), fall back to: treat left-cell as term **only if** it is quoted or fully Title-Case AND right-cell is non-empty. **Err toward capturing** (a missed definition table causes Critical false positives downstream, which is worse than a spurious capture that simply adds a harmless key).

### 6.3 `extractTerm(cell)` and column offset

- Determine `firstContentCol`: 0, unless every cell in column 0 across the table is a bare clause number (`^\s*\d+(\.\d+)*\.?\s*$`), in which case `firstContentCol = 1`.
- From the term cell text: strip a leading enumerator (`(a)`, `1.`, `(i)`); strip surrounding quotes; apply `normaliseTerm`. If the cell holds multiple bolded quoted terms (`"Buyer" and "Purchaser"`), split on `and`/`,` between quoted runs and create one record per term, all pointing at the same definition cell.
- Record `styling` from the col font (bold left-cell terms are the norm and are fine).

### 6.4 Merge

Each table term writes into the shared `DefinitionsMap` exactly as in Clause 4.4. Identical key from table + prose ⇒ candidate for sub-check 3 (subject to the cross-reference exemption).

## 7. UI / task-pane behaviour (Proof tab)

### 7.1 Entry & run

One card in the Proof tab. Runs on demand (and/or as part of "Run all Proof checks"). Computation is synchronous over the in-memory Paragraph Index after the Document Model's single load — sub-second on typical contracts.

### 7.2 Grouping — Critical vs Informational

Findings are grouped by severity:

- **Critical** — Used-but-not-defined (1), Defined-multiple-times (3).
- **Informational** — Defined-but-not-used (2), Capitalisation inconsistency (4), optional `shall mean` nudge (if enabled).

Within each group, ordered by document position.

### 7.3 Click-to-navigate

Every finding carries a `location` Range; clicking uses the [Findings model](../ARCHITECTURE.md#findings-model) jump-to-navigation. For multi-location findings (defined-multiple-times), expose each definition site as a sub-link. Range is resolved on click (lazy); if invalid, fall back to a short `body.search()` on the (short) term.

### 7.4 Mapping to the Findings model

| Sub-check | category | severity | suggested action (text) |
|-----------|----------|----------|--------------------------|
| 1 used-not-defined | `defined-terms/used-not-defined` | Critical | "Define '\<term\>' or, if it is an ordinary word, remove the capitalisation." |
| 2 defined-not-used | `defined-terms/defined-not-used` | Informational | "'\<term\>' is defined but never used — consider removing the definition." |
| 3 defined-multiple | `defined-terms/defined-multiple` | Critical | "'\<term\>' is defined N times (locations…). Keep one definition." |
| 4 capitalisation | `defined-terms/capitalisation` | Informational | "'\<surface\>' appears in lowercase but is a defined term — capitalise it or confirm ordinary use." |

### 7.5 Strict/smart toggle

A toggle on the card: **Smart (default) / Strict**. Changing it re-runs only sub-check 4 (cheap; maps already built) and updates the capitalisation group in place. Tooltip: "Smart flags lowercase uses of defined terms only after a determiner like 'the'. Strict flags every lowercase use."

### 7.6 Cross-reference / cross-document banner

A small informational line at the top of the card:

- "**N** term(s) defined by cross-reference within this document (treated as defined)." — links to each.
- "**M** term(s) defined by cross-reference to external documents (not checked)." — the `isCrossDoc` definitions; explicitly stated as recognised-but-not-resolved.

## 8. Edge cases

| Edge case | Handling |
|-----------|----------|
| **Sentence-initial capitalisation** ("Property is transferred…" at sentence start where "property" is ordinary) | A Title-Case candidate capitalised *only* at sentence-initial/heading positions is **not** flagged (sub-check 1). Require ≥1 mid-sentence capitalised occurrence before treating a token as a term. |
| **Multi-word partial match** ("Closing" vs "Closing Date") | Longest-match-first (Clause 4.1). "Closing Date" consumes both words; a bare "Closing" only matches `closing` if that is itself a key. Never increment the short key from a long-key match. |
| **Plurals / possessives** | Lookup-time singularisation + possessive stripping (Clause 4.1 rules 3–4). "Shareholders" → `shareholder`; "Buyer's" → `buyer`. Never auto-pluralise a key into a collision. |
| **Defined terms inside other definitions** | A term used *within another term's definition text* counts as a normal usage. Only the quoted headword of a definition is the `isDefinitionOccurrence`. Keeps nested usages out of "defined-but-not-used". |
| **Intentional clause-specific redefinition** ("for the purposes of Clause 7, X means…") | Captured with `contextScope`. Not counted as defined-multiple-times against the global definition (different scope). Reported only if `config.definedTerms.flagLocalRedefinitions` (default false). |
| **Cross-document ("as defined in the \[Other Agreement]")** | `isCrossDoc = true`; suppressed from used-but-not-defined; surfaced in the Clause 7.6 banner; **never resolved**. |
| **All-caps acronyms** ("the SPA", "VAT", "GST") | Treated as candidate terms (length ≥2, all-caps). Flagged used-but-not-defined only if not allowlisted and not introduced as `(the "SPA")`. Common ones (VAT, GST, PDPA, IRAS, ACRA, CPF) belong on the config allowlist. |
| **Smart-quote vs straight-quote** | Normalised in `normaliseTerm` rule 1 before any keying or matching. Definition recognition accepts mixed/curly pairs. |
| **Tracked-changes documents** | Office JS text reflects the host's current (displayed) revision view. v1 reads whatever the Document Model exposes; deleted-but-not-accepted definition text may not appear. Documented as known behaviour; no revision introspection in v1. |
| **Definition with no usage at all** (e.g. orphan schedule) | Still produces a key; surfaces under defined-but-not-used. Correct. |
| **Empty / no definitions found** | If `DefinitionsMap` is empty, suppress sub-checks 2–4 and report "No definitions detected" rather than flagging every capitalised word. Sub-check 1 runs with high suppression. |
| **S$ / £ amounts and symbols in terms** | `"S$"`, `"£"`, `"SGD"` may be defined. The term regex permits currency glyphs; `normaliseTerm` preserves them. Amounts *inside* definitions ("S$1,000,000") are definition text, not terms. |
| **Term appears in a heading** | Headings are paragraphs in the index; a Title-Case heading word is treated like sentence-initial (suppressed as candidate evidence). |
| **Performance / very large documents** | Single pass over the in-memory index; alternation regex built once; navigation ranges resolved lazily. See [Performance guidelines](../ARCHITECTURE.md#performance-guidelines). |

## 9. Test cases / acceptance criteria

Each case is a minimal fragment with the expected finding set (`**…**` = bold).

**T1 — SG/UK bold-double-quote definition, used correctly.** `1.1 **"Completion Date"** means the date on which Completion occurs. … on the Completion Date the Buyer shall pay.` (allowlist: "Buyer", "Completion"). → **no** findings; no double-quote error; "Buyer"/"Completion" not flagged.

**T2 — Used-but-not-defined.** `The Seller shall deliver the Disclosure Letter to the Buyer.` (no definition of "Disclosure Letter"; "Seller"/"Buyer" allowlisted). → **1 Critical** used-but-not-defined: "Disclosure Letter".

**T3 — Defined-but-not-used.** `**"Escrow Agent"** means Bank X.` and never appears again. → **1 Informational** defined-but-not-used.

**T4 — Defined-multiple-times.** Clause 1 table row `"Group" | the Company and its subsidiaries`; later prose `**"Group"** means the Company and each subsidiary undertaking.` → **1 Critical** defined-multiple-times (table cell + prose).

**T5 — Multiple-times exemption (cross-ref + substance).** Prose `**"Group"** has the meaning given in Clause 14.2.`; Clause 14.2 `In this Agreement, "Group" means …`. → **no** multiple-definition finding; banner: 1 cross-reference within document.

**T6 — Capitalisation, smart vs strict.** Defined `**"Property"**`; uses "the property is sold" (after determiner) and "all property and assets" (generic). → smart (default): **1 Informational** ("the property"). strict: **2 Informational**.

**T7 — Footnote definition.** Footnote body `**"Material Adverse Effect"** means …`; used in body. → recognised (source=footnote); no used-but-not-defined; no defined-not-used.

**T8 — Table-format (Schedule).** Schedule 1 heading "Definitions"; 2-col table with leading numbering column `1. | "Long-Stop Date" | 31 December 2026`. → recognised (firstContentCol=1); used elsewhere ⇒ no findings.

**T9 — Statutory / party-role suppression.** `Each Director must comply with the Companies Act 2006 and the Act.` → **no** used-but-not-defined for "Director", "Companies Act 2006", "the Act".

**T10 — Cross-document suppression.** `**"Permitted Encumbrances"** has the meaning given in the Facility Agreement.` (used later). → no used-but-not-defined; `isCrossDoc=true`; counted in the "external documents (not checked)" banner; never resolved.

**T11 — en-GB spelling.** Defined `**"Authorised Signatory"**`; body "each authorised signatory". → smart: **1 Informational** capitalisation; **no** spelling finding; "Authorised" never treated as a misspelling of "Authorized".

**T12 — Longest-match.** Defined both `"Closing"` and `"Closing Date"`; "on the Closing Date" and "at Closing". → "Closing Date" → `closing date`; "Closing" → `closing`. Both defined+used ⇒ no findings; no double counting.

**T13 — S$ currency term.** `**"S$"** means Singapore dollars.`; body "S$1,000,000". → "S$" recognised; the amount is body text, not a separate undefined term; no false used-but-not-defined.

**T14 — Sentence-initial false positive.** "Property" NOT defined; `Property of the Company. The property shall vest.` ("Property" only sentence-initial). → **no** used-but-not-defined.

**T15 — Empty definitions.** Capitalised words, zero detected definitions. → "No definitions detected"; sub-checks 2–4 suppressed; sub-check 1 suppressed-to-quiet.

**Acceptance (rollup):** all T1–T15 produce exactly the expected finding sets; no double-quote or bold-styling false positives; smart mode is default; the toggle re-runs only sub-check 4; footnote and table definitions merge into the same map; cross-document definitions are detected and never resolved.

## 10. Known limitations

- **Capitalisation false-positive ceiling (sub-check 4).** When a defined term is also a common noun ("Property", "Business", "Interest", "Default"), even smart mode occasionally flags a genuinely ordinary lowercase use, and *misses* an intended-defined lowercase use lacking a determiner. This is why sub-check 4 ships **Informational**. The ceiling cannot be fully removed without semantics; the toggle and determiner gate are the mitigations.
- **Sub-check 1 residual false positives.** Despite the allowlist, statutory pattern, and sentence-initial guard, novel proper nouns and product/company names may be flagged. The allowlist is config-extensible per matter.
- **Footnote API.** Footnote bodies are scanned; **endnotes** are a separate collection and are **not** scanned in v1 (gap — see [Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) row 1).
- **Tracked changes.** v1 analyses the displayed/final text only; pending deletions/insertions are not reconciled (Clause 8).
- **Table heuristics.** Unusual layouts (definition on the left, term on the right; or 3 content columns) may be missed; v1 captures the dominant 2-content-column form (with tolerated numbering column). Misses degrade to Critical noise, mitigated by "err toward capturing" (Clause 6.2).
- **No semantic resolution.** Cross-document and cross-reference substance is never opened/verified; the checker confirms a definition *exists*, not that it is *correct*.

## 11. Build milestones

1. **M1 — Normalisation & data structures.** `normaliseTerm`, `DefinitionsMap`, `UsageIndex` types. Unit-test quote folding, plural/possessive, longest-match alternation. (No Office JS yet.)
2. **M2 — Prose definition pass.** Shapes A & B over the Paragraph Index; cross-ref vs cross-doc classification. Test T1, T5, T10 as plain text.
3. **M3 — Usage index + sub-checks 1–3.** Definition-occurrence exclusion; used-not-defined (allowlist + sentence-initial + statutory guards), defined-not-used, defined-multiple-times. Test T2, T3, T4, T9, T12, T14, T15.
4. **M4 — Capitalisation sub-check 4 + smart/strict.** Determiner detection, styling/en-GB guards. Test T6, T11.
5. **M5 — Footnote pass.** Run prose patterns over footnote bodies. Test T7.
6. **M6 — Table-format pass.** `isDefinitionTable`, `extractTerm`, column-offset, merge. Test T8, T13.
7. **M7 — Findings & UI.** Map to Findings model; grouping; click-to-navigate (lazy range + search fallback); strict/smart toggle re-run; cross-ref/cross-doc banner (Clause 7.6).
8. **M8 — Config wiring & hardening.** Replace inlined defaults with reads from the Config layer (quoteStyle, verbs, exclusions, capMode, enGB, styleNudges, determiners); full T1–T15 acceptance pass; performance check on a >100-page contract.
