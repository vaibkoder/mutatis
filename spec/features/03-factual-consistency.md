# 03 — Factual Consistency Checks (Proof tab)

**Status:** ✅ v1 — three deterministic sub-checks (weekend execution dates, number/word mismatch, placeholder detection).

> Three deterministic, fully-local sweeps over the document body that catch self-contradicting clerical facts surviving substantive review: an amount in words disagreeing with its parenthetical figure, an "execution date" landing on a weekend, or an unfilled placeholder escaping into an executed deed. All findings are **explainable** (every Finding cites the exact text it matched) and **conservative** (where genuinely ambiguous — a bare `$`, a century-ambiguous date — it flags rather than guesses, and downgrades severity rather than emitting a false Critical).

## 1. Purpose & user value

Legal documents routinely contain self-contradicting facts that are clerical, not conceptual: a one-digit slip between "Ten Thousand Dollars (S$10,000)" and "(S$1,000)" is a material drafting error; an unfilled placeholder in a registered instrument can require rectification. These checks live on the **Proof tab** and run three sweeps:

1. **Weekend execution-date warning** — stated execution/completion dates that fall on a Saturday or Sunday.
2. **Number/word mismatch** — spelled-out monetary amounts vs their parenthetical numeric figures, and currency words vs currency symbols.
3. **Placeholder detection** — unfinished drafting markers (`[●]`, `[TBC]`, `[insert ...]`, `[NTD: ...]`) while suppressing valid conventions and case citations.

## 2. Scope

### 2.1 In scope (v1) ✅

| Sub-check | Emits | Default severity |
|---|---|---|
| Weekend execution-date warning | One Finding per stated execution/completion date landing on Sat/Sun | **Critical** |
| Number/word mismatch (amount) | One Finding per word↔figure disagreement | **Critical** |
| Currency symbol/word mismatch | One Finding per currency disagreement (number may still match) | **Critical** |
| Bare-`$` ambiguity | One Finding per amount pair whose currency cannot be resolved | **Informational** (softer) |
| Placeholder detection | One Finding per placeholder hit (after whitelist/suppression) | **Informational** |

### 2.2 Deferred to v2 🔜

- **Public-holiday calendars** and the **defined-"Business Day" cross-check** — v1 detects weekends only and does **not** read the definitions clause. The v2 holiday calendars (SG gazetted holidays; E&W bank holidays — Scotland/NI out of scope) and the bespoke Business-Day cross-check are specified in [Dates conventions](../CONVENTIONS.md#dates); see [README.md](../README.md#deferred-to-v2-).
- 🔜 **Non-monetary number/word checks** (e.g. "thirty (30) days"). v1 restricts number/word matching to amounts carrying a currency indicator, to control false positives against structural identifiers (Clause 8).

### 2.3 Out of scope ❌ (all versions unless re-scoped)

- US date conventions (Month-Day-Year) — the product is SG/UK only (per [Dates conventions](../CONVENTIONS.md#dates)).
- Indian lakh/crore numbering — confirmed not relevant to SG/UK drafting; not parsed (per [Currency conventions](../CONVENTIONS.md#currency)).

## 3. Dependencies

Read these first.

- [Document Model / Paragraph Index](../ARCHITECTURE.md#document-model--paragraph-index) — ordered paragraphs with normalised plain text, stable IDs, and a `(paragraphId, charOffset) → Word.Range` resolver. All three sub-checks operate over this text and use the resolver for navigation, avoiding both the 255-char `search()` limit and per-check OOXML re-walks.
- [Findings model](../ARCHITECTURE.md#findings-model) — the canonical `Finding` shape; this feature emits into it.
- [Localisation / Config layer](../ARCHITECTURE.md#localisation--config-layer) — runtime carrier for the config keys in Clause 6.
- [Office JS constraints reference](../ARCHITECTURE.md#office-js-constraints-reference) — the 255-char `body.search()` limit (row 8): `search()` is used only for short *anchored* lookups (locating "executed"/"completion" near a date), never to match a whole amount-pair span; long-span matching is done over Document-Model text with offsets mapped back to ranges.
- Conventions: [Currency](../CONVENTIONS.md#currency), [Dates](../CONVENTIONS.md#dates), [Placeholders](../CONVENTIONS.md#placeholders) — currency table + ambiguous-`$` policy, date locale + weekend/anchor keywords, placeholder vocabulary + suppression rules. **All literals are config-driven, not hardcoded.**
- Libraries: **`chrono-node`** (configured little-endian/day-first) and **`words-to-numbers`** (wrapped — Clause 5.2) — see [Technology stack](../ARCHITECTURE.md#technology-stack). No network calls; both bundled.

## 4. Data model

All records carry an `anchor` (`{ paragraphId, startOffset, endOffset }`) resolvable to a `Word.Range`. Offsets are into the **normalised paragraph text** produced by the Document Model.

### 4.1 `DateRecord`

```ts
interface DateRecord {
  anchor: TextAnchor;
  rawText: string;          // exact matched substring, e.g. "6 June 2026"
  parsed: Date | null;      // chrono little-endian result; null if unresolved
  format: 'dd-month-yyyy' | 'numeric-slash' | 'numeric-dash' | 'ordinal-long' | 'template' | 'unknown';
  centuryAmbiguous: boolean;// true for 2-digit years ("20__", "06")
  isTemplate: boolean;      // partially-filled "this ___ day of ____ 20__"
  roleHint: 'execution' | 'completion' | 'effective' | 'asof' | 'other';
                            // derived from nearby anchor words
  dayOfWeek: 0|1|2|3|4|5|6 | null; // parsed?.getDay(); 0=Sun, 6=Sat
}
```

### 4.2 `AmountPairRecord`

```ts
interface AmountPairRecord {
  anchor: TextAnchor;
  rawText: string;              // full matched span
  wordsText: string;            // "One Hundred and Fifty Thousand Dollars"
  wordsCurrency: CurrencyCode | 'AMBIGUOUS' | null; // from "Dollars"/"Pounds"/...
  wordsValue: number | null;    // normalised numeric value of the words
  figureText: string;           // "S$150,000" / "£10,000.00"
  figureCurrency: CurrencyCode | 'AMBIGUOUS' | null; // from symbol/code
  figureValue: number | null;   // 150000 / 10000.00
  isRange: boolean;             // "between X and Y"
  qualifier: 'up-to' | 'not-less-than' | 'at-least' | 'approx' | null;
  status: 'match' | 'value-mismatch' | 'currency-mismatch'
        | 'both-mismatch' | 'ambiguous-currency' | 'unparseable';
}
```

### 4.3 `PlaceholderHit`

```ts
interface PlaceholderHit {
  anchor: TextAnchor;
  rawText: string;              // "[●]", "[NTD: confirm completion date]"
  category: 'filled-bullet' | 'keyword' | 'note-to-draft' | 'insert-directive'
          | 'bare-brackets' | 'percent-bullet';
  suppressed: boolean;          // matched a suppression rule
  suppressReason?: 'whitelist' | 'case-citation' | 'signature-line';
  confidence: 'high' | 'medium'; // bullet/keyword = high; bare = medium
}
```

Findings are produced only from **non-suppressed** records.

## 5. Algorithm

Each sub-check: **extract (regex/scan) → parse → compare → emit Finding**, all over the per-paragraph normalised text from the Document Model (one shared pass).

### 5.1 Weekend execution-date check

**Extract.**
1. Multi-format date regex sweep per paragraph:
   - `dd-month-yyyy`: `\b(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|…|December)\s+(\d{4})\b`
   - `ordinal-long` (deed style): `this\s+(\d{1,2}|[A-Za-z]+)(?:st|nd|rd|th)?\s+day\s+of\s+(January|…|December)\s*,?\s+(\d{4}|\d{2})` and the **template** variant `this\s+_+\s+day\s+of\s+_+\s+20_*`
   - `numeric-slash` / `numeric-dash`: `\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\b`
2. Pass each raw span to **chrono-node** configured little-endian. Numeric `d/m/y` is therefore day-first by construction.

**Role hint.** Within the same paragraph (and the immediately preceding clause), search the Paragraph Index text for execution anchor keywords (from [Dates conventions](../CONVENTIONS.md#dates)): `executed`, `execution date`, `dated`, `signed`, `completion`, `completion date`, `shall complete`. Within a configurable window (default 120 chars) of the date, set `roleHint`. Effective-date language (`with effect from`, `as of`, `effective date`) sets `roleHint='effective'/'asof'` — reported differently (Clause 8, retroactive dates).

**Compare.**
- `parsed` null → no Finding (placeholder/template handling covers `20__`).
- `centuryAmbiguous` → **flag, do not compute weekend** (Informational "century-ambiguous date — cannot verify day of week").
- Else `dayOfWeek = parsed.getDay()`; **weekend iff `=== 0 || === 6`** (`weekendDays = [0,6]`).

**Emit.**
- `roleHint ∈ {execution, completion}` **and** weekend → **Critical**: *"Stated {execution|completion} date {rawText} falls on a {Saturday|Sunday}."*
- `roleHint = other` and weekend → **Informational** only if config `flagAllWeekendDates` is on (default **off**).
- Retroactive `effective`/`asof` weekend dates → not flagged as execution errors (Clause 8).

### 5.2 Number/word mismatch (amount) check

**Extract the pair.** The canonical SG/UK pattern is *words → currency word → parenthetical figure*: `Ten Thousand Dollars (S$10,000)`, `Pounds Sterling £10,000`, `One Hundred and Fifty Thousand Dollars (S$150,000.00)`.

Primary regex (illustrative; the production version is assembled from the config currency table):

```
(?<words>(?:[A-Za-z]+[\s-]+){1,12}?(?<wcur>Dollars?|Pounds?(?:\s+Sterling)?|Euros?|Ringgit|Cents?))   // spelled amount + currency word
\s*
(?:\(|—|–|-|,|:)?\s*                                                                                   // optional opener
(?<fig>(?<fcur>S\$|SGD|£|GBP|US\$|USD|€|EUR|RM|HK\$|\$)\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)              // numeric with currency
\)?
```

Because a full pair can exceed 255 chars in pathological cases, matching is done over Document-Model text, **not** `body.search()`.

**Parse the words side (the wrapped `words-to-numbers` pipeline):**
1. **Lower-case & strip trailing fillers:** `only`, `exactly`, `nett`, `(and) No/100 Cents`, `and Zero Cents`, `and …/100`. Regex strip e.g. `\band\s+(no|zero|\d{1,2})\s*\/\s*100(\s+cents?)?\b` and `\b(only|exactly|nett)\b$`.
2. **Capture & strip the currency word** → `wordsCurrency` via the config currency-word map (`Dollars`→ resolve by symbol later / `AMBIGUOUS`; `Pounds`/`Pounds Sterling`→GBP; `Euros`→EUR; `Ringgit`→MYR; `Cents`→sub-unit handling). See [Currency conventions](../CONVENTIONS.md#currency).
3. **Sub-unit handling:** if the words contain `… Dollars and Fifty Cents`, split on `and`-before-`Cents`, convert the cents fragment separately, combine as `dollars + cents/100`.
4. **British "and" normalisation:** the inserted connective "and" between hundreds and the remainder ("One Hundred **and** Fifty Thousand") is *grammatical*, not additive. Pre-normalise by removing `and` tokens between number-words (preserving the `and` separating major and sub-units, handled in step 3). This is the wrapper stock `words-to-numbers` lacks.
5. Feed the cleaned words to **`words-to-numbers`** → `wordsValue`.

**Parse the figure side:**
1. Resolve `figureCurrency` from the leading symbol/code via the config currency table. **Bare `$` → `AMBIGUOUS`** (do *not* assume USD; in a Singapore document a bare `$` is frequently SGD — per [Currency conventions](../CONVENTIONS.md#currency)).
2. Strip the currency token and thousands commas; parse the remainder as a float (`,`=thousands, `.`=decimal; lakh/crore not handled by design).

**Compare:**
- **Value:** `|wordsValue − figureValue| ≤ centTolerance` (default `0.005`). Outside tolerance → value mismatch.
- **Currency agreement:** both resolve and **disagree** → currency mismatch. *Note:* "Dollars" alone is under-specified (S$/US$/HK$); assert a mismatch only when the word maps to a *family* incompatible with the symbol (e.g. "Pounds (€10,000)" → mismatch; "Dollars (S$…)" → consistent dollar family → no mismatch). Either side `AMBIGUOUS` (bare `$`) → **do not** raise a currency-mismatch Critical; raise the softer ambiguity Informational instead.
- Combine into `status`.

**Emit:**
- `value-mismatch`/`both-mismatch` → **Critical**: *"Amount in words ({wordsText} = {wordsValue}) does not match the figure {figureText} ({figureValue})."* Suggested-fix payload offers the figure that matches the words and vice-versa (user chooses; never auto-edit).
- `currency-mismatch` (value matches) → **Critical**: *"Currency words ({wordsCurrency}) disagree with figure currency ({figureCurrency})."*
- `ambiguous-currency` → **Informational**: *"Bare '$' — currency not specified (S$/US$/HK$?). Consider qualifying."*
- `unparseable` → no Finding (silence beats noise).
- **Ranges / qualifiers:** if `isRange` or `qualifier` set, the words side is usually descriptive prose without a parenthetical-figure echo ⇒ typically no pair to compare; if a pair *is* present per endpoint, compare each endpoint independently (Clause 8).

### 5.3 Placeholder detection sweep

A **multi-pattern** scan, ordered high-confidence → low-confidence, with whitelist and suppression applied **before** emitting.

**Pattern set (config-driven `placeholderVocab` — see [Placeholders conventions](../CONVENTIONS.md#placeholders)):**

| Category | Pattern (illustrative) | Confidence |
|---|---|---|
| filled-bullet | `\[\s*[●•∙·]\s*\]` (U+25CF dominant; variants U+2022, U+2219, U+00B7) | high |
| filled-bullet variants | `\[\s*\]` (empty), `\[\s*\*\s*\]` (`[*]`) | high |
| percent-bullet | `\[\s*[●•]\s*\]\s*%` (`[●]%`) | high |
| keyword | `\[(TBC|TBD|DATE|NAME|AMOUNT|PARTY|ADDRESS)\]` (case-insensitive; **`[TBC]` prioritised — more common in UK than TBD**) | high |
| insert-directive | `\[(INSERT|insert)\b[^\]]*\]` (`[INSERT]`, `[insert date of completion]`) | high |
| note-to-draft | `\[(NTD|drafting note|note)\b[^\]]*\]` (`[NTD: …]`, `[drafting note]`) | high |
| bare-brackets | `\[[^\]]{1,80}\]` (generic) | **medium — heavily suppressed** |

**Whitelist (never flag):** `[INTENTIONALLY LEFT BLANK]`, `[Reserved]` (config `placeholderWhitelist`).

**Suppression rules (drop the hit even if matched):**
1. **Case citations.** A 4-digit bracketed year `\[(19|20)\d{2}\]` immediately followed by a court reporter/abbreviation (`EWHC|EWCA|UKSC|UKHL|SGHC|SGCA|SGHC\(A\)|SGCA\(I\)|SLR|WLR|AC|All ER|…` — config `citationReporters`), or preceded by a case-name pattern (`v\.?`/`v\b` between proper nouns) — a neutral citation, not a placeholder.
2. **Signature lines.** Underscore runs and bracket forms near `Signature`, `By:`, `Name:`, `Title:`, `Witnessed by`, `In the presence of` (config `signatureAnchors`; default window: same paragraph or preceding 60 chars).
3. **Whitelist** (above).

**False-positive control.** The generic `bare-brackets` pattern is the main FP risk because UK/SG use square brackets around *optional/conditional* text (`[and the Guarantor]`, `[Singapore][London]`):
- `bare-brackets` hits are emitted **only if** config `flagBareBrackets` is enabled (**default off**), and even then at **medium** confidence and **Informational** severity.
- `filled-bullet`, `keyword`, `insert-directive`, `note-to-draft` are the **default-on** categories.

**Emit:** one **Informational** Finding per non-suppressed hit: *"Unfinished placeholder {rawText} ({category})."* Dedupe key `(paragraphId, startOffset, rawText)` so identical repeated placeholders each surface once at their own location.

## 6. SG/UK config keys used

This feature reads the following from the [Localisation / Config layer](../ARCHITECTURE.md#localisation--config-layer); the *values* are owned by CONVENTIONS.md.

- **Currency** (`config.currencies`, `config.ambiguousDollarPolicy = 'flag-soft'`, `config.centTolerance = 0.005`) — per [Currency conventions](../CONVENTIONS.md#currency). Bare `$` → `AMBIGUOUS`, never auto-resolved to USD; soft Informational only. "Dollars" maps to a family; mismatch asserted only across incompatible families.
- **Date locale** (`config.dateLocale`: `order='little-endian'`, `defaultDisplayFormat='DD Month YYYY'`, `executionAnchorKeywords`, `effectiveDateKeywords`, `flagAllWeekendDates=false`, `weekendDays=[0,6]`) — per [Dates conventions](../CONVENTIONS.md#dates).
- **Placeholder vocabulary** (`config.placeholderVocab`: `bulletChars`, `keywords` (TBC prioritised), `placeholderWhitelist`, `citationReporters`, `signatureAnchors`, `flagBareBrackets=false`) — per [Placeholders conventions](../CONVENTIONS.md#placeholders).

The SG profile and the UK (E&W) profile differ only in default currency emphasis and citation-reporter sets; both ship enabled.

## 7. UI / task-pane behaviour (Proof tab)

### 7.1 Grouping & severity

Findings render grouped by **severity**, then sub-check:
- **Critical** — number/word **value** mismatches; currency word/symbol mismatches; weekend **execution/completion** dates.
- **Informational** — placeholder hits; bare-`$` ambiguity; century-ambiguous dates.

Each group is collapsible with a count badge. Critical expanded by default; Informational collapsed.

### 7.2 Click-to-navigate

Each Finding exposes `rawText` and a **"Go to"** affordance that resolves the `anchor` → `Word.Range`, calls `range.select()` + `range.scrollIntoView()`. No edits on navigation.

### 7.3 Suggested fixes (non-destructive)

For amount mismatches, two **read-only** suggestions ("figure to match words: S$10,000" / "words to match figure: One Thousand Dollars"). Applied manually or via an explicit Apply button (tracked-changes-friendly); the feature **never silently rewrites** a legal amount.

### 7.4 Ambiguity surfaced softly

Bare-`$` and century-ambiguous items use a distinct **amber "Check"** styling (not red "Error"), with copy that *asks* rather than *asserts*: "Currency not specified — is this S$, US$ or HK$?"

### 7.5 Re-run model

The sweep runs on demand ("Run Proof checks") and is debounced on document change if live-mode is enabled. All three sub-checks share one pass over the Document-Model text; results merge into the Findings model in a single update.

## 8. Edge cases

| # | Case | Handling |
|---|---|---|
| 1 | **Partially-filled date template** "this ___ day of ____ 20__" | `format='template'`, `isTemplate=true`. Not weekend-checked. Surfaces via **placeholder** detection as Informational "unfinished date". |
| 2 | **Retroactive "as of" effective date** differing from execution date | `roleHint='effective'/'asof'`. Weekend rule **not** applied as Critical (back-dated effective dates legitimately fall on weekends). Divergence from the execution date is intentional in practice; no contradiction asserted. |
| 3 | **Century-ambiguous numeric date** ("6/6/06", "20__") | `centuryAmbiguous=true`. **Flag, do not guess** — Informational "cannot verify day of week"; no weekend Critical. |
| 4 | **Range** "between S$10,000 and S$20,000" | `isRange=true`. Per-endpoint pairs compared independently; bare prose → no comparison. |
| 5 | **Qualifiers** "up to S$10,000", "not less than £5,000", "at least" | `qualifier` captured. Figure still compared to accompanying words like-for-like; the qualifier alone does not trigger a Finding. |
| 6 | **Currency mismatch without value mismatch** "Ten Thousand Dollars (€10,000)" | Value matches but `wordsCurrency=dollar-family` vs `figureCurrency=EUR` → **currency-mismatch Critical**. Value agreement alone is insufficient. |
| 7 | **Numbers as structural identifiers** "Clause Three (3)", "Schedule Two (2)" | **Excluded** — the extractor requires a **currency indicator** on the figure side. No currency → not an `AmountPairRecord`. |
| 8 | **Non-monetary word/number pairs** "thirty (30) days" | Out of scope v1 (no currency indicator); not flagged. |
| 9 | **"Dollars" without a symbol** "Ten Thousand Dollars (10,000)" | `figureCurrency=null`; value still compared; currency agreement skipped. Bare `$` → ambiguity Informational. |
| 10 | **Defined "Business Day"** in the document | v1 ignores the definitions clause and uses the generic weekend rule. Recorded for **v2** cross-check (Clause 2.2). |
| 11 | **Sub-unit amounts** "One Thousand Dollars and Fifty Cents (S$1,000.50)" | Cents fragment parsed separately and combined; compared within `centTolerance`. |
| 12 | **"and No/100" / "only" / "exactly" suffixes** | Stripped before `words-to-numbers`; clean compare. |
| 13 | **British inserted "and"** "One Hundred and Fifty Thousand" | Pre-normaliser removes the grammatical "and" so the wrapped `words-to-numbers` yields 150,000 (not 100 + 50,000). |
| 14 | **Case citation in brackets** "[2024] SGHC 45", "[2024] EWHC 123 (Comm)" | Suppressed by citation rule; **not** a placeholder. |
| 15 | **Signature-line blanks** "Name: ____" / "By: [   ]" | Suppressed by signature-anchor proximity. |
| 16 | **Optional-text brackets** "[and the Guarantor]" | Not flagged by default (`flagBareBrackets=false`); generic bare-brackets pattern opt-in only. |

## 9. Test cases / acceptance criteria

Samples use SG and UK forms. Each asserts the expected Findings.

### 9.1 Number/word — value
| Input | Expected |
|---|---|
| `Ten Thousand Dollars (S$10,000)` | **No Finding**. |
| `Ten Thousand Dollars (S$1,000)` | **Critical** value-mismatch. |
| `One Hundred and Fifty Thousand Dollars (S$150,000)` | **No Finding** (British "and"). |
| `Pounds Sterling £10,000` | **No Finding** (GBP↔£). |
| `Five Thousand Pounds (£5,000.00)` | **No Finding** (cent tolerance absorbs `.00`). |
| `One Thousand Dollars and Fifty Cents (S$1,000.50)` | **No Finding** (sub-unit). |
| `Ten Thousand Dollars and No/100 (S$10,000) only` | **No Finding** (suffixes stripped). |

### 9.2 Number/word — currency
| Input | Expected |
|---|---|
| `Ten Thousand Dollars (€10,000)` | **Critical** currency-mismatch (value OK, dollar vs EUR). |
| `Ten Thousand Dollars ($10,000)` | **Informational** ambiguous-`$` (not Critical). |
| `Ten Thousand Pounds (£10,000)` | **No Finding**. |

### 9.3 Weekend execution dates (verify with a real calendar)
| Input | Expected |
|---|---|
| `This Agreement is executed on 6 June 2026.` (6 Jun 2026 = **Saturday**) | **Critical** weekend execution date. |
| `Completion shall take place on 7 June 2026.` (Sunday) | **Critical** weekend completion date. |
| `Executed on 5 June 2026.` (Friday) | **No Finding**. |
| `With effect from 7 June 2026.` (Sunday, effective date) | **No Critical** (effective date). |
| `Dated this 6th day of June 2026.` (Saturday, deed form) | **Critical** weekend execution date. |
| `Dated 6/6/2026.` | Parsed **day-first** → 6 June 2026 (Saturday) → **Critical** (never US M/D, never April). |
| `Dated this ___ day of ____ 20__.` | **No weekend Finding**; Informational template placeholder. |

### 9.4 Placeholders
| Input | Expected |
|---|---|
| `The Purchase Price is [●].` | **Informational** filled-bullet. |
| `Completion Date: [TBC]` | **Informational** keyword. |
| `[insert date of completion]` | **Informational** insert-directive. |
| `[NTD: confirm with client]` | **Informational** note-to-draft. |
| `Interest at [●]% per annum.` | **Informational** percent-bullet. |
| `This page is [INTENTIONALLY LEFT BLANK].` | **No Finding** (whitelist). |
| `[Reserved]` | **No Finding** (whitelist). |
| `As held in Lim v Tan [2024] SGHC 45 …` | **No Finding** (citation suppressed). |
| `… [2024] EWHC 123 (Comm) …` | **No Finding** (citation suppressed). |
| `Name: ______  By: [   ]` (signature block) | **No Finding** (signature suppression). |
| `[and the Guarantor]` (optional text) | **No Finding** by default (`flagBareBrackets=false`). |

### 9.5 Cross-cutting acceptance
- AC-1: No Finding ever auto-edits the document.
- AC-2: Bare `$` never produces a Critical.
- AC-3: Numeric dates are always day-first; no US Month-Day-Year interpretation.
- AC-4: `Clause Three (3)` and `thirty (30) days` produce **no** amount Finding.
- AC-5: Every Finding navigates to the exact source substring via the Paragraph Index resolver.
- AC-6: Whole-document run on a 100-page deed completes in one Document-Model pass (no per-check re-walk) and never relies on `search()` for span matching.

## 10. Known limitations

1. **No holiday calendars in v1.** 🔜 A weekday execution date that is a Singapore public holiday or an E&W bank holiday is not flagged. Deferred to v2 (Clause 2.2); Scotland/NI out of scope.
2. **No Business-Day cross-check.** 🔜 The document's own defined "Business Day" is ignored in v1; the generic weekend rule may diverge from a bespoke definition.
3. **Semantic ambiguity of bracketed optional text.** Generic `[...]` is legitimately used for optional drafting; the feature deliberately under-reports bare brackets (opt-in, off by default), at the cost of missing some genuinely unfinished bare-bracket markers.
4. **"Dollars" is under-specified.** It maps to a *family*, not a single currency; a mismatch is asserted only across incompatible families (per [Currency conventions](../CONVENTIONS.md#currency)).
5. **Century-ambiguous and template dates are reported, not resolved.** No guessing of century or completion of blanks.
6. **English-language amounts/dates only.** Non-English spelled-out numbers are not parsed.
7. **No cross-reference reconciliation.** v1 compares a words↔figure pair *locally*; it does not check an amount used consistently elsewhere.

## 11. Build milestones

| Milestone | Deliverable | Exit criteria |
|---|---|---|
| **M0 — Scaffolding** | Proof-tab sub-check registration; shared single-pass driver over Document-Model text; Findings emit path; config wiring. | Empty sweep runs end-to-end, emits zero Findings, navigates a stub Finding. |
| **M1 — Placeholders** | Multi-pattern sweep, whitelist, citation + signature suppression, severity = Informational. | §9.4 cases pass; `[2024] SGHC 45` not flagged; bare-bracket default-off. |
| **M2 — Amounts** | Pair extractor; words-to-numbers wrapper (British "and", currency-word strip, suffix strip, sub-units); currency table; bare-`$` soft path; cent tolerance. | §9.1 & §9.2 pass; structural identifiers excluded (AC-4). |
| **M3 — Dates/Weekend** | Multi-format date sweep; chrono little-endian; role-hint anchoring; weekend logic; century-ambiguous + template handling. | §9.3 pass; AC-3 holds; templates → placeholder not weekend. |
| **M4 — UI & Severity** | Findings grouping, click-to-navigate, suggested-fix payloads, amber "Check" styling for ambiguity. | §7 behaviours verified; AC-1, AC-2, AC-5 hold. |
| **M5 — Hardening** | Performance pass (single-walk, 255-char compliance), config-profile switching (SG/UK), edge-case suite. | §8 table green; AC-6 holds on a 100-page document. |
| **M6 (v2 prep, non-blocking)** | Interfaces/stubs for holiday-calendar provider and Business-Day definition reader. | Stubs typed and injectable; no behaviour change in v1. |
