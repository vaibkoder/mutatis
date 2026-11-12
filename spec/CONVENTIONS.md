# CONVENTIONS — SG/UK Legal & Localisation (Config Layer)

> The **single source of truth** for Singapore and UK (England & Wales) drafting conventions — the concrete contents of the [Localisation / Config Layer](ARCHITECTURE.md#localisation--config-layer). Every feature reads these values from the central config; **no feature hardcodes US conventions**. Feature docs reference these headings instead of restating the values.
>
> See also [ARCHITECTURE.md](ARCHITECTURE.md) and [README.md](README.md).

---

## Jurisdiction Config Layer

All checkers, formatters, and structure builders read their conventions from one central `JurisdictionProfile` object. Two profiles ship: **`SG`** (Singapore) and **`UK`** (England & Wales). A default is selected at first run and persisted via Storage (`localStorage` for the user default; optionally `document.settings` to pin a profile to a document).

**Hard rules:**
- US-centric defaults are a **bug**: no `$`-as-USD default, no `mm/dd/yyyy`, no "Section" as the sole clause noun, no US quotation/punctuation rules, no "Inc./LLC/EIN/State-of-incorporation".
- **Singapore == England & Wales** for most drafting purposes (SG legal drafting derives from UK precedent / shared common-law heritage); the profiles diverge mainly in currency emphasis, entity suffixes, registration-number schemes, court formatting, and citation reporters.
- New conventions are added **to the profile**, not inline in a feature. Adding a third jurisdiction later means adding a profile, not editing checkers.

```ts
interface JurisdictionProfile {
  id: 'SG' | 'UK';
  terminology: { topLevelUnit: 'Clause'; scheduleWord: 'Schedule';
                 scheduleSubUnit: 'paragraph'; recitalMarker: 'Recitals' | 'Background'; /* … */ };
  currency: CurrencyTable;
  dateLocale: { order: 'little-endian'; defaultDisplayFormat: 'DD Month YYYY'; weekendDays: [0, 6]; /* … */ };
  definedTerms: DefinedTermConfig;
  placeholderVocab: PlaceholderConfig;
  numberingPresets: NumberingScheme[];
  entityFields: EntityFieldConfig;
  courtProfiles: StyleProfile[];
  spellingLocale: 'en-GB';
}
```

---

## Terminology

- **Clause** — the **default** top-level operative unit (**not** "Section"). "Section" / "Sec." / "§" are **recognised** (UK statutes, LMA/finance house styles) but never the default; treated as a top-level division only when a firm sets `section.isTopLevelDivision = true`.
- **Schedule** — the **default** annexure unit (**not** "Exhibit"). "Exhibit" / "Exh." is recognised but rare; never offered as default replacement text.
- **paragraph X of Schedule Y** — the canonical SG/UK way to reference inside a Schedule. Schedules divide into **paragraphs** (lowercase), not clauses. Compound references must validate both halves.
- **Annex / Annexure / Appendix / Part** — other divisions; "Part" is often roman (`Part I/II/III`), common in LMA/finance docs.
- **Recitals** — lettered **uppercase** in parentheses: `(A) (B) (C)`. **"Background"** is an accepted synonym heading.
- **Recognise-but-not-default list:** `Section`, `Sec.`, `§`, `Exhibit`, `Exh.` — matched so existing documents resolve, never emitted as house-style defaults.

```jsonc
{
  "units.primaryBody":     "Clause",
  "units.bodyAliases":     ["Clause","Cl","Section","Sec","§"],
  "units.annexurePrimary": "Schedule",
  "units.annexureAliases": ["Schedule","Sch","Exhibit","Exh"],
  "units.scheduleSubUnit": "paragraph",
  "units.otherDivisions":  ["Annex","Annexure","Appendix","Part"],
  "units.recitalNouns":    ["Recital","Background"],
  "units.recitalStyle":    "letteredUpper",
  "section.isTopLevelDivision": false,
  "selfReferenceTerms": ["this Agreement","this Deed","this Clause","this Schedule",
    "this Annex","this Appendix","this Part","this paragraph","hereof","herein","hereunder","the foregoing"]
}
```

---

## Defined-Term Conventions

- **Quote style:** **double quotes default**; **recognise single** too (and curly variants `" " ' '`). Double quotes are **NOT** an error in UK or SG defined-term drafting (the "UK = single quotes" rule is a prose/citation convention, not a defined-term-delimiter rule). Never flag double-quote usage.
- **Definition markers:** **bold and/or italic** on a quoted term is a **definition marker**, never a capitalisation/formatting anomaly. A definition is valid regardless of styling (bold-double-quoted is the SG/UK norm; absence of bold is not a defect).
- **Definitional verbs (default set):** `means`, `shall mean`, `means and includes`, `includes`, `refers to`, `has the meaning given (to it) in Clause X`, `has the meaning set out in`, `has the meaning specified in`. (`shall mean` is recognised unconditionally; an optional, default-**off** nudge can suggest `means`.)
- **Definition locations:** a dedicated **Clause 1 "Definitions and Interpretation"**; an **interpretation Schedule**; **inline** (`(the "Purchaser")`, `(each a "Shareholder")`); and **footnotes**. Two-column term/definition **tables** carry the meaning structurally (often no verb).
- **Capitalised-but-not-defined exclusion allowlist** (suppresses "used-but-not-defined" only; does not prevent pickup if actually defined):
  - **Party-role words:** the Buyer, Seller, Purchaser, Vendor, the Parties, Party, Disclosing Party, Receiving Party, Lender, Borrower, Guarantor, Company, Shareholder(s), Director(s).
  - **Statutory / proper nouns:** `<X> Act <Year>`, `<X> Regulations <Year>`, "the Act", "the Regulations", "the Companies Act"; SG (Companies Act 1967), UK (Companies Act 2006). Also caught by a statutory-citation pattern.
  - **Conventionally-capitalised:** Board, Business Day, Working Day, Director, Schedule, Clause, Recital, Annex, Appendix, Group, Effective Date; common acronyms VAT, GST, PDPA, IRAS, ACRA, CPF.
- Term identity does **not** fold en-GB↔en-US (`Authorised` ≠ `Authorized` as keys); the capitalisation check is a pure letter-case comparison and never routes to spell-flagging.

```jsonc
{
  "definedTerms.quoteStyle": "double",                // recognise single + curly
  "definedTerms.verbs": ["means","shall mean","means and includes","includes","refers to",
    "has the meaning given to it in","has the meaning given in","has the meaning set out in","has the meaning specified in"],
  "definedTerms.styleNudges.shallMean": false,
  "definedTerms.capMode": "smart",                    // smart | strict
  "definedTerms.determiners": ["the","a","an","this","that","these","those","such","its","their","any","each","no"],
  "definedTerms.exclusions": [ /* allowlist above + statutory regex family */ ]
}
```

---

## Currency

Recognise and pair: **`S$`/`SGD`** (SG), **`£`/`GBP`** (UK), **`$`/`USD`** (explicit US only), **`€`/`EUR`**, **`RM`/`MYR`**, **`HK$`/`HKD`**. Currency words: `Dollars`, `Pounds`/`Pounds Sterling`→GBP, `Euros`→EUR, `Ringgit`→MYR, `Cents` (sub-unit).

- **Bare `$` is AMBIGUOUS** — **never** auto-resolve to USD (in a Singapore document a bare `$` is frequently SGD). Bare-`$` raises a soft **Informational** "currency not specified" warning, never a Critical guess.
- **Strip currency words** (and trailing fillers `only`/`exactly`/`nett`/`and No/100 Cents`/`and Zero Cents`) before number parsing.
- **Tolerate the British inserted "and"** ("One Hundred **and** Fifty Thousand" = 150,000, not 100 + 50,000) when parsing spelled amounts; preserve the `and` that separates major and sub-units ("Dollars and Fifty Cents").
- **"Dollars" maps to a family**, not a single currency: assert a currency mismatch only across incompatible families (dollar vs pound/euro), not within the dollar family (S$/US$/HK$).
- **lakh/crore are not used** in SG/UK drafting — not parsed. Thousands `,` / decimal `.`. `centTolerance = 0.005`.

```jsonc
{
  "currencies": {
    "S$":"SGD","SGD":"SGD","£":"GBP","GBP":"GBP","Pounds":"GBP","Pounds Sterling":"GBP",
    "US$":"USD","USD":"USD","€":"EUR","EUR":"EUR","Euros":"EUR","RM":"MYR","MYR":"MYR","Ringgit":"MYR",
    "HK$":"HKD","HKD":"HKD","$":"AMBIGUOUS","Dollars":"DOLLAR_FAMILY"
  },
  "ambiguousDollarPolicy": "flag-soft",
  "centTolerance": 0.005
}
```

---

## Dates

- **DD Month YYYY, day-first** default (e.g. "6 June 2026"). `chrono-node` configured **little-endian**: numeric `6/6/2026` parses as 6 June, **never** US month-first.
- **Weekend detection:** `weekendDays = [0, 6]` (Sun/Sat). A stated **execution/completion** date on a weekend → Critical; a weekend **effective/"as of"** (retroactive) date is legitimate and not flagged.
- **Century-ambiguous** (2-digit year, `20__`) and **template** dates (`this ___ day of ____ 20__`) are **flagged, not guessed** (the latter surfaces via placeholder detection).
- **Anchor keywords** for the execution role: `executed`, `execution date`, `dated`, `signed`, `completion`, `completion date`, `shall complete`. Effective-date keywords: `with effect from`, `as of`, `effective date`, `effective as of`.
- **v2 — holiday calendars:** SG gazetted public holidays (Holidays Act, MOM) + E&W bank holidays (gov.uk; some movable). **Scotland and NI have different bank-holiday sets** and are **out of scope** (target E&W only). v2 also adds a **defined-"Business Day" cross-check** that reads the document's own definition and cross-checks performance dates against it rather than the generic weekend rule.

---

## Placeholders

- **Dominant marker:** `[●]` (**U+25CF**). Variants: `[•]` (U+2022), `[∙]` (U+2219), `[·]` (U+00B7), `[ ]` (empty), `[*]`, `[TBC]` (prioritised — more common than TBD in UK/SG), `[TBD]`, `[NTD: …]`, `[drafting note]`, `[DATE]`, `[NAME]`, `[AMOUNT]`, `[PARTY]`, `[ADDRESS]`, `[INSERT]`/`[insert …]`, `[●]%` (percent-bullet). High-confidence categories (bullet/keyword/insert-directive/note-to-draft) are default-on.
- **Whitelist (never flag):** `[INTENTIONALLY LEFT BLANK]`, `[Reserved]`.
- **Suppression rules (drop even if matched):**
  - **Case citations** — `[2024] SGHC 45`, `[2024] EWHC 123 (Comm)`, `[2023] SGCA 1`, `[2019] UKSC 7`: a 4-digit bracketed year followed by a court reporter/abbreviation (`EWHC, EWCA, UKSC, UKHL, SGHC, SGHC(A), SGCA, SGCA(I), SLR, WLR, AC, All ER, QB, Ch …`) or preceded by a `v`/`v.` case-name pattern.
  - **Signature lines** — underscore runs and bracket forms near `Signature`, `By:`, `Name:`, `Title:`, `Witnessed by`, `In the presence of` (same paragraph / preceding ~60 chars).
- **General bracketed optional text caution:** `[and the Guarantor]`, `[Singapore][London]` are legitimate optional-drafting brackets. The generic `bare-brackets` pattern is **opt-in, default off** (`flagBareBrackets = false`); even when on it is medium-confidence Informational only.

```jsonc
{
  "placeholderVocab.bulletChars": ["●","•","∙","·","*",""],
  "placeholderVocab.keywords": ["TBC","TBD","DATE","NAME","AMOUNT","PARTY","ADDRESS","INSERT","NTD","drafting note"],
  "placeholderWhitelist": ["[INTENTIONALLY LEFT BLANK]","[Reserved]"],
  "citationReporters": ["EWHC","EWCA","UKSC","UKHL","SGHC","SGHC(A)","SGCA","SGCA(I)","SLR","WLR","AC","All ER","QB","Ch"],
  "signatureAnchors": ["Signature","By:","Name:","Title:","Witnessed by","In the presence of"],
  "flagBareBrackets": false
}
```

---

## Numbering Presets

The numbering scheme schema (`NumberingScheme` / `NumberingLevel`, with `setLevelNumbering` formatStrings using `%1..%9` tokens and points-based indents) is detailed in the Numbering feature doc. The **jurisdiction defaults** owned here:

- **UK/SG Commercial — DEFAULT** (`jurisdiction: "SG-UK"`): hierarchy **`1 / 1.1 / 1.1.1 / (a) / (i) / (A)`**, lower-roman and letters in **closed parentheses** — `(a)`, `(i)`, `(A)`, **never** `a.`, `a)`, `i)`. Top-level clauses run continuously (no reset); deeper levels reset under their immediate parent. Unit = **"clause"**.
- **Schedule preset:** `1 / 1.1 / (a)`, unit = **"paragraph"**, often grouped under **"Part 1 / Part 2"** headings (a `preset-schedule-part` variant sets level 0 to `Part %1`).
- **Recital preset:** single-level **`(A) (B) (C)`** (parenthesised upper-letter), unit = "recital".
- **Plain Outline (memo):** `1 / 1.1 / 1.1.1 …` decimal, no parenthesised limbs.
- **US litigation style** (`I. / A. / 1. / a.`) is **non-default**, clearly labelled "US style — not for SG/UK drafting", and **suppressed entirely** in SG/UK config unless the user enables "show non-SG/UK presets".

**Cross-reference regex vocabulary** (compiled from the terminology config): default unit "Clause", sub-units `(a)`→`(i)`, ranges/lists connectors (`to`, `–`, `—`, `-`; `,`, `and`, `&`), `(inclusive)` marker consumed-and-ignored. "Section"/"Exhibit" recognised, never default.

---

## Entity & Template Fields

- **Registration numbers:**
  - **UK Companies House number** — **8 characters**, two-letter prefixes allowed (`SC` Scotland, `NI`, `OC`/`SO` for LLPs). Validator `UK_CH`: `^[A-Z0-9]{8}$`.
  - **Singapore UEN** — validator `SG_UEN` (business `nnnnnnnnX`, local company `yyyynnnnnX`, others `Tyypqnnnnx`); pattern set held in config, not hardcoded.
- **Entity-type suffixes (jurisdiction-aware):** **UK:** `Limited`, `Ltd`, `plc`, `LLP`. **Singapore:** `Pte. Ltd.`, `Ltd.`, `LLP`. **NO** `Inc.`, `LLC`, `EIN`, "State of incorporation" — forbidden in option lists and field templates.
- **Standard variables:** registered office address; **jurisdiction of incorporation** (E&W / Singapore / other); **Business Day location** (e.g. "London" / "Singapore").
- **Execution / attestation blocks (shipped templates):**
  - **UK deed — Companies Act 2006 s.44:** signed by **two authorised signatories** (two directors, or a director + company secretary), **or** a **single director whose signature is witnessed**. Header "**EXECUTED as a DEED**"; correct witness attestation wording ("in the presence of …") for the director+witness variant. Ship both valid s.44 variants.
  - **Singapore execution block:** execution per the company's constitution / common seal; director + secretary or two directors as applicable; Singapore-appropriate attestation wording.
- Dates in template content controls default to **DD Month YYYY**, never month-first (ISO accepted as input, rendered day-first).

---

## Court Formatting Profiles

Style profiles (typography, margins, spacing; points-based, `1 mm = 2.83464567 pt`) for the Repair feature. **Normalisation is opt-in, never silent**; orientation section breaks are **protected**.

- **SG Court Filing (PD 2021)** — Practice Directions 2021: **Times New Roman 12** body; **TNR 10** quoted authorities / footnotes; **double spacing** (`lineSpacingRule: multiple`, `lineSpacing: 2.0`); **≥35 mm all margins** (`99.21 pt`); A4 portrait; **page numbers top-right** (flagged, not auto-inserted in v1); paragraph cross-citation in **square brackets `[xx]`**. When this profile is active and the document already matches, spacing/margin normalisation checkboxes default **unchecked with a warning**.
- **UK Court Filing** — TNR 12; **1.5 line spacing** (court bundles commonly require ≥1.5; double also accepted); ~25 mm margins (≈70.87 pt), optional larger left binding margin; A4 portrait; page numbers bottom-centre or top-right (flagged); neutral-citation conventions. Court-bundle requirements vary by court/division — the profile is a sensible default, firms clone.
- **Firm Commercial** — generic non-court default: TNR 11; single spacing with modest `spaceAfter`; 25.4 mm (72 pt) margins; portrait, landscape Schedules preserved; normalisation most useful here.

**Mandatory behaviours:** never apply single-spacing under a court profile; warn if a commercial profile is selected on a court-looking document (double-spaced + ≥35 mm + TNR12 heuristic); never strip an **orientation** section break without explicit per-break confirmation (legitimate landscape Schedules); shipped profiles are `locked` — user clones to edit.

---

## en-GB Spelling

- Both jurisdictions draft in **en-GB**. Checkers must be **en-GB tolerant**: `authorise`/`organisation` are correct, not misspellings; the defined-term capitalisation check treats `Authorised`/`Organisation` as pure case comparisons, never spell-flags.
- **Compare** surfaces en-GB↔en-US spelling changes (`organize→organise`, `-ize→-ise`, `color→colour`, `-or→-our`, `program→programme`) as a **real text delta** — classified `substantiveText` with a `spellingVariant` sub-flag, **never** "formatting-only". Detection via a small bidirectional en-GB/en-US variant map in config; the rule errs toward visibility.
- Clause/library language facet is `en-GB` / `en-SG`; `en-US` spelling variants are not promoted.
