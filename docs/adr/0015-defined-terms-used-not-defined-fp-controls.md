# ADR-0015: Defined Terms — used-but-not-defined false-positive controls, all-caps recurrence rule, and shared exclusion lexicon

- **Status:** Accepted
- **Date:** 2026-06-07

## Context

Sub-check 1 of the Defined Terms checker (used-but-not-defined, Critical — `spec/features/01-defined-terms.md` §5.4) is the highest false-positive-risk check in the product. First live testing in Word Online surfaced a stream of false Criticals on text that is never a bespoke defined term:

- **Proper nouns** — country/territory names (`Singapore`, `Southeast Asia`), cities, months, weekdays, nationalities (`British`), and bare court references.
- **Company forms** — `LLC`, `LP`, `BV`, `AG`, etc., appearing inside party names.
- **Legal boilerplate all-caps** — `THEREFORE`, `WHEREAS`, `WITNESSETH`, `HEREBY`.
- **Unknown technical acronyms** — `FEED`, `LIBOR`, `SPA` — and compounds containing them (`Commencement of FEED`).

The pre-existing controls (the `exclusions` allowlist, the sentence-initial guard, the `≥1 mid-sentence occurrence` rule, the statutory-citation pattern) did not cover these classes.

A strategy review proposed several directions. The maintainer reviewed them and **selected a specific subset**; the rest were explicitly declined or deferred. This ADR records **only the decisions the maintainer proceeded with**. The declined/deferred options are listed under *Rejected alternatives* and *Consequences* so the reasoning survives and is not relitigated.

## Decision

### 1. Enumerated exclusion of proper nouns and boilerplate

Extend the `DefinedTermExclusions` allowlist with a new `properNouns` category (months, weekdays, countries/territories, major commercial cities, geographic regions, nationality/language adjectives, bare `Court`/`Courts`) and extend `acronyms` with company-form suffixes and legal-boilerplate all-caps words. A candidate whose normalised key (or any constituent all-caps token) is on the allowlist is suppressed from Sub-check 1.

### 2. All-caps recurrence rule (the one place a frequency threshold is used)

Replace the previous "suppress any candidate containing an undefined all-caps token" rule with a three-way classifier (`classifyAllCaps`, `src/features/defined-terms/subchecks.ts`):

- **defined** — an all-caps token that IS defined in the document takes precedence; the candidate is handled normally (defining `"FEED"` makes `Commencement of FEED` a normal candidate).
- **noise** — the all-caps token is undefined but allowlisted boilerplate/country (`THEREFORE`, `UAE`): **always suppressed**, even when recurring.
- **unknown** — the all-caps token is undefined and not allowlisted (`FEED`, `LIBOR`): **surfaced only when the candidate recurs (≥2 occurrences); a single occurrence is suppressed** as presumed noise.

The recurrence threshold applies **only to the `unknown` all-caps case**. It is deliberately *not* applied to ordinary (mixed-case) candidates — see decision 3.

### 3. Ordinary undefined terms are flagged on first occurrence; frequency is surfaced, not gated

A mixed-case undefined term (e.g. `Disclosure Letter`) is flagged on its **first** occurrence, regardless of how many times it appears. To help the user triage volume without losing any finding:

- Sub-check 1 clusters all occurrences of one term into a single finding (already the behaviour) and now records **every** occurrence anchor.
- The shared `Finding` model gains an optional additive field `occurrences?: FindingLocation[]` (set only when length > 1).
- The task pane shows an **occurrence counter** (`N×`), a **Next ▸ i/n** control to step through occurrences, and a **frequency sort toggle** (most-used first) on the Proof tab.

### 4. The exclusion lexicon is shared static data, single-sourced

The proper-noun / acronym / verb / role / boilerplate lexicons are **static English-language and global reference data, not per-matter configuration**. They live once in `src/shared/config/defined-term-lexicons.ts` and are composed into each `JurisdictionProfile`, rather than copied inline in `sg.ts` and `uk.ts`. Genuinely jurisdiction-specific entries (SG vs UK regulators / institutions — `MAS`/`SGX` vs `HMRC`/`FCA`, etc.) are kept in their own constants and spread on top of the shared base. Per-matter tuning remains possible by extending a profile's exclusion arrays at runtime.

## Rationale

- The maintainer is **unwilling to risk a recurrence/evidence gate on ordinary candidates** — for a checker whose entire value is catching a missing definition, silently dropping a genuine single-use undefined term is the worst failure. A visible false positive is recoverable (dismiss); a false negative is invisible. The counter + sort give the user frequency-based triage **without** hiding anything.
- All-caps tokens are the one class where the maintainer accepts "noise by default": unknown acronyms (`FEED`, `LIBOR`) are far more often stray technical shorthand than bespoke defined terms, so requiring recurrence before surfacing them is the right default — with the explicit escape hatch that defining the term, or its recurrence, brings it back.
- Enumeration is imperfect but immediate and zero-false-positive-risk (allowlisting only ever suppresses). It was chosen as the pragmatic first-line control for the live-test FPs.
- Co-locating the lexicon honours the project principles (two implementations of the same thing is a defect): the list was previously duplicated verbatim across both profiles.

## Rejected alternatives

- **Recurrence/evidence gate on all candidates (only flag a term used ≥2×).** Proposed in the strategy review; **explicitly declined by the maintainer** as too risky — it would hide genuine single-use undefined terms. Adopted the occurrence-counter + frequency-sort instead (decision 3), keeping every finding visible while still letting the user prioritise by frequency. The recurrence threshold is therefore confined to the `unknown` all-caps case only.
- **Remove the all-caps suppression entirely** (let the allowlist alone handle it). Rejected: the maintainer wants unknown all-caps treated as noise by default, with recurrence as the surfacing signal.
- **Replace enumeration with an additive, evidence-based candidate model** (flag only tokens with positive evidence — definition-section membership, definition-related paragraph style, determiner-stability, identical-recurrence). Considered and **deferred, not adopted** this round. These remain open follow-ons; they are not part of this decision and the checker remains subtractive (allowlist-based) for now.
- **Party-role consistency as a new checker** (flag a role defined as "Buyer" but later called "Purchaser"). Considered; **deferred** to a future feature with its own scope check — not built here.
- **Keep the lexicon duplicated inline per profile.** Rejected — the project principles defect.
- **Put the lexicon in per-matter config.** Rejected: it is static language data, not matter configuration. Matter-specific additions are still supported by extending a profile's arrays; the baseline is shared.

## Consequences

- **Accepted tradeoff (documented):** a genuinely-bespoke all-caps term that is **used only once and never defined** is now silently suppressed (the `unknown && <2` case). This is the maintainer's deliberate choice. Mitigations: it surfaces as soon as it recurs; defining it (any recognised form) restores normal handling; known-noise words are allowlisted so the rule is not over-broad.
- **No false-negative on ordinary terms:** mixed-case undefined terms are unaffected by any frequency gate — they flag on first use, as before.
- **Finding model:** `Finding.occurrences?: FindingLocation[]` is additive and optional; no other feature sets or depends on it. Consistent with `spec/features/01-defined-terms.md` §7.3 (multi-location findings).
- **Lexicon single-sourced:** `defined-term-lexicons.ts` is the one home; SG/UK splits live there. Future entries are added once.
- **Invariants intact:** all processing remains client-side (no network/ML); jurisdiction values remain config-driven (SG/UK only; the lexicon is global reference vocabulary used as an allowlist, not a hardcoded US convention); Sub-check 1 stays in-scope feature 01; no Compare suppression. the scope register row 1 unchanged.
- **Open follow-ons (deferred, not decided here):** evidence-based candidate signals (definition-section / paragraph-style / determiner-stability) and a party-role-consistency checker.
