# ADR-0002: Singapore & England-and-Wales jurisdiction focus

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

Legal drafting conventions are jurisdiction-specific: terminology, citation systems, date order, currency assumptions, entity suffixes, and numbering all differ between common-law markets. Most commercial proofing/compare add-ins are US-first, which makes them subtly wrong for non-US lawyers (defaulting `$` to USD, month-first dates, "Section/Exhibit" as the only structural nouns). Mutatis targets a specific audience and must fit it precisely rather than broadly.

## Decision

Target **Singapore and England & Wales only**. US conventions are **never hardcoded** anywhere in the codebase; US forms are *recognised* where sensible but never defaulted to. All locale-specific values come from the Config Layer (`shared/config`), whose data lives in `spec/CONVENTIONS.md`. Two profiles ship: `SG` and `UK`.

## Rationale

SG and E&W drafting derive from shared common-law / UK precedent, so one well-tuned convention set covers both with minor divergence (currency emphasis, entity suffixes, registration numbers, court formatting). Tuning precisely for this audience is the product's differentiator. A US-centric default is treated as a bug, not a feature gap.

## Rejected alternatives

- **US-first.** Rejected: wrong audience entirely; would default the product against its own users (USD `$`, mm/dd/yyyy, Section/Exhibit) and require them to fight the tool.
- **US-included (support SG/UK *and* US).** Rejected: US drafting uses different terminology (Section/Exhibit vs Clause/Schedule), citation systems (Bluebook vs neutral citations/OSCOLA), date order, and currency assumptions. Supporting it would dilute SG/UK fit and add a large surface — citation tooling, US entity fields, US numbering outlines — for an audience the product deliberately is not for. US citation features are consequently cut (ADR-0005).

## Consequences

- Defaults are **Clause / Schedule / paragraph**, **DD Month YYYY** (day-first), **S$ / £**, **en-GB** spelling, parenthesised `(a)/(i)/(A)` numbering limbs.
- US-style citation, Table of Authorities, and Bluebook/OSCOLA tooling are cut (ADR-0005).
- Adding a third jurisdiction later means **adding a profile**, not editing checkers — no convention may be hardcoded inline in a feature.
