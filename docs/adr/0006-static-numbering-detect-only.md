# ADR-0006: Static numbering — detect-and-flag, never auto-convert

- **Status:** Superseded by [ADR-0023](0023-numbering-upgrade-assisted-static-removal.md)
- **Date:** 2026-06-06

> **Superseded 2026-06-10 by [ADR-0023](0023-numbering-upgrade-assisted-static-removal.md).** The *one-shot OOXML `abstractNum`/`num` conversion* this ADR cut remains cut. What ADR-0023 reopens is narrower: an **assisted, tracked-changes, human-driven** path where the tool deletes hand-typed number prefixes as native Word revisions (typed-API, reviewable, rejectable) and the **user** then selects which cleared paragraphs to re-number with the existing apply engine. That defuses this ADR's stated fears (OOXML fragility, no undo, tool-made structural decisions). Read ADR-0023 before acting.

## Context

Legal documents often contain **manually-typed (static) clause numbers** — "1.", "1.1", "(a)" typed as literal text rather than Word's dynamic list numbering. Converting static numbering to dynamic Word list numbering would be valuable, but doing so reliably requires injecting and wiring `abstractNum`/`num` definitions in `numbering.xml` and reassigning every paragraph — fragile OOXML surgery that can corrupt a document, with no Office JS undo for the OOXML loop.

## Decision

The Numbering feature (and the Cross-Reference static-numbering flag) **detects and flags** manually-typed static numbering but **does not auto-convert** it to dynamic Word list numbering. Detection is the deliverable; remediation is left to the user.

## Rationale

Detection is safe and still valuable — it surfaces a real risk (static numbers don't renumber when clauses move). Auto-conversion is fragile enough to corrupt a legal document, which is an unacceptable outcome for this audience. This follows the constitution's "detect-and-flag over risky automation" principle.

## Rejected alternatives

- **Full static→dynamic conversion.** Rejected: corruption risk on real legal documents; the OOXML write path has no reliable undo, so a bad conversion is hard to recover from in-place.
- **Beta-flagged / opt-in conversion.** Rejected: still too risky for legal documents in v1 — a "beta" label does not make a corrupted contract acceptable, and supporting it would demand the same fragile OOXML injection.

## Consequences

- Users fix flagged numbering themselves; the product points precisely at it but does not rewrite it.
- The detection heuristic (leading-number regex distinguishing literal text from `listItem.listString` auto-numbers) is the shipped deliverable — see feature specs `02-cross-reference-and-static-numbering.md` and `06-numbering.md`.
- Statically-numbered units are still indexed by the Section Index (so they don't produce false "Not Found" cross-references).
- Reopening auto-conversion requires superseding this ADR with user approval.
