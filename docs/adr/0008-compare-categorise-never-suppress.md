# ADR-0008: Compare — categorise, never suppress

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

The Compare feature authors a Word-vs-Word OOXML redline and summarises the changes. A large redline between two contract versions is noisy: formatting tweaks, renumbering cascades, and substantive edits all appear together. It is tempting to reduce noise by hiding "unimportant" changes — but in a legal redline, a change the tool *hides* is a change the lawyer never sees. Wrongly suppressing one substantive edit is a malpractice-grade error.

## Decision

Compare **classifies and lets the user filter** changes — formatting vs substantive, renumbering-cascade collapse, en-GB↔en-US spelling-variant sub-flag — but **never hides a change**. There is **no ML / semantic-equivalence judgement**. Classification is **rule-based** only. Every change is always reachable in the UI; any "formatting-only" view is a reversible filter, never a silent drop.

## Rationale

A false positive that suppresses a substantive edit from a legal redline is unacceptable — the cost of missing a real change vastly outweighs the convenience of a quieter list. Rule-based categorisation gives the noise-reduction benefit (filter, collapse, group) while keeping every change visible and the classification auditable. This encodes the constitution's "categorise, never suppress" principle.

## Rejected alternatives

- **ML semantic-equivalence detection ("this change doesn't matter, hide it").** Rejected: unacceptable suppression risk — a model deciding a substantive edit is equivalent and dropping it from the redline is a malpractice exposure no convenience justifies. Also opaque/unauditable.
- **Hiding "formatting-only" changes outright.** Rejected: even a correct formatting classification must be a *reversible filter*, never a silent removal — the user must always be able to reveal it. A hidden change the user can't reach is the same failure mode in a milder form.

## Consequences

- Classification is rule-based and auditable; the en-GB↔en-US spelling map errs toward visibility (`substantiveText` + `spellingVariant` sub-flag, never "formatting-only").
- All changes are always reachable in the UI; filters and cascade-collapse are presentation only.
- No model inference path exists in Compare; reopening semantic suppression requires superseding this ADR with user approval.
