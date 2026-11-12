# ADR-0016: Revive a narrowed Party-Role Consistency checker

- **Status:** Accepted
- **Date:** 2026-06-07
- **Supersedes:** the party-consistency clause of ADR-0005 only

## Context

ADR-0005 cut an item named "**Execution & Party Consistency Check**" with the
single rationale "removed by the maintainer", and the scope register lists it on the ❌
cut table. That name bundled **two distinct capabilities**:
1. **Execution / attestation-block validation** — checking that a signature
   block satisfies e.g. Companies Act 2006 s.44 (two signatories, or director +
   witness) or the Singapore constitution/common-seal rules.
2. **Party naming consistency** — checking that each party is referred to by one
   consistent label throughout.

Three things have changed since the cut:
- **Live testing of feature 01 (Defined Terms)** confirmed that precedent-copy
  slips — a party introduced as "the Buyer" but later called "the Purchaser",
  and a role like "the Agent" used throughout but never actually defined — are a
  real, high-frequency, legally-consequential defect class. Defined-Terms
  deliberately does **not** catch these: sub-check 1 allowlists party-role words
  to avoid false-positive floods (ADR-0015), so a party-role used-but-undefined
  or referred-to-by-two-labels passes silently.
- **The harvesting machinery now exists.** The Defined-Terms `DefinitionsMap` +
  usage scanner, the SG/UK party-role lexicon (`defined-term-lexicons.ts`), and
  the shared Section Index are all built and client-side. The naming-consistency
  half is now cheap to build by **reuse**, not the from-scratch effort imagined
  when it was cut.
- ADR-0015 explicitly logged "party-role consistency as a new checker" as a
  **deferred** follow-on; the maintainer now wants it.

The maintainer wants to revive **only the party-naming-consistency half**, narrowed
further to **label consistency**. The execution/attestation-block half stays cut.

## Decision

Supersede the "Execution & Party Consistency Check" cut of ADR-0005 **only insofar
as it cut party-role naming consistency**. Build a new v1 Proof-tab checker —
**Party-Role Consistency** — with this exact scope:

**In scope (THIS checker):**
- **(a) Inconsistent label / same party, two names.** A single party referred to
  by two or more party-role labels that the document's own text shows denote the
  **same** party (the classic Buyer/Purchaser precedent-copy slip). Flagged
  **only on explicit textual evidence** of co-reference (see Rationale) — never
  on mere co-presence of two role words.
- **(b) Role used-but-never-defined-as-a-party.** A party-role label (Buyer,
  Seller, Purchaser, Vendor, Lender, Borrower, Guarantor, Agent, …, from the
  config lexicon) used repeatedly in the body but never introduced in the parties
  clause, the definitions, or an inline parenthetical `(the "Agent")`.
- **Detect-and-navigate only**; no auto-fix (the project principles, parity with
  feature 01).
- **Reuse**: consumes the Defined-Terms harvester (`DefinitionsMap`, `normalise`,
  usage scan) and the shared **Section Index** (to locate the parties/definitions
  region); emits the shared `Finding` model to the **Proof** tab. Issues no
  document-walk of its own — reads the in-memory Document Model snapshot.

**Explicitly NOT in scope (stays ❌ cut under ADR-0005):**
- **Execution / attestation-block validation** (s.44 two-signatory / witness
  variants; SG common-seal / constitution attestation). The other half of the
  original cut item **remains cut**; this ADR does not revive it.
- Cross-document party resolution (Office JS cannot read other files — ADR-0005).
- Auto-fix / one-click relabelling.
- ML / embedding-based coreference or semantic entity resolution.

## Rationale

- **Why a separate checker, not an extension of Defined-Terms sub-check 1.**
  Sub-check 1 deliberately allowlists party-role words (ADR-0015) to stay quiet;
  folding party-role logic in would either reintroduce those FPs or entangle two
  different heuristics. A standalone checker with its own `featureId`, its own
  evidence model, and its own findings category keeps both clean.
- **Why evidence-gated, conservative on (a).** Dual-role deals — a genuine
  "Buyer" **and** a separate "Purchaser", or "Lender" + "Agent" as different
  entities — are legitimate and common. Flagging on co-presence would be
  corrosive. We require explicit textual evidence of co-reference before flagging:
  e.g. both labels introduced in a parenthetical against the **same named entity**
  (`ACME Pte Ltd (the "Buyer")` … `ACME Pte Ltd (the "Purchaser")`), or one
  label defined in terms of the other (`"Purchaser" means the Buyer`). Absent
  such evidence, **do not flag** — false negatives are preferred to false
  positives, mirroring the maintainer's stance in ADR-0015.
- **Why no ML coreference.** It adds no confidentiality risk by itself but
  invites the suppression-risk philosophy ADR-0008 rejects; explicit textual
  evidence is auditable and conservative.

## Rejected alternatives

- **Keep it fully cut.** Rejected: the maintainer now wants the naming-consistency
  value; reuse makes it cheap; it catches a real high-frequency defect class the
  other checkers deliberately skip.
- **Revive the whole "Execution & Party Consistency Check"** including
  execution-block validation. Rejected: execution-block correctness is a large,
  jurisdiction-deep, fragile capability (s.44 variants, seal/constitution rules);
  the maintainer wants only naming consistency now. **Kept cut.**
- **Extend Defined-Terms sub-check 1 instead of a new checker.** Rejected — see
  Rationale (allowlist entanglement / FP reintroduction).
- **ML / embedding coreference to detect same-party-two-labels.** Rejected:
  suppression-risk philosophy (ADR-0008) and unnecessary complexity; require
  explicit textual evidence instead, accepting missed subtle coreference.
- **Aggressive flagging of any two co-present party-role labels.** Rejected:
  dual-role deals are legitimate; evidence-gated instead.

## Consequences

- **the scope register** moves a new row "**Party-Role Consistency Check (naming
  consistency only)**" to ✅ v1, and the residual **Execution Consistency Check**
  is split out and **stays ❌** with a pointer to this ADR (so the cut survives).
- **ADR-0005** status is annotated "Partially superseded by ADR-0016 (party-role
  naming-consistency clause only)"; the cut record otherwise stands.
- A new feature spec `spec/features/09-party-role-consistency.md` and an
  implementation plan follow.
- **Accepted tradeoff:** the checker is evidence-gated and will miss coreference
  that has no textual evidence — a deliberate false-negatives-over-false-positives
  choice, consistent with ADR-0015.
- **Invariants intact:** all analysis is client-side (no network/ML); party-role
  vocabulary is read from the SG/UK config lexicon (no US convention); not a
  Compare feature, so nothing is suppressed.
