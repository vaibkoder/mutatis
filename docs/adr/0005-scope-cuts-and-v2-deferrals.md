# ADR-0005: Scope cuts and v2 deferrals

- **Status:** Accepted (partially superseded by ADR-0016 — party-role naming-consistency clause only)
- **Date:** 2026-06-06

## Context

Mutatis must resist scope creep across many independent working sessions. Numerous capabilities were considered and deliberately cut or deferred — some are platform impossibilities (no Office JS API), some are US-centric (wrong audience), some are confidentiality violations, and some are too fragile to ship safely for legal documents. Without a single durable record, a future contributor re-proposes a cut feature as if it were new.

## Decision

Record the consolidated **cut list** (do not build) and **v2 deferrals** (do not build in v1, do not block v1) exactly as in the scope register. This ADR is the **master anti-scope-creep record**.

## Rationale

A cut is never silent and a deferral always names its v1 replacement. Pinning each item with its rejection reason here means a future contributor reads the decision rather than relitigating it.

## Rejected alternatives

Each item below was considered and rejected for the stated reason.

- **Citations Checker / Table of Authorities / Bluebook / OSCOLA** — US/litigation-centric; not the SG/UK transactional audience (ADR-0002).
- **Execution & Party Consistency Check** — removed by the maintainer. (The **party-role naming-consistency** half is later revived as a narrowed standalone v1 checker — see ADR-0016. The **execution / attestation-block validation** half stays cut.)
- **Text-box / drawing-layer detection & removal** — `body.shapes` is desktop-only (`WordApiDesktop`); platform limitation.
- **Split-screen / docked target view** — no Office JS API to split the viewport; the task pane is the only surface.
- **Real-time per-keystroke typing feedback** — the event model makes it noise; replaced by debounced, selection-triggered feedback.
- **Cross-document / external-file reference & exhibit resolution** — Office JS cannot read other files; detected and labelled "external", never resolved.
- **Live REF-field insertion in Assign-source** — fragile async chain; v1 does static text replacement (v2 seam left).
- **True "out-of-date" staleness detection** — needs version history; ship a "static reference" badge instead.
- **Per-author tracked-change management** — Office JS is bulk accept/reject only.
- **OneDrive / SharePoint scheme sync** — replaced by JSON export/import.
- **Auto cross-reference updating after renumbering** — error-prone; warn-and-highlight instead.
- **Table-cell numbering / style-linked list modification** — edge-heavy, template-corruption risk.
- **Repeating-section cloning in Templates** — Repeating Section Content Control not exposed; fixed slots (default 4, max 6) instead.
- **KMS integrations (iManage / NetDocuments / OpenText) in core** — enterprise plugin adapter only, not the open-source core.
- **v2 deferrals:** holiday-calendar date checks (v1 does weekend detection only); Word-vs-PDF compare via PDF.js; fuzzy moved-and-edited paragraph detection in Compare (v1 does verbatim hash-match moves).

(ML semantic-equivalence in Compare is cut separately — see ADR-0008. Any backend/server is barred by ADR-0001.)

## Consequences

Bringing any cut item back into scope requires **superseding this ADR** with new context and rationale, **explicit user approval**, and updating the scope register in the same change. Do not build first and document later.
