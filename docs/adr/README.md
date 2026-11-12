# Architecture Decision Records (ADRs)

This directory records the architecture-affecting decisions for Mutatis in **Michael Nygard's ADR format** — one file per decision, immutable once accepted.

## What an ADR is

An ADR captures a single architectural decision: the **context** (forces in play), the **decision** itself, the **rationale**, the **rejected alternatives**, and the **consequences**. The point is to record *why* a choice was made so a future contributor reads the decision instead of relitigating it from scratch.

## Format

Each ADR follows `0000-template.md`:

```
# ADR-NNNN: <title>
- **Status:** Accepted
- **Date:** YYYY-MM-DD
## Context
## Decision
## Rationale
## Rejected alternatives   ← mandatory; the most important section
## Consequences
```

The **Rejected alternatives** section is mandatory and carries the most weight: it is the anti-relitigation record. An ADR without it is incomplete.

## Rules

- **Rejected ADRs are kept forever.** A decision that was reversed is *superseded*, not deleted. The old ADR stays on disk with its status updated (e.g. `Superseded by ADR-00NN`) so the reasoning survives.
- **Architecture-affecting changes require a new ADR, committed with the change.** You do not change architecture and document it later. The ADR lands in the same commit as the code (or before it).
- **Read the relevant ADR before re-opening a question.** If a request conflicts with an accepted ADR, stop and raise it — supersede the ADR with explicit user approval, or drop the request.
- **Scope changes go through the project's scope register too (not published).** Moving anything off the cut/deferred lists requires superseding the relevant ADR (usually ADR-0005) *and* updating the scope register in the same change.

## Numbering convention

ADRs are numbered sequentially, zero-padded to four digits. **The next ADR is the highest existing number + 1** (currently ADR-0046). Numbers are never reused. `0000` is reserved for the template. **There is intentionally no ADR-0007** — the number is reserved/skipped in this seed set; do not backfill it. **ADR-0025 is likewise a permanent gap** (a draft was started and never landed); do not backfill it either.

## Index

| ADR | Title | Status |
|---|---|---|
| [0001](0001-fully-client-side-no-backend.md) | Fully client-side, no backend | Accepted |
| [0002](0002-singapore-uk-jurisdiction-focus.md) | Singapore & England-and-Wales jurisdiction focus | Accepted |
| [0003](0003-client-side-ooxml-loop.md) | Client-side OOXML read-modify-write loop | Accepted |
| [0004](0004-tech-stack-and-requirement-set.md) | Technology stack and WordApi requirement-set floor | Accepted (front-end/UX partially superseded by ADR-0012) |
| [0005](0005-scope-cuts-and-v2-deferrals.md) | Scope cuts and v2 deferrals | Accepted (partially superseded by ADR-0016) |
| [0006](0006-static-numbering-detect-only.md) | Static numbering: detect-and-flag, never auto-convert | Superseded by ADR-0023 |
| 0007 | *(reserved — intentionally unused)* | — |
| [0008](0008-compare-categorise-never-suppress.md) | Compare: categorise, never suppress | Accepted |
| 0009–0011 | *(internal process decisions — not published)* | — |
| [0012](0012-frontend-ux-and-product-name.md) | Front-end framework, UX design, and product name (Mutatis) | Accepted |
| [0013](0013-defined-terms-broadened-recognition.md) | Defined Terms: broaden recognition beyond quoted-term + default verbs | Accepted |
| [0014](0014-cross-reference-feature-local-office-reads.md) | Cross-Reference reads REF fields, bookmarks, and OOXML feature-locally | Accepted |
| [0015](0015-defined-terms-used-not-defined-fp-controls.md) | Defined Terms: used-not-defined FP controls, all-caps recurrence, shared exclusion lexicon | Accepted |
| [0016](0016-revive-party-role-consistency.md) | Revive a narrowed Party-Role Consistency checker | Accepted |
| [0017](0017-pdf-comparison.md) | PDF comparison sources (PDF-vs-PDF + gated Word-vs-PDF) | Accepted |
| [0018](0018-text-vs-text-paste-compare.md) | Text-vs-text paste comparison with instant standalone redline | Accepted |
| [0019](0019-text-compare-inline-redline-preview.md) | In-pane inline redline preview for the Text-vs-text path | Accepted |
| [0020](0020-compare-export-report-and-rendered-redline.md) | Compare export — candidate-differences report (PDF modes) + rendered-redline snapshot (Word/Text modes) | Accepted (Word/Text snapshot generation path superseded by ADR-0024; artefact framing, never-suppress, disclosures, and the PDF-modes report stand) |
| [0021](0021-ai-readable-export.md) | AI-readable export ("Copy for AI") — client-side structured data export for downstream LLM use | Accepted |
| [0022](0022-standalone-host-portable-implementation.md) | Standalone host-portable implementation — backendless web app / PWA sharing one feature core with the add-in (host-adapter layering) | Accepted |
| [0023](0023-numbering-upgrade-assisted-static-removal.md) | Numbering upgrade — assisted static-number removal (tracked changes) + human-driven re-apply (supersedes ADR-0006) | Accepted |
| [0024](0024-rendercore-own-client-side-redline-rendering-engine.md) | RenderCore — own client-side paginated layout/rendering engine for redlines (worker-resident HarfBuzz core, dual preview + vector-PDF output; partially supersedes ADR-0020) | Accepted |
| [0026](0026-compare-replace-dissimilarity-threshold.md) | Compare — change-complexity threshold (demote a too-dissimilar `replace` to clean delete+insert) | Accepted |
| [0027](0027-markdown-table-compare-text-path.md) | Markdown-table compare in the Text-vs-text path (row/cell-aware structured diff) | Accepted |
| [0028](0028-compare-edit-boundary-anchoring.md) | Compare — edit-boundary anchoring (place a mark where it reads) + restore the clause number on an edited paragraph | Accepted |
| [0029](0029-compare-cross-reference-tokenisation.md) | Compare — cross-reference tokenisation (recognise the lower-case noun, mark only the value) | Accepted (+ dated correction 2026-09-12: the possessive apostrophe is not a quote delimiter) |
| [0030](0030-compare-replace-block-refinement.md) | Compare — replace-block refinement (undo the `cleanupSemantic` avalanche where a finer rendering earns its marks) | Accepted |
| [0031](0031-compare-adjacent-transposition.md) | Compare — adjacent transposition reads as one relocation, not two replacements | Accepted |
| [0032](0032-compare-paragraph-merge-split.md) | Compare — a paragraph merge is one edit, not a deletion beside an insertion | Accepted |
| [0033](0033-compare-defined-term-classification-tests-the-change.md) | Compare — the defined-term class tests the change, not the neighbourhood | Accepted |
| [0034](0034-compare-relocation-preserving-cleanup.md) | Compare — a slot relocation must survive `diff_cleanupSemantic` (apply the cleanup everywhere except across a verbatim relocation) | Accepted |
| [0035](0035-compare-whole-paragraph-mark-revision.md) | Compare — a whole-paragraph change must revise the paragraph mark, or accepting/rejecting it leaves a blank paragraph | Accepted |
| [0036](0036-compare-no-whitespace-punct-filter.md) | Compare — remove the "hide whitespace/punctuation" filter rather than repair it (the class can never be a primary classification) | Accepted |
| [0037](0037-compare-at-scale-one-to-many-many-to-many.md) | Compare at scale — one-to-many + many-to-many batch redlining in the standalone target, with proposed-and-confirmable pairing (extends ADR-0008 to pairing) | Accepted |
| [0038](0038-real-document-readability-instrument.md) | Compare — a readability instrument that can see a real document (real cases on their own scale, a character yardstick beside the token one, the real documents never committed) | Accepted (+ dated note 2026-09-12: fixture set replaced by a clean pair; contamination caveat lifted, false-positive class broadened) |
| [0039](0039-compare-similarity-gap-pairing.md) | Compare — inside a gap, pair paragraphs by similarity, not by position (positional pairs that already render inline are locked; never across an anchor) | Accepted |
| [0040](0040-compare-shared-character-affix.md) | Compare — recover a shared character affix, and run the replace-block refinement to a fixed point (third branch of ADR-0030's acceptance rule) | Accepted |
| [0041](0041-compare-contiguity-escape-from-demotion.md) | Compare — a large contiguous cut is not a rewrite (contiguity escape from ADR-0026's change-complexity demotion) | Accepted |
| [0042](0042-compare-non-text-identity.md) | Compare — text identity anchors, non-text identity annotates (style and resolved numbering as comparable properties) | Accepted |
| [0043](0043-compare-insertion-before-deletion.md) | Compare — the insertion reads before the deletion it replaces (one authored mark order; merge/split seams exempt) | Accepted |
| [0044](0044-web-target-visual-identity.md) | The standalone web target has its own visual identity and an explicit theme preference (partially supersedes ADR-0012 for the web target; AA measured, five values deviate from the desktop reference) | Accepted |
| [0045](0045-compare-inline-objects-survive-a-changed-paragraph.md) | Compare — what is not text survives a changed paragraph (footnote references, links, fields, bookmarks, text boxes re-emitted or revised inside), and an equation's text is compared; amended 2026-10-18: a footnote or picture added, removed or moved alone is reported and marked | Accepted |
| [0046](0046-compare-word-tables-stay-tables.md) | Compare — a Word table stays a table in the redline (body tables as structured alignment units, revised cell by cell through the body's own stages) | Accepted |
