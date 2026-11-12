# ADR-0021: AI-readable export ("Copy for AI") — client-side structured data export for downstream LLM use

- **Status:** Accepted
- **Date:** 2026-06-09
- **Extends (does not supersede):** **ADR-0001/ADR-0003** (client-side-only confidentiality invariant), **ADR-0008** (categorise, never suppress — extended here to the machine-export surface), **ADR-0020** (human-facing compare export artefacts). Complements, does not replace, any human-facing export.

## Context

Mutatis computes rich structured data during every analysis run — a definitions map, a clause-tree index, a classified redline change set, normalised dates and amounts, a party/role cast list, and more. Each of these is valuable priming context for an external contract-analyser LLM: handing it a clean glossary of defined terms, a before/after change set, and a deal-facts sheet radically reduces the work the AI must do to reason about the document.

Today none of this data is portable; it is trapped in the task pane. A user who wants to "ask an AI about this contract" must copy raw document text and ask the AI to reconstruct structure that Mutatis has already computed precisely.

Two prior decisions constrain how this capability can be designed:

- **ADR-0020 (proposed)** draws the human-facing export boundary: a "candidate differences report" (Compare/PDF modes) and a "rendered redline snapshot" (Word/Text modes), both read-only artefacts for filing or e-mail. That is a *human-readable* export — a PDF or `.docx` a colleague can open. The present ADR covers a *machine-readable* export — structured text (JSON or Markdown) an LLM ingests as a prompt prefix.

- **ADR-0001/ADR-0003** prohibit the *app* from transmitting document content over any network. A clipboard/file export is local-only: the app writes to the user's clipboard or the user's filesystem; the app is never the transmitting party. The user makes an independent, informed choice about what to do with the text. These are categorically different acts, and the confidentiality invariant is not breached.

However, "documents never leave your machine" is the product's core positioning, and a frictionless "copy everything for an AI" button could be read as undercutting it absent an explicit scope decision and a user-facing awareness signal. This ADR makes both.

## Decision

Add **AI-readable export** ("Copy for AI") as a v1 in-scope cross-cutting capability, governed by five principles:

### 1. Local-only, always

Every export writes to the clipboard (`navigator.clipboard.writeText`) or triggers a local file download (`Blob` + object URL). Zero network calls by the app. The user decides independently what to do with the text.

### 2. One-time awareness notice

On first use of any copy/download action, a non-blocking toast reads:

> "This copies document content to your clipboard for use outside Mutatis. Mutatis makes no network calls. Your firm's policies on AI and data handling apply to what you do next."

Persisted via `localStorage` key `ai-export.noticeShown`. The notice is informational, not a consent gate — the action proceeds whether or not the user reads it. It fires once per browser profile, not once per document.

### 3. Two serialisations, one schema

**JSON** is the canonical machine format (a contract-analyser API ingests it directly). **Markdown** is the paste-into-chat rendering (tables and headings an LLM reads natively). The same underlying export object is serialised by two formatters. A format toggle (JSON | Markdown) is persisted per user in `localStorage`.

### 4. Blind-spot disclosures on every export (ADR-0008 extended)

Every exported section — and the top-level envelope — carries a `blind_spots` array stating what was *not* covered: headers/footers not scanned, image-only PDF pages not text-compared, external references not resolved, unchecked features. An LLM must not be able to assert "no issues" over a region Mutatis never analysed. This is the machine-export analogue of the "what is NOT compared" banner ADR-0020 mandates for the human report.

### 5. Before/after, never markup

Change exports use `{ "original": "...", "revised": "..." }` string pairs per change. Never strikethrough formatting, never `<w:del>`, never caret notation. The consumer is an LLM, not a diff-viewer. `DiffOp.wordDiffs` (del atoms = original, ins atoms = revised) already carries this decomposition; the serialiser traverses it.

### v1 scope

Per the feasibility analysis (2026-06-09):

**Tier-1 per-panel exports (v1):**
| # | Export | Source | Panel |
|---|--------|--------|-------|
| 1 | Defined-terms glossary | `DefinitionsMap` (feature 01) | Defined Terms |
| 2 | Section outline / clause tree | `SectionIndex` (feature 02) | Cross-Reference |
| 3 | Change set with before/after | `DiffOp.wordDiffs` (feature 08) | Compare |
| 4 | Deal-facts sheet (dates, amounts, placeholders) | `DateRecord / AmountPairRecord / PlaceholderHit` (feature 03) | Factual Consistency |
| 5 | Party/role cast list | `PartyAnchor / RoleRecord` (feature 09) | Party-Role |

**Composite (v1):**
| # | Export | Contents |
|---|--------|---------|
| 14 | AI briefing pack (one-click) | #5 + #1 + #2 + #4 + open items from #7 + provenance header |

**Deferred to v2 (Tier 2+):** cross-ref integrity report, open-items checklist, change summary, extracted comments, filled template variables, term-usage consistency map, clause/precedent pair, document hygiene report. See `spec/features/11-ai-export.md` §2.2 for the full deferred list.

## Rationale

- **The confidentiality invariant is technical, not behavioural.** ADR-0001/0003 prohibit network calls by the app; they cannot and should not prohibit the user from copying text to their own clipboard. The awareness notice is the right instrument for the positioning concern — it is informational, respectful of user agency, and fires once.
- **Structured data is worth far more than raw text.** A resolved glossary, a normalised deal-facts table, and a classified before/after change set let the AI skip reconstruction work and immediately reason about substance. This is the principal value proposition.
- **Before/after not markup is non-negotiable.** Strikethrough formatting requires an LLM to parse visual conventions. The `wordDiffs` data already splits every replace-op into original/revised atoms; emitting them as strings is free.
- **Blind-spot disclosures preserve the ADR-0008 no-suppression guarantee.** A machine export that omits disclosures could cause an AI to assert "no issues found" with false confidence over an unanalysed region. The disclosure array is the technical enforcement of this invariant for LLM consumers.

## Rejected alternatives

- **Sending data to a cloud AI endpoint directly** — violates ADR-0001/0003. Non-negotiable and not considered.
- **Silently exporting without any awareness notice** — would undermine the confidentiality-first positioning; a user who doesn't know they are copying document content cannot make an informed choice about their firm's AI policies.
- **Making the notice a blocking consent gate (user must click "I understand")** — needlessly adds friction. The user chose to click the copy button; a modal gate is paternalistic and would be dismissed without reading. An informational toast is sufficient.
- **Exporting only the Findings list (severity/title/location short message)** — too thin. A category name and a clause reference gives an LLM almost nothing to reason about. Full semantic data is what makes the export useful.
- **Exporting the on-screen filtered view** — a filtered export can silently omit a substantive item, identical to the risk ADR-0008 guards against in Compare. Full data set always; same reasoning, same rule.
- **Calling the machine export an "AI summary" or letting the app generate the summary** — the app exports *data*; the AI generates the summary. Calling it a summary would misrepresent what Mutatis does and could erode trust in the underlying data.

## Consequences

- New `spec/features/11-ai-export.md` governs the implementation.
- the scope register gains row 13.
- The task pane gains "Copy for AI" buttons on five panels (Defined Terms, Cross-Reference, Factual Consistency, Party-Role, Compare) and a "Copy Briefing Pack" composite button on the Proof tab header.
- No new scanning code: the feature is a pure serialisation layer over in-memory data already computed by the checkers. The source features must be built first (this feature has no independent document-scanning logic).
- The confidentiality guarantee is unchanged: no bytes leave the machine via the app. The awareness notice is the user-facing signal of the positioning boundary.
- The blind-spot disclosure array extends the ADR-0008 no-suppression guarantee to machine-export consumers.
