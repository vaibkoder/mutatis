# ADR-0019: In-pane inline redline preview for the Text-vs-text Compare path

- **Status:** Accepted  <!-- user-approved 2026-06-09, in the same change as the scope register update below -->
- **Date:** 2026-06-09
- **Extends (does not supersede):** **ADR-0018** (Text-vs-text paste → instant standalone `.docx` redline) and **ADR-0008** (categorise, never suppress). ADR-0018's decision stands in full; this ADR *adds* a read-only surface alongside it.

## Context

ADR-0018 shipped Text-vs-text as **`.docx`-out only**: paste two blobs → an editable tracked-changes Word document, built eagerly, **opened only on the user's click**. In use, the panel therefore shows the user *nothing about the changes themselves* until they open a separate Word window — only a "Open redline (.docx)" button and the category summary. For the snippet use case ADR-0018 targets (a clause two ways, a counterparty line vs the house position), the user wants to **see what changed immediately, in the pane**, without spawning Word.

The result is already in memory. The driver holds `ops: DiffOp[]` (the same array feeding the summary and the change list); each `replace` op carries `wordDiffs` (`eq`/`ins`/`del` tokens) and `origRange`/`revRange` into the per-side `ParaRecord[]`. The PDF path **already renders this inline** — `components/SideBySideDiff.tsx` + the `.dp-compare__tok--ins|del|eq` styles — precisely because a PDF has no `.docx` redline target. So an inline preview for the text path is a **pure rendering addition over data that already exists**: no new diff, no new classify pass, no second pipeline, no extra memory beyond DOM nodes.

ADR-0018's *Rejected alternatives* include *"a bespoke in-pane text differ with **no `.docx` output**"*, rejected because *"it would rebuild the classifier and forgo the editable-Word-redline value."* This decision is **materially different** and the rejection's rationale does not apply: the preview is **additive** (the `.docx` remains the editable artefact and the primary output), and it **reuses the existing `ops`/classifier** verbatim (nothing is rebuilt). What it changes is the boundary ADR-0018 drew — text path = `.docx`-out only — so it is recorded here as its own decision, per the scope register "Changing scope".

The one genuine engineering concern is **scaling with text length**: an inline redline is a linear token stream, and a long passage can be tens of thousands of DOM nodes, which would jank the task pane.

## Decision

Add a **read-only inline redline preview** to the Text-vs-text results, rendered automatically when a comparison completes, **in addition to** the unchanged eager-`.docx` / "Open redline" flow.

1. **Unified inline view (not side-by-side).** A single flowing passage in document order — `eq` text plain, `ins` underlined, `del` struck through — reusing the existing `.dp-compare__tok--*` token styles. The task pane is ~340px; a Word-style unified redline reads far better than two columns here. `SideBySideDiff`'s two-column variant stays PDF-only.
2. **Hunk + collapsible unchanged context (scales with *number of changes*, not document length).** Render each changed op plus a small, fixed amount of surrounding unchanged context; collapse long unchanged runs into an **explicit, one-click** `⋯ N unchanged paragraphs ⋯` expander. Render cost tracks the number of edits — the quantity a reviewer actually reads — so a long document with sparse edits stays cheap.
3. **Hard guardrail against silent truncation (the project's "no silent caps" rule).** If the op or token count exceeds a configured threshold, the view defaults to changes-only **and shows a visible notice** naming what is collapsed. Nothing is ever dropped without the user being told; "expand all" remains available.
4. **Virtualisation deferred behind the threshold.** Windowed rendering (mount only ops near the viewport) is *not* built in v1; the hunk+collapse model already bounds DOM for the realistic case. It is recorded as the next lever if a measured case (e.g. a global find/replace producing thousands of edits) needs continuous full-passage scroll. No heavy virtualization dependency (consistent with spec §9.10 deferring list virtualization).

**v1 boundary:** read-only preview only — no accept/reject, no editing in the pane (the `.docx` is where the user edits). Text path only (the PDF path already has its in-pane view; the Word-vs-Word path keeps writing back to the live document and is out of this ADR).

## Rationale

- **Immediate feedback, zero friction.** The user sees the redline the instant Compare finishes — the missing half of the snippet workflow — without a Word window opening.
- **Maximum reuse, minimum new surface.** No new diff/classify/pipeline; the preview is a presentational component over the existing `ops` plus the per-side `ParaRecord[]`, using styles already on disk.
- **ADR-0008 preserved.** Every change renders; only **unchanged** context collapses, behind an explicit expander. The summary total stays the full count and the list filters remain bound to the navigation list only — the preview shows everything.
- **Confidentiality / jurisdiction untouched.** Pure client-side rendering of in-memory data; no network, no locale assumptions.
- **The `.docx` value is intact.** ADR-0018's editable native redline is unchanged and remains the primary, editable output; the preview is a complement, not a replacement.

## Rejected alternatives

- **Replace the `.docx` with the in-pane view (the ADR-0018 rejected alternative).** Still rejected. The editable Word redline is the thing lawyers act on; the preview is read-only and additive. ADR-0018 stands.
- **Side-by-side two columns for the text path.** Rejected for v1 — at ~340px the columns are too narrow to read; unified inline is the better fit. The column component remains for PDF.
- **Render the entire passage eagerly (no hunking/collapse).** Rejected — does not scale with text length; a long passage janks the pane. Hunk+collapse bounds cost to the number of changes.
- **Build windowed virtualisation now.** Rejected for v1 — premature for the realistic sparse-edit case and adds complexity/possible dependency; deferred behind the measured threshold (see Decision §4).
- **Silently cap the preview at N changes.** Rejected — violates the "no silent caps" rule and risks reading as "no further changes." Over-threshold ⇒ changes-only **with a visible notice**.
- **Accept/reject or edit in the preview.** Rejected for v1 — that is exactly what the editable `.docx` provides; duplicating it in-pane is scope creep and Office-JON-free editing of synthesised content has no target document.
- **ML semantic-equivalence / any suppression of a change.** Remains ❌ (ADR-0008).

## Consequences

- A **small, presentational** feature lands on the existing Text-vs-text path. New code is confined to: a new `compare/components/InlineRedline.tsx` (unified renderer with hunk/collapse), a pure `compare/components/hunks.ts` helper (group `ops` into changed hunks + collapsible context; exported for unit tests), the threshold constant, and wiring into `TextComparePanel.tsx`'s results block. `output/**`, the diff, and the classifier are untouched.
- **ADR-0018's `.docx` flow is unchanged** — same eager-generate, same click-to-open button, same bytes in memory. The preview reads the `ops` the panel already has.
- Confidentiality, jurisdiction, scope, and categorise-never-suppress are all preserved; this ADR + the user's approval + the scope register update **is** the scope gate that moves the Text-vs-text boundary from `.docx`-out-only to `.docx`-out **plus** read-only in-pane preview.
