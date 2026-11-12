# Backlog — proposed features & enhancements

Ideas captured for future work that are **not yet greenlit for build**. This file
is a holding area, **not** an authorisation to build.

- The canonical in/out boundary remains the project's scope document (not part of this repository).
- Anything here that collides with a ❌ cut requires a **superseding ADR + explicit
  user approval + a same-change update to the scope register** before it can move into the build plan.
- An item that is merely an enhancement of an existing ✅ feature can graduate
  straight into the build plan when scheduled (no ADR).

Newest first.

---

## 2026-06-08 — Defined Terms (Feature 01): two recognition refinements from a review of an open-source comparator

**Origin:** a reference-architecture review of an open-source contract-analysis project. Surfaced as the only two *in-scope, scope-compatible* ideas in an otherwise out-of-scope (substantive-review) comparator. Both are **additive flags** consistent with categorise-don't-suppress (ADR-0008); neither moves a boundary.

- **Circular-definition detection** — a defined term whose definition resolves (directly or transitively) back to itself: term A defined via B, B defined via A. A new structural signal over the existing defined-terms graph.
- **Scope-gated / by-reference definitions** — first-class handling of *"for the purposes of this Clause, X means…"* (locally-scoped) and *"X has the meaning given in Schedule N"* (by-reference) definitions, which today's recogniser does not model as distinct cases.

**Status:** enhancement of ✅ Feature 01 — could graduate straight into the build plan when scheduled (no ADR for the feature boundary). **BUT** both are false-positive-sensitive, so they must first pass an **ADR-0013 / ADR-0015 FP check** (broadened recognition + used-not-defined FP controls) — do not ship either without that gate, or they will manufacture spurious findings on ordinary drafting. Not greenlit for build.

---

## 2026-06-07 — Interactive Lookup: definition-list / expanded-card enhancement (+ split-view request, BLOCKED)

**Origin:** user request — click a defined word in the document body → its definition
pops up in a *list* in the side panel; the relevant definition is jumped-to within the
list and its card *expanded for fuller view* (long / cut-off text is acceptable);
"opens the definition in a split-view that keeps the user where they are." Desktop-only
(not Word Online). Feasibility + scoping investigated during planning.

**Key finding:** ~80% of this is **already shipped** as Feature 04 (Interactive Reference
Lookup — `spec/features/04-interactive-lookup.md`, code in `src/features/interactive-lookup/`).
Clicking a defined term in the body and seeing its definition in the side panel **with the
cursor unmoved** already works today (spec §2.1, §7.3; acceptance T3).

### Part 1 — In-scope enhancement of ✅ Feature 04 (BUILD when scheduled; no ADR needed)
A desktop-first UX polish over the shipped lookup. Each item reuses existing in-memory
`LookupResult` data — **no new Office round-trips per click, no document-size scaling**
(low performance risk; consistent with the build-once dictionary/index, spec §8.5):
- **Accumulating definition list** in the side panel (reuse the `LookupResult` /
  `AmbiguousPanel` patterns in `LookupPanel.tsx`).
- **Expand/collapse card** with the active definition expanded "for fuller view"
  (the panel already expands index 0 only — generalise to the selected card).
- **Relax `lookup.previewMaxChars`** so long definitions aren't aggressively trimmed
  (config change at `dictionary.ts` / `resolve.ts`; keep a sane upper bound + the panel's
  existing internal scroll — "OK if cut off" is satisfied by a generous cap, not unbounded text).
- *(Stretch, capabilities-gated)* a **transient in-document highlight** of the target range
  that does **not** move the viewport. A highlight that scrolls would contradict "keep the
  user where they are"; a non-scrolling highlight of an off-screen range is invisible — so
  this is optional and bounded. If it applies/clears range formatting it adds one Office
  write round-trip per click → debounce + gate it (nudges that sub-item to medium risk).

**Desktop-only gate:** cheap via `src/shared/office/capabilities.ts`, though the enhancement
subset does not technically require desktop (Feature 04 is cross-platform). Honour as the
user's stated preference.

### Part 2 — Literal in-document "split-view" (DO NOT BUILD without ADR + sign-off)
- **Collides with a ❌ cut:** "Split-screen view" (ADR-0005; the scope register ❌ table;
  restated in spec/features/04 §2.2 and §10.1).
- **Feasibility: not achievable via Office JS.** There is no viewport / window / split API
  (spec `ARCHITECTURE.md` Cross-Platform Matrix + Office JS Constraints Reference); there is
  no programmatic control of the desktop View-ribbon "Split"/"New Window". `range.select()`
  *moves* the user and scrolls — it cannot show two document locations at once, so it is
  mutually exclusive with "keep the user where they are."
- **Platform-correct substitute already shipped:** the **task-pane sidebar is the "second
  pane"** — Feature 04 renders the definition beside the live document with the cursor
  untouched, which is exactly the "keeps the user where they are" value.
- **To reopen:** supersede ADR-0005 with a **new ADR** (next free number) + explicit user
  approval + a same-change the scope register update. (NB: ADR-0017 is now PDF comparison and ADR-0018
  is text-vs-text — neither relates to split-view.) Note that even with sign-off the literal split-view remains
  infeasible through the current API. **Recommendation: decline the literal split-view; ship
  the sidebar substitute (already built) and the Part 1 enhancements.**
