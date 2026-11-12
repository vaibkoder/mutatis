# ADR-0014: Cross-Reference reads REF fields, bookmarks, and OOXML feature-locally

- **Status:** Accepted
- **Date:** 2026-06-06

## Context

Feature 02 (Cross-Reference Checker & Static Numbering Detector) needs three pieces
of document information that the shared Document Model (`shared/doc-model`) does
**not** expose, by design:

1. **REF/PAGEREF fields** — `paragraph.fields` → `field.code` — to decide whether a
   prose reference is backed by a live Word field (no static-reference badge) or is
   manually-typed (badge), and to detect a *broken* REF field whose target bookmark
   no longer exists (spec §5.1.4, §9 "REF field that is itself broken").
2. **Bookmarks** — `context.document.bookmarks` — to validate a REF field's target
   and to populate `SectionIndexEntry.bookmarkNames` (spec §4.2).
3. **Leading-token OOXML** — `paragraph.getOoxml()` — for Feature B's layer (c)
   residual-ambiguity probe (a number that is neither a `listItem` nor obviously a
   `<w:t>` run; spec §5.2.3).

The shared loader (`loadDocSnapshot`) is the project's most performance-critical
module and is deliberately a lean single pass capturing `text`, `style`,
`listItem`/`listString`, and table-cell tagging only (ARCHITECTURE.md §Document
Model). `getFileAsync`/OOXML is the sanctioned escape hatch but is owned by the
OOXML Service for **write** loops (Repair/Compare), not per-paragraph reads.

The question: where do these three reads live?

## Decision

Cross-Reference performs these reads **feature-locally**, within its own bounded,
batched `Word.run` passes — it does **not** enrich the shared loader:

- **REF-field + bookmark read** — one batched feature-local `Word.run` over the
  paragraphs that the prose scanner flagged as carrying reference matches (not the
  whole document): load `paragraph.fields` + `field.code`/`field.result`, load
  `document.bookmarks` once, correlate by range overlap. Lives in
  `src/features/cross-reference/fields.ts`.
- **Layer (c) OOXML probe** — `paragraph.getOoxml()` called **only** for Feature B's
  residual-ambiguous set, **capped** at `staticNumbering.maxOoxmlProbes` (default 200,
  config). Lives in `src/features/cross-reference/static-numbering` layer-c module.
- **`SectionIndexEntry.bookmarkNames`** is added to the shared type (additive,
  optional, default empty `[]`). The shared builder leaves it empty; Cross-Reference's
  feature-local field/bookmark pass is the only v1 populator.

All three are client-side typed-API / `getOoxml()` reads. No bytes leave the machine;
the confidentiality invariant (ADR-0001/0003) is untouched.

## Rationale

- **Lean single pass stays lean.** Loading `fields` and `getOoxml()` for every
  paragraph in every analysis run — when only Cross-Reference needs them in v1 —
  would degrade the load that *all* features share, against ARCHITECTURE.md
  §Performance ("single-pass loading", "batch `ctx.sync()`").
- **Dependencies flow features → shared.** Field/bookmark/OOXML reads scoped to one
  feature belong in that feature; promoting them to shared infra before a second
  consumer exists is speculative coupling (the project principles reuse is about *existing*
  cross-cutting logic, not pre-emptive hoisting).
- **Within the architecture, not against it.** Typed-API field reads are interaction
  model (A); `getOoxml()` is the documented per-paragraph read primitive. Neither
  contradicts the OOXML Service's remit (the write loop). ADR-0006 already says
  static-numbering detection may inspect content types.
- **Bounded cost.** The field read is scoped to reference-bearing paragraphs; the
  OOXML probe is hard-capped — both honour the large-document performance budget.

## Rejected alternatives

- **Enrich the shared `loadDocSnapshot` with fields/bookmarks/OOXML.** Rejected:
  over-loads the single pass for one v1 consumer; heavier sync on every run for every
  feature; couples shared infra to xref-only needs. If Interactive Lookup (04) or
  Numbering (06) later need bookmarks, *that* is the trigger to hoist — with its own
  trace/ADR — not now.
- **Route the reads through the OOXML Service (`shared/ooxml`).** Rejected: that
  service is the `getFileAsync` → unzip → edit → `insertFileFromBase64` **write**
  loop for Repair/Compare; using it for per-paragraph typed reads is the wrong tool
  and pulls in the two-document-in-memory serialisation constraints unnecessarily.
- **Uncapped `getOoxml()` probing.** Rejected: `getOoxml()` is expensive; an uncapped
  per-paragraph probe on a 400-page document blows the performance budget. The cap
  (config, default 200) with residual paragraphs reported `indeterminate` is the spec
  §5.2.3 / §10.7 behaviour.

## Consequences

- `SectionIndexEntry`/`SectionNode` gains an additive optional `bookmarkNames: string[]`
  (default `[]`); existing consumers (none populate it yet) are unaffected.
- The field/bookmark read is the plan's **escalate** task (Office-JS sequencing is
  fragile). **Documented fallback:** if `paragraph.fields` is unavailable/unreliable
  on a host, degrade to treating *all* matched references as un-backed (static badge
  applies, broken-field detection disabled) behind a non-blocking capability banner —
  never throw, never block the rest of the check (ARCHITECTURE.md graceful degradation).
- A future feature needing bookmarks/fields is the trigger to revisit hoisting into
  the shared loader; until then this stays feature-local.
- No scope boundary moves; the scope register unchanged. Confidentiality/jurisdiction/scope/
  compare invariants intact.
