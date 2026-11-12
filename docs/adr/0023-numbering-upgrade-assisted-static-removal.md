# ADR-0023: Numbering upgrade — assisted static-number removal (tracked changes) + human-driven re-apply

- **Status:** Accepted
- **Date:** 2026-06-10
- **Supersedes:** ADR-0006 (static numbering: detect-and-flag, never auto-convert)

## Context

ADR-0006 cut static→dynamic numbering conversion. Its rationale was specific: doing the conversion *as one operation* requires injecting and wiring `abstractNum`/`num` definitions in `numbering.xml` and reassigning every paragraph — **fragile OOXML surgery, with no reliable Office JS undo**, on documents where corruption is unacceptable. It therefore left **all** remediation to the user: the tool flags a manually-typed clause number but never touches it.

Real usage showed the cut was drawn one notch too wide. The tedious, error-prone part of fixing static numbering is not the OOXML — it is the *mechanical* work the user does by hand: locating each hand-typed prefix, selecting it precisely (across run splits and irregular whitespace), deleting it, and then re-applying a correct list at the right level. The product already owns the safe half of that (the typed-API apply engine in feature 06, `setLevelNumbering`/`attachToList`/`startNewList`). What was missing was a safe way to clear the old typed numbers.

The unlock is **Word's native tracked changes**. If the tool performs the prefix deletion *as a tracked change* rather than a silent edit, every one of ADR-0006's fears is defused:

- It is a **typed-API text delete** (`range.delete`/search-delete) — **no OOXML injection, no `abstractNum` wiring**. ADR-0003's OOXML loop is untouched.
- Each deletion is a **native Word revision**: reviewable and individually **rejectable** — the "reliable undo" ADR-0006 said did not exist. The user can accept-all-except-the-one-that-went-wrong.
- The decision about *which* paragraphs become a list, and where hanging/continuation paragraphs belong, is made by **human selection** in a separate second step — not by a heuristic.

This is materially safer than the "beta opt-in conversion" ADR-0006 rejected, because it is never a one-shot convert and never injects OOXML.

## Decision

Add a **Numbering upgrade** to feature 06 (`spec/features/06-numbering.md`): a two-step, human-gated power-up that lets a user clear hand-typed clause numbers and re-impose proper numbering, while keeping the tool's automation bounded to non-destructive and natively-reversible operations.

**Step 1 — Assisted removal (tracked changes).** The user batch-selects flagged paragraphs (via the sidebar queue and/or the live Word selection). The tool deletes, **as tracked changes**, the leading static-number token *and its trailing delimiter/whitespace* on each selected paragraph. The user then reviews the revisions in Word and accepts/rejects them.

**Step 2 — Human-driven re-apply.** The user selects whichever freshly-de-numbered paragraphs they judge should carry numbering (excluding hanging/continuation paragraphs by their own understanding) and runs the **existing** feature-06 apply engine to lay down the scheme. The existing §5.3 cross-reference warn-and-highlight review fires afterward.

The tool **never** decides which paragraphs get re-numbered and **never** numbers a hanging paragraph — that judgment is the user's, exercised through selection.

### Binding design constraints (the design must implement all six)

1. **Tracked changes are mandatory.** Set `document.changeTrackingMode = trackAll` before the deletions and **restore the user's prior mode afterward** (never silently leave tracking toggled on if it was off). The revisions persist regardless of the restored mode. If the host cannot enable change tracking, the removal step is **disabled** (no silent untracked delete).
2. **Batch guard.** Even on a sloppy multi-paragraph selection, only delete from a paragraph where a **leading static-number token is actually detected**. A paragraph with no detected leading number (e.g. a hanging paragraph swept into the selection) is **skipped, never blanked**.
3. **Whitespace auto-rule + single global override.** Per paragraph, auto-decide the delete span: number + delimiter + trailing horizontal whitespace run (spaces / **tab** / NBSP / figure space), and leading whitespace before the number, so body text ends flush (the real list supplies its own indent). The user may set **one global override** (e.g. "collapse to single space" / "strip to flush" / "keep one tab") in the scheme/config store — not per-item fiddling.
4. **Never auto-resequence.** Sequence anomalies (duplicate `1.5, 1.5`; gap `1.5 → 1.7`) are **surfaced with both readings** — typo *or* deliberately reserved number — and never silently "fixed". A gap is ambiguous and the user decides. (Conversion via Step 1+2 *dissolves* a typo because the list recomputes numbers, but that is a user choice, not an auto-fix.)
5. **Leading-occurrence-only selection.** Token selection is anchored to the paragraph start and must span run-split prefixes (e.g. bold `**1.**5`); it must never match a later in-body occurrence (`1.5 … see 1.5 above`).
6. **Re-apply safety unchanged.** Step 2 reuses the existing engine and its existing guards: skip table-cell/footnote, do not modify style-linked lists, respect section-break restarts, Word-Online per-level tolerance, and run the §5.3 cross-reference review after.

### Named edge cases this resolves

- **Hanging / lead-in / continuation paragraphs** — protected structurally: the batch guard (constraint 2) never deletes from a paragraph it did not flag as a number-bearer, and Step 2 numbers only paragraphs the *human* selects.
- **Typed-number typo** (`1.5` where `1.6` was meant) — surfaced as an anomaly with both readings (constraint 4); the user can retype, leave it, or convert (which dissolves it).
- **Irregular whitespace** (`1.5␣␣[clause]`, tab-separated, NBSP) — normalised by the whitespace auto-rule with a global override (constraint 3).

## Rationale

ADR-0006's rationale was a property of the *one-shot OOXML conversion*, not of "the tool helps with static numbers" in general. Tracked changes + typed-API delete + a fully manual re-apply removes exactly the three things ADR-0006 cited (OOXML fragility, no undo, tool-made structural decisions). The corruption-risk principle from the constitution is honoured, not bypassed: every mutation is native, reviewable, and reversible, and the only structural judgment — which paragraphs are a list and which are hanging — stays with the human. This follows "detect-and-flag over risky automation" by keeping the *risky* decision (structure) human and automating only the *safe, reversible* mechanics.

## Rejected alternatives

- **Keep ADR-0006 unchanged (detect-and-flag only).** Rejected: it leaves the entire mechanical burden — precise selection across run splits and irregular whitespace, repeated per paragraph — on the user, for no safety gain now that tracked changes provide native reversibility.
- **Full one-shot static→dynamic conversion via `numbering.xml` `abstractNum`/`num` injection.** **Still rejected** — this ADR does *not* revive it. The OOXML write path remains corruption-prone with no reliable in-place undo (ADR-0003/0006). The whole point is to avoid it: removal is a typed-API text delete, re-apply is the existing typed-API engine.
- **Silent (untracked) deletion of typed numbers.** Rejected: removes the very safety net that justifies reopening 0006. If change tracking is unavailable, the removal step is disabled, not run silently.
- **Tool auto-selects which paragraphs to re-number / auto-numbers the cleared block.** Rejected: this reintroduces the hanging-paragraph hazard (a continuation paragraph wrongly absorbing a number) as a heuristic. The human selects in Step 2; the tool never guesses structure.
- **Auto-fix sequence typos/gaps in place** (rewrite `1.5`→`1.6`). Rejected: a gap is ambiguous between a typo and a deliberately reserved number, and silently rewriting it can break existing cross-references. Surface both readings; let the user decide.
- **One-click "fix everything"** (remove + re-apply in a single action). Rejected: it collapses the human review gate (accept/reject the tracked deletions) and the human structural judgment (which paragraphs are a list) that make this safe. The two steps are deliberately separate.

## Consequences

- Feature 06 gains a **Numbering upgrade** milestone: the assisted-removal queue (Step 1) and the wiring to the existing apply engine (Step 2). the scope register and `spec/README.md` move *assisted, tracked, human-driven* static-number removal into ✅; **one-shot OOXML conversion stays ❌**.
- ADR-0006 is **Superseded by ADR-0023** (kept on disk; status updated).
- No new invariant exposure: typed-API only, zero network, no OOXML injection, no backend. Confidentiality and the OOXML-loop boundary (ADR-0001/0003) are untouched.
- Detection (feature 02 static-numbering flag) is unchanged and remains the source the upgrade reads from; the upgrade never writes that finding.
