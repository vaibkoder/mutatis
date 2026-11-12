# ADR-0028: Compare — edit-boundary anchoring (place a mark where it reads, and keep the clause number)

- **Status:** Accepted; Part A amended 2026-10-26 (§ Amendment — a changed number in an edited clause is marked)
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0026 (change-complexity threshold); feature 08 (`spec/features/08-compare.md` §5.0, §5.3)

## Context

A readability assessment ran 22 single-act indemnity amendments through the real
Word-vs-Word / Text-vs-text pipeline and read the result as a lawyer would. Alignment,
atomic-value handling, move detection and the ADR-0026 demotion were at or beyond
conventional-redline quality. Two defects were not.

**1. An edited clause silently lost its number.** `emitReplace` discards the template
paragraph's own runs and rebuilds it from `wordDiffs`, and those derive from `normText` —
which had the manual numbering prefix stripped by `stripNumberingPrefix` (§5.0) so a
renumbering cascade would not read as a change. The equal / insert / delete / move paths
emit `rawXml` wholesale and keep their prefixes. Net effect on a manually-numbered
contract: the redline keeps clause numbers everywhere **except on the clauses that
changed** — precisely the ones the reader navigates to, quotes in the reply email and
argues about. 18 of the 22 specimens lost their number. It is not a text-mode artefact;
`preprocess/pararecord.ts` strips the same prefix, so the Word-vs-Word path is equally
affected wherever numbering is hard-typed rather than a list.

**2. Marks were anchored without ever looking at real text.** `diff_main` returns *a*
minimal edit script, not the most readable one. Where several placements of the same edit
are equally minimal, DMP picks whichever its scan reaches first. DMP ships the right
remedy — `diff_cleanupSemanticLossless`, which slides an edit boundary to the nearest
logical break without changing *what* changed — but `diff/word.ts` diffs in **sentinel
space**: every token is mapped to an arbitrary code point before `diff_main`, so that
pass is scoring characters carrying no punctuation, whitespace or word information. No
boundary decision was ever taken on real text.

The recurring legal-drafting casualty is an appended proviso or list limb, where the new
material ends on the same word the original sentence ended on:

```
… ending on or before {+Completion, provided that … paid before +}Completion.
```

Minimal, and unreadable: the untouched clause is cut off mid-flow and then finished by a
word that looks like part of the new text, so the reader cannot see that the original
obligation is unamended. The same edit anchored one phrase later reads correctly:

```
… ending on or before Completion{+, provided that … paid before Completion+}.
```

A compounding cause is that `tokenize` glues sentence- and clause-level punctuation onto
the word it follows, so `Completion` and `Completion.` are different tokens and the final
occurrence of a word can never align with the first — the boundary is *unable* to move
even in principle.

## Decision

**A. Re-emit the manual numbering prefix on the replace path.** `emitReplace` prepends the
template record's `numberingPrefix` as one clean, unmarked run
(`numberingPrefixRun`). Deliberately **not** in `emitEqual`: that path emits `rawXml`,
which already carries the prefix in its run text, and re-injecting there is the
doubling bug (`"(1) (1)  Party"`) fixed in `2fd4eb5`. The prefix is taken from the
template — the revised record where there is one — matching `emitEqual`'s
"prefer the revised numbering" rule: a number that moved because clauses above it shifted
is a *consequence* of that edit, not an amendment of its own (§6.1), so it is emitted
unmarked. Auto-numbered paragraphs have no literal prefix (`numberingPrefix === ''`) and
are untouched; their labels come from the numbering scheme.

**B. Anchor each single-sided edit on real tokens** — a new pass,
`diff/anchor.ts → slideEditBoundaries`, run inside `wordDiff` on the decoded per-token
diff (real text, atomic kinds attached) immediately before coalescing. For every
`eq → (all-ins | all-del) → eq` triple it shifts the edit right through identical
neighbouring tokens and keeps the placement scoring best on three purely structural
signals — no vocabulary lists, so nothing is language- or jurisdiction-specific:

- **−6** the span *opens* on the same word that reappears immediately after it closes
  (the tell-tale of a mis-anchored edit: the reader cannot tell which occurrence belongs
  to the untouched sentence);
- **+3** the span opens on punctuation, or directly after it (a clean join);
- **+2** the span closes on punctuation (self-contained);
- **+1** each for whole-word edges either side (tie-breaks only).

Ties prefer the later placement: appended material reads better anchored after the text it
follows than before it.

To let a boundary reach a word carrying a punctuation tail, the pass splits that tail into
its own token *locally, at the boundary only* (`Completion.` → `Completion` + `.`). It
never does so to a token `atomicKindOf` recognises, so amounts, dates, periods, clause
references and quoted defined terms stay indivisible (§5.3). This is a surgical fix at the
one place it is needed, **not** a change to the global tokenizer, whose atomic-token
behaviour the classifier depends on.

`flatten` stamps every token with the op of the segment it ended in: a token that slid
across a boundary still carries the op it was born with, and that op is what the author
and the guard read, so the segment has to win — otherwise the boundary moves while the
mark stays put.

## This does not violate the never-suppress invariant (ADR-0008 / the project principles)

Anchoring is **length-preserving by construction**. A shift is legal only when the text
leaving the front of the span is the same text arriving at its back, so the concatenation
of each side is unchanged; the pass moves *where a mark sits*, never *what is marked*.
Nothing is dropped, hidden, filtered or merged, and no change leaves the redline: the same
insertions and deletions are present, individually accept/rejectable in Word, and still
classified. `slideEditBoundaries` additionally **verifies** the invariant on every call —
it reconstructs both sides before and after and returns its input unchanged if they ever
differ — and that guard is asserted by test.

Part A likewise adds only an unmarked run carrying text the paragraph already had; no
revision markup is added or removed.

## Rejected alternatives

- **Do nothing.** Rejected: an edited clause with no number, and a proviso that appears to
  amend the sentence it merely follows, are the two readability failures a lawyer notices
  first. Both are in the core value of the feature.
- **Re-inject the numbering prefix in `preprocess`, so `normText` keeps it.** Rejected:
  the prefix is stripped from `normText` on purpose — it is what stops a renumbering
  cascade reading as a change on every downstream clause (§5.0, §6.1). Putting it back at
  the comparison layer trades a rendering defect for a far worse alignment defect.
- **Re-inject in `emitEqual` as well, for symmetry.** Rejected: that path emits `rawXml`,
  which already carries the prefix; this is the doubling bug fixed in `2fd4eb5` and now
  pinned by regression test.
- **Call DMP's `diff_cleanupSemanticLossless` on the decoded real-text diff.** Rejected as
  insufficient, having been measured: its scoring rewards whitespace-aligned boundaries,
  which is exactly the placement that produces the unreadable proviso. On the worked
  example it scores the bad placement 4 and the good one 2 — it would actively keep the
  defect. The problem is not where the *surrounding characters* break but whether the
  *span itself* is a well-formed unit, which its scoring does not model.
- **Split terminal punctuation into its own token globally in `tokenize`.** Rejected for
  this slice: it changes the token stream every classifier rule and atomic test reads, for
  a benefit fully obtained by splitting locally at a boundary that is about to move.
  Recorded as the broader follow-up, not forgotten.
- **Extend the pass to two-sided replaces** (the `[-basis; it-]{+basis. It+}` widening,
  where a word identical on both sides sits inside both marks). Rejected **for this
  ADR/slice**: a replace has no free boundary to slide — the fix there is to stop
  `diff_cleanupSemantic` fusing tight edits into coarse blocks, a separate change to a
  DMP call that many existing tests pin. Deferred as a follow-up; recorded here so the
  boundary is explicit.
- **Drop `diff_cleanupSemantic` altogether.** Rejected: measured on the same corpus it
  tightens one specimen and fragments another into two adjacent replacements. The
  replacement has to be a *bounded* cleanup, which is the deferred follow-up above, not a
  deletion.

## Amendment (2026-10-26) — a changed number in an edited clause is marked

Part A emitted the revised number **unmarked** in every edited clause. Where the
number itself changed with the edit — `(a)` became `(b)`, or the number was added or
removed — rejecting every change kept the new number on the old clause: reject-all
was not the original document (ADR-0008's reconstruction). The recorded bound in
`compare.acceptance.test.ts` (*rejects back to the original text AND number*) said so
and named the fix; it now expects `1.1 …` on reject-all.

**Decision.** In an edited clause (`emitReplace`, and the first paragraph of an
ADR-0032 merge/split rebuilt from text), the same number on both sides is still one
clean, unmarked run — the `2fd4eb5` doubling case stays pinned. A number that
differs is the revised number **inserted** and the original **struck**, insertion
first (ADR-0043), as part of the same op: no new op kind, the classification
unchanged. Where the paragraph is rebuilt from its source runs (spec 08 §5.5,
a later change) the number is written from its own runs, in its own formatting
(`numberFromSource`); the text rebuild writes it as plain runs (`numberRuns`).

**Not changed.** A clause whose *words* did not change but whose manual number moved
because a clause above it was inserted or deleted is an `equal` op, emitted from the
revised `rawXml`; reject-all reads the revised number there. That is the renumbering
cascade (§6.1), outside this amendment; recorded as spec 08 §12 limit 10, with the
readability corpus's count (2 clauses each in C18, C19, C20, C28).

**Never-suppress.** The amendment adds revision markup round text both documents
already carry; nothing is hidden, and both views now reconstruct.
