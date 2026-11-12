# ADR-0043: Compare — the insertion reads before the deletion it replaces

- **Status:** Accepted
- **Date:** 2026-09-17
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0028 (edit-boundary anchoring); ADR-0030 (replace-block refinement); ADR-0031 (adjacent transposition — the precedent this generalises); ADR-0032 (paragraph merge/split seams — the one place this can do damage); ADR-0034 (relocation-preserving cleanup); feature 08 (`spec/features/08-compare.md` §5.3, §5.5)

## Context

Inside a changed paragraph the shipped redline always struck first and inserted
second:

```
[-exclusive-]{+non-exclusive+}
```

**That order was never a decision.** `emitReplace` (`output/author.ts`) emits
`op.wordDiffs` in list order with no reordering; the list is
diff-match-patch's, and DMP's internal `diff_cleanupMerge` normalises every
changed region to DELETE-then-INSERT. Nothing between `diff_main` and authoring
swapped the two sides — `refineReplaceBlocks`, `slideEditBoundaries`,
`cleanupPreservingRelocations` and `detectTransposition` are all
length-preserving and self-verify with `sideText(out, 'del') === before.orig`,
which re-partitions marks but never exchanges the sides. §5.5 of the spec said
only "interleaved". So the order the product shipped was an inherited library
default.

The project's reading preference is for the new text to come first:

```
{+non-exclusive+}[-exclusive-]
```

Three things make that the better default, none of them a matter of taste:

1. **The text that survives is what the reader is being asked to approve.** A
   redline is read to answer "what does the clause say now?", and the struck
   text is context for that answer, not the answer.
2. ~~**It is what the market's reference products do.**~~ **WITHDRAWN.** This was asserted without being measured, and the claim was dropped. **Reasons 1 and 3 stand on their own and are why the decision holds.**
3. **This codebase already did it in one place.** `detectTransposition`
   (ADR-0031) emits `ins(B+mid) eq(A) del(mid+B)` — insertion first — so before
   this ADR the product used two orders depending on which pass had last touched
   the stream.

### A claim about another product is a measurement, not a citation

Reason 2 was the only one of the three that was a claim about the world rather than
a judgement, so it was the only one that could be checked, and it was withdrawn
rather than defended. The decision rests on reason 1 and reason 3 alone. Where a
conventional redline puts the insertion relative to the deletion is a matter of
presentation, and nothing here depends on matching any other product.

**The measurement that shaped the plan.** A throwaway implementation was wired
into both pipelines and the full suite run at baseline, with the pass, and with
its seam barrier ablated. Only **6 tests** in **one file** changed outcome
(`compare.readability.test.ts`: five reviewer ideals and one fixture-shape
guard). `compare.output.author.test.ts`, all four export tests, the acceptance
suite, the side-by-side panel tests and both UI panel tests **passed unchanged**.
The emit path had **no order assertion anywhere**. The order was not just
undecided; it was unpinned.

## Decision

**One authored order, insertion before deletion, set once on the `WordDiff`
stream after classification.**

> **Re-confirmed after reason 2 was withdrawn.** The decision was re-opened with
> three options: keep it and fix only the record; keep it and close the
> inline/whole-paragraph split; or revert to struck-first. The maintainer chose to
> keep the current behaviour, so the order is now a **deliberate, informed choice
> resting on reason 1 alone**, not an inherited assumption. What the decision does
> NOT settle: the table-cell gap, the *Go to change* anchor, or the change-list row
> preview.

`diff/mark-order.ts → insertBeforeDelete(WordDiff[]): WordDiff[]` swaps each
maximal `del` run with the maximal `ins` run immediately following it,
re-coalesces, reconstructs both sides and returns its input unchanged on any
mismatch — the fifth length-preserving word-diff pass, built to the shape of the
four that precede it (ADR-0028/0030/0031/0034).

It is applied at **exactly two call sites**: stage 6.5 of `runComparePipeline`
(which serves Word-vs-Word, Text-vs-text and all three batch entry points) and
stage 6.5 of `runPipelinePdf`. The spec's §5 pipeline diagram names it **4a**.

### Where it runs, and why not the two obvious alternatives

- **Not in `word.ts`.** Classification, the SG/UK rules and the cascade collapse
  all read `wordDiffs`. Flipping the order before them makes every classifier
  read a stream in a new order for no gain, and any classifier that is
  order-sensitive becomes a latent defect.
- **Not in `author.ts`.** The authored redline is only one of the consumers. The
  side-by-side panel, the paginated preview, the AI-export change set and the
  exported report all read `op.wordDiffs` directly. Flipping at the author would
  leave them disagreeing with the redline about the same edit.

After stage 6.5, every consumer sees one order.

### Whole runs only

The pass never re-orders within a side. That is what makes it safe by
construction — each side's text is a concatenation in stream order, and moving a
whole run past another whole run of the opposite op leaves both sides' text
untouched. It self-verifies anyway, like its four siblings.

### The seam barrier

The pass **skips any candidate window containing a `PARA_SEAM`.**

A seam is a paragraph MARK, not a word (ADR-0032). `splitAtSeams` cuts the
word-diff stream at each seam and turns the seam-bearing entry's op into the
revision on that mark. A run that crosses a seam in the swap therefore crosses a
paragraph boundary, and the runs either side land in different `<w:p>` elements
than ADR-0032 authored them into.

Two facts about that risk, both measured rather than assumed, because the
barrier's width is a judgement against a cost:

- **The shape is easy to reach.** Every ordinary merge and split puts a
  seam-bearing `del` run beside an `ins` run. Instrumented during planning it
  fired 4× across the suite, with **zero** test coverage: ablating the barrier
  entirely changed no test outcome at all.
- **The worst case is unreachable, and the reachable case is still not
  acceptable.** Rendering a merge as a split — the two mark revisions trading
  places — would need seams on BOTH sides of one window, and
  `detectMergesAndSplits → tryPair` only ever produces 2-original→1-revised or
  1→2, never 2→2. So an op carries at most one seam and that corruption cannot
  occur. What CAN occur, and what the barrier prevents, is text moving between
  the emitted paragraphs while accept-all and reject-all still round-trip: a
  silent rewrite of ADR-0032's output for a presentation preference. Merge/split
  is also where the order matters least to the reader — the two marks sit either
  side of a paragraph mark, not side by side in one sentence.

### It must run exactly once per op

`insertBeforeDelete` is **not idempotent** on a stream holding two changed
regions with no equality between them: `[-A-]{+B+}[-C-]{+D+}` swaps to
`{+B+}[-A-]{+D+}[-C-]`, which leaves `[-A-]` abutting `{+D+}`, and a second
invocation re-windows that pair and merges two independent edits into one
`{+BD+}[-AC-]` block. Both sides still reconstruct, so **the never-suppress
guard cannot see it** — it is a readability regression, not a suppression.

Measured across the suite — and these are two different populations, which the
figure here elided until the fourth review pass said so: **4,535** streams leave
`wordDiff`, **327** of them reach stage 6.5 at all, and **0** carry the chained
shape, because `diff_cleanupMerge` leaves an equality between changed regions
and none of the four upstream passes chain them. So the answer is "call it once",
which the two call sites satisfy, rather than provenance tracking on a swapped
run.

**And that invariant needs a SOURCE guard, because nothing else can hold it.** A
behavioural fixture would need the chained shape, and the shape does not occur —
so the invariant is unfalsifiable from the outside. Measured by the correctness
review pass: with a second `applyMarkOrder(ops);` added beside the first, the full
suite stayed **green** at 211 files / 2454 tests. The sharp form of the finding is
that this ADR and the module header each asserted a double-apply is
invisible to every other guard, and then nothing guarded it.
`tests/unit/compare/pipeline.mark-order.test.ts` now counts the call sites, per the
repo's existing precedent for the same shape (`batch.acceptance.test.ts` → *"the
bundle driver runs a pipeline from exactly one place"*). Re-measured with that
guard in place: the double-apply fails **exactly one test in the whole suite**, and
it is that one.

## Accepted limits

**The order is decided on the `WordDiff` stream, so two authoring sites are not
reached by it.** Found by the first spec-conformance review pass on this diff, not by the
plan — which is worth recording, because an earlier task swept the SPEC for every section
that describes the order and nothing swept the CODE for other authoring sites.
That is an earlier failure mode inverted: the code moved and a description was left
behind, except here the description was the new one and it over-claimed.

Stage 3 skips the word diff for a table-unit record (ADR-0027), so such an op's
`wordDiffs` is `undefined` and stage 4a passes over it. Therefore:

1. **A changed table cell** is authored old-struck-then-new-inserted inside its one
   `<w:p>` — `output/author-table.ts → changedCellParagraph` returns
   `` `<w:p>${del}${ins}</w:p>` ``. Measured end to end through
   `runComparePipeline` on a Markdown table with a changed cell value: the
   `<w:del>` precedes the `<w:ins>`. **This is the del-first order surviving in the
   one place it still reads as "a changed paragraph", and it is a genuine gap
   against this ADR's Decision, not a case the Decision excludes.**
2. **A table paired with a paragraph** (a kind mismatch) emits the whole deletion
   then the whole insertion (`output/author.ts`). This is the **block**-level
   shape, which this ADR never governed: an ordinary `delete` op followed by an
   `insert` op has it too, and so does ADR-0026's demotion. Ordering whole blocks
   is a separate question and is not opened here.

   **Its reader-facing consequence is, though, and it was not stated until
   2026-09-18:** one document now carries BOTH orders. An inline swap reads
   insertion-first; the paragraph beside it that ADR-0026 demoted to a whole
   rewrite still reads struck-first, as do a changed table cell and a merge/split
   seam. "Set once, so no surface can disagree" is true of the surfaces — redline,
   preview, panel, exports — and not of the shapes within one document.

`ParaRecord.table` is produced **only** by `io/text/extract-text-records.ts`, so
both are confined to the Text-vs-text Markdown-table path today.

3. **On the PDF path the seam barrier can only fire on document content.**
   `runPipelinePdf` never runs `detectMergesAndSplits`, so no `PARA_SEAM` is ever
   synthesised there; a U+E000 in a `wordDiff` is the DOCUMENT'S own character,
   which `merge-split.ts` names as reachable from broken-`ToUnicode` PDF
   extraction. The effect is a silent refusal to reorder that one paragraph:
   conservative and harmless, and recorded here rather than left unremarked.

**Why (1) is recorded rather than fixed.** That code belongs to the Markdown-table path
(ADR-0027), which is held behind a live-host smoke test, and this change leaves
those diffs untouched. Changing a gated path's authored output
would invalidate a gate that cannot be re-run here. **The decision is the maintainer's**
and is carried as a named next action; it is a one-line change
(`${ins}${del}`) plus a test, and it wants doing together with clearing that
gate, so the two are smoke-tested together.

## Consequences

**Which worked examples were restated, and which were not.** The test applied was
*does this text claim to show what the tool emits TODAY?* — not which file it lives in.

Brought into line, because they describe the redline a reader is shown:
`spec/features/08-compare.md` §5, §5.3 and §5.5; the corpus ideals the suite asserts;
`refine.ts:15-19` (the avalanche bad/good pair) and `:24-27` (the P3 two-acts pair);
`transpose.ts:14` (the diff-salad example — `:26`, the relocation it is replaced by,
was already ins-first and is the precedent this ADR generalises); `relocate.ts:20`
(the rendering `diff_cleanupSemantic` hands the reader); and
`tests/fixtures/readability-corpus.ts` C32, whose comment began *"Today it reads"*.

Left as they stand, each for a reason that is NOT "historical record":
- `mark-order.ts` — its examples are the pass's own INPUT, which is del-first by
  definition.
- `refine.ts:227-232` (the ADR-0040 affix examples) and `:380` (the fixed-point
  arithmetic) — a stream inside `refineReplaceBlocks`, which runs within `wordDiff`,
  long before stage 6.5. The file says so at the first of them.
- `relocate.ts:11` — the `del`/`ins` there are separated by an `eq`, so no swap
  applies and the rendering is unchanged.
- `readability-corpus.ts` C21 — **measured, not assumed.** It is a paragraph merge,
  the window carries a `PARA_SEAM`, the barrier refuses it, and the tool still emits
  `[-The-]{+, and the+}`. Restating it would make the file lie.

Left because they ARE a historical record: every worked example in ADR-0028 through
ADR-0041. Each prints what the tool emitted
when it was written, and rewriting them would falsify the record of what was decided
when.

- A changed paragraph reads `{+new+}[-old-]` in the authored redline, the
  paginated preview, the side-by-side panel and every export — **wherever the
  paragraph was authored from a `WordDiff` stream.** The one exception is a changed
  table CELL; see § Accepted limits above, which also says why it is recorded and
  not fixed here.
- **The authored bytes change and every `w:id` shifts, but the length does not.**
  On the private `v1`/`v1.1` benchmark pair the authored length is identical before
  and after: a whole-run swap is length-preserving, which is the clearest
  corroboration available that this is presentation and nothing else. Every other
  figure on that pair is unchanged: paragraph and op counts, the same two property
  ops, `summary.total` and the identical `byCategory`. The earlier byte-level pin is
  annotated as superseded rather than overwritten.
- **No `op`, no `classification`, no summary count and no cascade changes** — the
  pass runs after all of them, and that is asserted rather than assumed.

  **This bullet said "Nothing else moves" until the third review pass falsified
  it, and the correction was itself too broad until the fourth.** Two consumers are
  order-sensitive.

  **(i) `searchAnchorFor` (`driver.ts:127`)** takes the first changed line whose
  `anchorWeight` reaches `MIN_WEIGHT` (4) — the first *distinctive* changed line,
  without regard to side. So the *Go to change* anchor is drawn from whichever side's
  first distinctive run now sorts first, which after this ADR is usually the revised
  text. **Not always**, and the case it forbids is constructible: where the insertion
  is short enough to weigh under 4 (`{+X+}[-the Vendor's solicitors-]`) the anchor is
  still the deletion, exactly as before. Measured on this ADR's own example, where
  both sides are distinctive: `"an exclusive"` → `"a non-exclusive"`. **It was never in both,
  in either direction** — `findAndReveal` has always missed for whichever single
  input document is not the anchor's side, so this ADR swapped which flow misses
  rather than creating a miss; and the authored REDLINE carries both sides' runs,
  so navigating the default output (§8.1) is unaffected. Now pinned end to end in
  `compare.driver.anchor.test.ts`, whose other cases hand-build their streams and
  could not see it. **Whether the anchor should be SIDE-AWARE — drawn from whichever
  document is live, or preferring unchanged text present in both — is a design
  question this ADR does not settle; it is carried as an open question.**

  **(ii) `previewText` (`components/ChangeList.tsx:121`)**, which builds §8.3's
  change-list row preview. It walks `wordDiffs` in stream order across BOTH sides
  (`-old +new`, joined) and then truncates **from the head** at 80 characters. So the
  row now leads with the insertion, and for a replace whose insertion alone exceeds
  80 characters the struck text is cut off entirely — where before, the struck text
  led and the insertion was cut. Measured:

  ```
  OLD ORDER → "-shall be liable for loss +shall indemnify and hold harmless each Indemnified Pe…"
  NEW ORDER → "+shall indemnify and hold harmless each Indemnified Person against all Losses ar…"
  ```

  **This is not an ADR-0008 suppression** — the row, its category chip and its
  drill-down are all unchanged, and the old order truncated the other side just as
  hard. It is a change to which side a long replace's preview favours, and it is the
  one consequence of this ADR that is not purely presentational. an earlier task
  recorded *"both UI panels passed unchanged under the flip"*; that measured **that no
  test broke**, and no test covered this path. Now pinned by
  `tests/unit/compare.change-list.preview.test.ts`. **Whether a long replace's preview
  should show both sides — a head-and-tail truncation, or a per-side budget — is a UI
  design question this ADR does not settle; it is carried as an open question beside (i).**

  The lesson is the one this repo already owns: *"nothing
  else moves" is a single confident sentence, so construct the case it forbids* — and
  *a consumer that no test covers is a consumer a green suite cannot clear.* Accept-all still yields the revised document and reject-all the
  original **words** — and, since 2026-10-26, the original **number**
  too. Until then `emitReplace` re-emitted the surviving record's
  `numberingPrefix` as a leading UNMARKED run (ADR-0028), so a paragraph whose
  number also changed rejected back to the original text under the revised
  number: **ADR-0028's bound, which predated this ADR** — this sentence said
  "reject-all the original" unqualified until the correctness review pass
  constructed the case it forbade, and the fixture it was asserted on could not
  show the difference because its prefixes matched. ADR-0028's amendment marks a
  changed number (the new inserted, the old struck, insertion first), and the
  test in `compare.acceptance.test.ts` is now *'rejects back to the original text
  AND number (ADR-0028, as amended)'*. A clause renumbered with
  its words unchanged still rejects to its new number (spec 08 §12 limit 10).
  Nothing is suppressed — the words round-trip exactly, which is what ADR-0008
  governs. On the private pair, `wholeParaChurn`, `alignmentRate`,
  `charOffendingBlocks`, `charRedundancy`, `redundancy` and `marks` are all
  identical before and after.
- The corpus's five reviewer ideals (C03, C16, C23, C24, C29) are re-authored in
  the new order and the fixture-shape guard now looks backwards from its
  contentless `del`. All nine `idealMatch` pins hold — 8 true, C21 deliberately
  false. `TOTAL_REDUNDANCY_CEILING` stays 0.11 and every per-case `marks` is
  unchanged: `replaceBlocks` in the yardstick was already written order-agnostic.
- Three previously duplicated private copies of `sideText` (`transpose.ts`,
  `refine.ts`, `anchor.ts`) are now one module (the project principles), with a direct
  unit test — the three guards compare that function against itself, so a
  symmetrically wrong `sideText` is invisible to every one of them.
- **The order is now pinned by tests that can fail.** It was not before. The seam
  barrier, both wiring lines and the emit path each have a named test that
  breaks when they are removed.

## Alternatives considered and rejected

- **A config key (`compare.markOrder`).** The maintainer expressed a preference, not a
  request for a setting. A second authored order doubles the surface of the one
  artefact that must round-trip through Word — accept-all and reject-all have to
  be right under both — for a choice no user has asked to make. If a knob is ever
  wanted it belongs in the style profile, the single home for presentation
  config, not as a ninth `compare.*` key beside the behavioural ones in §7.5.
- **Flipping inside `word.ts`, before classification.** Rejected: the classifiers,
  the SG/UK rules and the cascade collapse would all read the new order for no
  gain, widening the blast radius from one stage to five.
- **Flipping inside `author.ts`, at emit.** Rejected: only the authored redline
  would change. The side-by-side panel, the preview, the AI export and the
  exported report read `wordDiffs` directly and would disagree with the redline
  about the same edit.
- **Letting `detectTransposition` stay the exception.** Rejected: that is the
  status quo — two orders in one product, decided by which pass last touched the
  stream. This ADR makes the transposition rendering the rule rather than the
  outlier.
- **Swapping across a seam as well, and re-deriving the paragraph split.** Rejected
  as the wrong trade: it would rewrite ADR-0032's authored merge/split output for
  a presentation preference, in the one place where the order matters least to
  the reader and where getting it wrong moves a paragraph boundary.
