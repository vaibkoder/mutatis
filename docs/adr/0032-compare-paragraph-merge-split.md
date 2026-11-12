# ADR-0032: Compare — a paragraph merge is one edit, not a deletion beside an insertion

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0026 (change-complexity threshold); ADR-0027 (a table is one alignment unit); ADR-0030 (replace-block refinement); feature 08 (`spec/features/08-compare.md` §5.1, §5.3, §5.5)

## Context

Level-1 (`structuralDiff`) matches paragraphs **one-to-one**. When a drafter
joins two paragraphs into one, that model has no way to say so: paragraph 1
pairs with the joined paragraph and becomes a large `replace`, and paragraph 2
is left without a partner and struck whole. The reader sees the same twelve
words twice — once struck at the end of the redline, once inserted into the
paragraph above — and has to diff the redline themselves to discover that
nothing happened to them. That is the P2 failure in its purest form, at
paragraph scale:

```
bad   … notify the Vendor of any Third Party Claim{+, and the notice shall specify
      the nature of the claim in reasonable detail.+}
      [-The notice shall specify the nature of the claim in reasonable detail.-]
```

A split (one paragraph becoming two) is the exact mirror.

### The metric had a blind spot, and it hid this case

`redundancy` — the objective function the whole readability programme is scored
against — was computed from `replaceBlocks` **within a single paragraph**. C21's
rendering above therefore scored **0**: the insertion and the deletion sat in
different paragraphs, so the metric could not see that they were the same words.
One of the worst-reading cases in the corpus was invisible to the yardstick.

`tests/fixtures/readability-metrics.ts` now also measures redundancy **across**
paragraphs: text inserted into one paragraph that reappears verbatim in a
wholly-struck neighbour. Insertions are counted only from paragraphs that are
not themselves wholly inserted, so ADR-0026's clean paragraph-for-paragraph
swaps stay exempt — a whole paragraph replaced by a whole paragraph is one
honest editing act, not a redundancy.

**This raises previously-reported scores, and the earlier numbers must be read
in that light.** Under the corrected metric the corpus total *before* this
change is **1.05**, not the 0.11 recorded in ADR-0030 — C21 alone accounts for
0.94 of it. ADR-0030's measurements are correct for the passes they describe;
they were simply taken with an instrument that could not see this defect.

## Decision

Add **stage 2.25** to the pipeline — `diff/merge-split.ts →
detectMergesAndSplits`, between Level-1 and the ADR-0026 dissimilarity gate —
which re-pairs a `replace` sitting beside an orphaned `delete` (a merge) or
`insert` (a split) into **one `replace` spanning the whole range** on the
many-side.

No new op type is needed. `DiffOp.origRange` / `revRange` are already ranges and
`recordsInRange` already walks them: a merge is a `replace` whose `origRange`
covers two paragraphs, a split one whose `revRange` does. Running before the
ADR-0026 gate means that gate scores the merge against the **joined** original
text, which is what it actually replaces.

**The paragraph mark is a first-class edit, not something glued away.**
`rangeDiffText` joins the range with a Private Use Area sentinel (`PARA_SEAM`,
U+E000) standing for each paragraph mark. The tokenizer emits it as its own
single-character token — it is neither whitespace nor a word character — so the
Level-2 diff strikes or inserts it as a unit, and `output/author.ts` turns that
into the revision Word actually records for a lost paragraph break:
`<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>` on the paragraph mark
(`markParagraphMark`). Accepting it joins the paragraphs; **rejecting it restores
the break.** A merge authored without this would not be a redline of the merge at
all — the break would simply vanish, unrejectable.

A single-record range is the record's `normText` **unchanged**, so every
existing one-to-one `replace` produces a byte-identical word diff. That is the
condition this pass had to meet before it could land, and it is pinned directly.

### Guards — structural only (ADR-0008: no semantic-equivalence scoring)

- **`MERGE_MIN_TOKENS` (4).** Never re-pair on a fragment. A limb stub, a
  heading or `"; and"` matches by chance, and the reading gain is nil.
- **`MERGE_CONTAINMENT_FLOOR` (0.85), positional.** The absorbed paragraph must
  reappear token-for-token **and in order** at the matching end of the survivor.
  A neighbour that merely shares vocabulary has not been absorbed. The floor is
  high on purpose: gluing two genuinely unrelated paragraphs together is a far
  worse defect than leaving a merge rendered the old way.
- **The remainder must still be its pair.** What is left of the host after the
  absorbed part is removed must still resemble the paragraph the `replace`
  matched, scored with the **ADR-0026 threshold** so the two gates cannot
  disagree. A host that is *entirely* the absorbed paragraph is a deletion of
  the other one, not a merge.
- **Strict adjacency**, and the orphan must be a single paragraph. (Written
  first for four orderings; only two exist — see the correction below.)
- **Never a table unit** (ADR-0027): a table is one alignment unit.

Never suppresses: the pass only changes **which op owns which record**. It
reconstructs both sides' full text before and after and returns its input
unchanged on any mismatch — the guard pattern from `diff/anchor.ts`.

## Consequences

| | before | after |
|---|---|---|
| C21 redundancy (corrected metric) | 0.94 | **0** |
| C21 marks | 2 | 3 |
| corpus total redundancy | 1.05 | **0.11** |
| corpus total marks | 82 | 83 |

C21 now reads as three small marks instead of two large ones, and the twelve
unchanged words are never shown twice:

```
good  9.26 The Purchaser shall notify the Vendor of any Third Party Claim[-.-]
      [-The-]{+, and the+} notice shall specify the nature of the claim in
      reasonable detail.
```

0.11 remains C07's score and the corpus floor.

**`idealMatch` for C21 stays `false`, deliberately.** The reviewer wrote the
wanted rendering as a *single* paragraph. A merge that the reader can still
**reject** cannot be one paragraph in the output: Word records the join as a
revision on the paragraph mark, and the two `<w:p>` elements must both exist
until that revision is accepted. Rendering it as one paragraph would mean
performing the merge unconditionally, which is a worse failure than the
cosmetic one. The gap between the wanted rendering and the honest one is a
**preview** question, not a diff question — the natural answer is to draw the
struck paragraph mark inline as a struck pilcrow in the RenderCore preview, the
way conventional redline tools present it, so the reader sees one flowing paragraph while the
document still holds two. That is not built here.

The reviewer's ideal also treats ` the notice` as unchanged where the original
has capital `The`; the pipeline marks the capitalisation honestly
(`[-The-]{+, and the+}`). Recorded rather than engineered around.

## Alternatives rejected

- **A new `merge` / `split` op type.** More surface in the model, the
  classifier, the author and every consumer, for something two existing range
  fields already express. Rejected as scope the spec does not ask for.
- **Glue the paragraphs in preprocessing and diff the joined text.** Loses the
  paragraph mark entirely — the merge becomes unrejectable, which breaks the
  point of a tracked redline.
- **Similarity-only pairing (Dice over the two paragraphs).** Bag-of-words
  overlap pairs neighbours that merely share vocabulary, which is how two
  unrelated paragraphs would get glued together. Containment is positional for
  that reason; Dice is used only for the leftover half, and only at the
  threshold ADR-0026 already sets.
- **Leave C21 alone and accept the metric's 0.** The metric was wrong, not the
  case. Fixing the yardstick first is what made the decision to build this
  measurable at all — before that correction the evidence pointed the other way.

---

## Correction — 2026-09-04 (independent spec-conformance review review)

A review of this diff found the decision sound and the headline numbers exactly
reproducible, but **four statements above were wrong about the code**, plus three
defects. Recorded here rather than silently edited, because the original claims
were recorded when the change was made.

**1. "In all four orderings" — only two exist.** `structuralDiff`'s `emitGap`
emits every `replace` in a segment before its orphans, and segments are
separated by an `equal` anchor, so an orphan is never the op immediately
*before* its replace. The `delete`+`replace` and `insert`+`replace` branches
were unreachable in principle and dead in fact — deleting them left the whole
suite green. **Both are now removed**, and the test named *"re-pairs delete +
replace, where the absorbed paragraph came first"* was renamed: its fixture
produced `replace`+`delete` and fired the first branch all along.

> **Note (2026-10-01, ADR-0039).** "An orphan is never the op immediately *before* its replace" is
> no longer true of Level 1 as a whole: similarity gap pairing can emit a re-paired `replace` after
> a positional surplus. The two branches stay removed, deliberately. ADR-0039's lock runs this pass
> over the positional ops and keeps every merge and split it finds, so an orphan-first shape arises
> only among records that were whole-paragraph churn before. ADR-0039 §1 records that as an
> accepted limit.
>
> **And one new refusal.** `detectMergesAndSplits` takes an optional `mayAbsorb(record, side)`
> predicate, and the pipeline passes ADR-0039's `GapPairing.mayAbsorb`. It refuses, as the
> absorbed orphan, a record whose text appears verbatim on the other side (one headed for a
> Level-3 move), unless positional pairing's own run of this pass absorbed it. Re-pairing could
> otherwise place such a record straight after a replace it never followed, and absorbing it there
> took its move away. Without the predicate (the lock probe's own run of it) the pass
> behaves exactly as before.

**2. "Each guard is ablation-verified" was false for two of them.** Honest
position, now recorded in the tests themselves:

| guard | status |
|---|---|
| `MERGE_MIN_TOKENS`, positional containment, remainder/Dice, table unit | pinned — ablating fails a test |
| single-paragraph orphan | **now** pinned, by a white-box op-list test |
| `mayAbsorb` (ADR-0039, 2026-10-01) | pinned, merge and split side each — `compare.pipeline.gap-pairing.test.ts` |
| seam-collision (new, see 4) | pinned |
| `host.length <= part.length` | **pinned** — see Correction (2). The "not separately observable" claim here was itself wrong: `dice()` (since ADR-0039 the shared `bagDice`, same convention) returns 1 when *both* sides are empty, so an empty `keep` lets the Dice check pass and this guard is the only brake left |
| strict adjacency | **not separately observable** — widening across a skipped record changes the set of records claimed, which the never-suppress guard detects independently |

Strict adjacency is **kept** as a stated precondition that also keeps
`endContainment` off a negative slice. What changed there is the claim, not the
code.

**3. `markParagraphMark` emitted schema-invalid `document.xml`.** §5.5 stated the
`CT_PPr` rule correctly (`rPr` before `sectPr`/`pPrChange`, one only); the code
appended before `</w:pPr>` unconditionally. Three reachable breakages, all on
shapes Word writes routinely: a **second `w:rPr`** when the paragraph mark
already carried direct formatting (or an emptied one left by §9.2 quarantine);
the mark placed **after `sectPr`** on the last paragraph of a section; and, worst,
a lazy `</w:pPr>` match landing the revision **inside `w:pPrChange`'s nested
`w:pPr`** — a `CT_PPrBase` with no `rPr` at all, so the paragraph mark was then
not marked and *the merge became unrejectable*, the exact failure this element
exists to prevent. Fixed with one depth-counting `matchLeadingPPr`, now shared
with `wrapParaRuns` and `splitParagraph`, which each carried their own copy of
the same lazy idiom (and the same latent flaw).

**4. `PARA_SEAM` is not "impossible to collide with document text".** `normaliseText`
does not strip Private Use Area code points, and PDF extraction with a broken
`ToUnicode` map — or verbatim pasted text — yields them. On such a record the
author cut at the document's *own* U+E000, emitting an extra `<w:p>` carrying **no
revision** (invisible and unrejectable) and dropping the character with **no
`<w:del>`** — an ADR-0008 breach the record-level never-suppress guard cannot
see, since it compares which records are claimed, not their text. The pass now
**declines to re-pair any record whose text contains the sentinel**.

**5. The sentinel leaked out of the diff into four consumer surfaces.** The
author strips it, but `op.wordDiffs` retained it, and every other consumer reads
that field: the AI-export change set (whose consumer is an LLM), the exported
compare report, the change-list row preview (rendered as tofu), and
`goToChange`'s host-search anchor — which could never match document text, so
**"go to change" was a silent no-op for every merged or split op**. The seam now
stops at the author, spelled as the plain newline it stands for.

**6. The corrected metric had a false positive.** The cross-paragraph measure
this ADR introduced was written as an LCS against *every* insertion anywhere in
the change, though the text above says "reappears **verbatim**". An honest
insertion that merely reuses the document's defined terms, beside an unrelated
whole-paragraph deletion, scored **0.47** on shared drafting vocabulary alone.
Fixed by counting one *contiguous* run.

That fix originally came with a second narrowing — to *adjacent* paragraphs
only — which was **wrong and has since been reverted**; see Correction (2).

Suite after the corrections: **1722 tests + 1 todo**, green.

---

## Correction (2) — 2026-09-04 (second independent spec-conformance review review)

The first correction closed seven findings. A second review of that work found
three of the closures incomplete or wrong, and two further defects in the
yardstick. Every fix below is ablation-verified — each named ablation was
performed, and the stated test failed.

**A. The `goToChange` fix in §5 replaced a silent failure with a wrong one.**
Spelling the seam as a newline stopped the anchor holding U+E000, but
`goToChange` still passed the op's first *mark* to the host search — and the
first mark of a merge is the lost sentence break, i.e. `"."`. `findAndReveal`
selects `results.items[0]`, so "go to change" jumped to **the first full stop in
the document**. A no-op became a mis-navigation, which is worse. Replaced with an
exported `searchAnchorFor`: the first changed line carrying at least four letters
or digits, else the longest unchanged line, else `''` (do nothing). Pinned by
`tests/unit/compare.driver.anchor.test.ts`; three of its five tests fail under
the previous behaviour.

**B. `matchLeadingPPr` had no test at all.** Replacing its body with the old lazy
regex left the entire suite green — the depth-counting was asserted, never
exercised. `markParagraphMark`'s share was covered; `wrapParaRuns` and
`splitParagraph` were not. Now covered on a `w:pPrChange`-bearing paragraph (any
paragraph a reviewer restyled under track-changes), including an explicit
well-formedness check: under the lazy regex `wrapParaRuns` emits a `<w:del>` with
a close tag it never opened.

**C. `markParagraphMark` mis-read one more `rPr` shape.** The paired-tag probe
`/<w:rPr(?:\s[^>]*)?>/` matches a self-closed `<w:rPr w:x="1"/>` — `[^>]*`
swallows the slash — putting `w:del` directly under `CT_PPr` and leaving the
paragraph mark unmarked. The self-closed branch also *discarded* the tag's
attributes rather than expanding it. `CT_ParaRPr` declares no attributes, so this
is not reachable from Word's own output, but both failures were in the dangerous
direction. Fixed and pinned.

**D. `spellOutSeams` was unconditional.** It rewrote U+E000 to a newline in
*every* op, so a document carrying its own PUA code point — the very case
finding 4 above exists to handle — had it silently turned into a paragraph break
in the AI-export change set, the exported report and the change-list preview. Now
confined to ops claiming more than one record, which is sound precisely because
`isPlain` declines to re-pair a record containing the sentinel. The
"can never collide with document text" sentence in `merge-split.ts` is retracted
in place.

**E. The metric scored contentless struck spans at the maximum.** `floor` is
`min(struck, inserted)`, so a **one-character** struck span whose character also
appears in the insertion beside it scores **1.00**. This is live for the shape
this workstream is trying to reach: a paragraph join renders the lost break as
`[-.-]`, and if the inserted joining words end in a full stop, the *wanted*
rendering scores the maximum while the bad one scores 0. No corpus case hits it,
which is why it had to be found by construction. Now guarded, corpus unmoved at
**0.11 / 83 marks / 8-of-9 ideals**.

**F. The "adjacent paragraphs only" narrowing traded a false positive for a
false negative, on the exact shape this ADR exists to catch.** When the paragraph
between a merge's two halves is *itself* deleted, the re-insertion lands two
paragraphs away: the reader meets the clause twice and the measure scored it
**0.11** instead of **0.91**. Distance is not the discriminator — contiguity is.
The distance limit is removed. The measure's three decisions (contiguity, no
distance limit, the contentless-span guard) are each pinned by a test in
`compare.readability.test.ts` that fails when that decision alone is undone; the
earlier regression test pinned **none** of them, scoring 0 under both the old and
new measures.

The whole-for-whole swap exemption remains a **disclosed blind spot**: a
near-merge this pass declines and ADR-0026 then demotes lets the reader see most
of a paragraph twice and scores 0. The exemption is principled (P2 admits that
rendering), so narrowing it is a measurement-design question for the reviewer,
recorded in the plan's open questions.

Suite after Correction (2): **1737 tests + 1 todo**, green.
