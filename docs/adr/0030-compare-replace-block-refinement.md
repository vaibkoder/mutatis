# ADR-0030: Compare — replace-block refinement (undo the cleanup avalanche where a finer rendering earns its marks)

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0026 (change-complexity threshold); ADR-0028 (edit-boundary anchoring); ADR-0029 (cross-reference tokenisation); feature 08 (`spec/features/08-compare.md` §5.3)

## Context

The worst-reading case in the 30-case readability corpus — the only one the
reviewer graded **bad** — was a collapsed doublet:

```
bad   The Vendor shall indemnify [-and hold harmless the Purchaser against any
                                   and all Losses whatsoever.-]{+the Purchaser
                                   against all Losses.+}
```

Twelve struck words beside six inserted ones, and the reader has to compare the
two spans word by word to find the three small deletions actually made. Its
`redundancy` score was **1.00** — every word of the shorter side appears
verbatim in the longer one. C21 (a paragraph merge) scored the same, C03 0.67,
C12 0.41.

The cause is `diff_cleanupSemantic`. Its rule — *eliminate an equality smaller
than or equal to the edits on both sides of it* — measures the equality with
`String.length`, and `diff/word.ts` runs it in **sentinel** space, where one
character is one **token**. The tokenizer emits every run of whitespace as a
token of its own, so `all ` counts 2 and a lone space counts 1 against a
six-token deletion. Real shared prose is out-weighed and destroyed; the
neighbouring edits grow by what they swallowed; the next equality along becomes
eligible in turn. That avalanche is what turns three small deletions into one
whole-paragraph swap.

Deleting the cleanup outright is not the answer: it was measured and leaves the
`[-a-] {+b+}` two-adjacent-replaces shape that reads worse than the fusion it
avoids (ADR-0028's rejected alternatives), and it fragments C12.

## Decision

Keep `diff_cleanupSemantic`, and add a pass that **undoes its over-fusion
locally, where the finer rendering earns its extra marks**. `diff/refine.ts →
refineReplaceBlocks` takes each struck span sitting immediately beside an
inserted one and re-diffs the two sides on **real tokens** with no cleanup at
all, then applies two acceptance rules:

1. **It must recover a shared word.** Recovering nothing but whitespace buys the
   reader nothing and costs two extra marks. `[-out of any-]{+from each+}` is one
   rewritten phrase; splitting it at the space it happens to share into
   `[-out of-]{+from+} [-any-]{+each+}` reads as two editing acts, not one (P3).
2. **It must not hand the reader the tool's job.** A word struck inside one mark
   and re-inserted inside another is text the reader has to prove identical
   before they can dismiss it — the P2 failure that makes fragmenting a replace
   worse than leaving it coarse. A transposition therefore stays coarse; phrase-
   scale move detection, not fragmentation, is its answer.

When the re-diff is rejected, one fallback is tried: pull out only the run the
deletion **ends** on and the insertion **begins** with, so that shared run
belongs to neither mark. DMP has this move (`diff_commonOverlap`) but gates it on
a character-count ratio a token-space diff never reaches. It is what recovers
`arising` in C12.

Two supporting details. A trailing-punctuation tail is split off each token
locally, at the block only (`Losses.` → `Losses` + `.`), never for an atomic
value — the same local-split device ADR-0028 uses at a boundary, and the reason
C11 can end on `Losses[- whatsoever-].` rather than `[-Losses whatsoever.-]`.
And sentinels are keyed on **atomic kind + text**, matching `diff/word.ts`, so a
reference value never aligns with an identical-looking ordinary word (ADR-0029).

The pass runs **before** `slideEditBoundaries`, so the single-sided deletions and
insertions it frees are anchored where they read best.

## Consequences

Measured over the corpus, before → after:

| case | grade | before | after | rendering after |
|---|---|---|---|---|
| C11 doublet collapsed | **bad** | 1.00 | **0** | `indemnify [-and hold harmless -]the Purchaser against [-any and -]all Losses[- whatsoever-].` |
| C21 paragraphs merged | partial | 1.00 | **0** | insertion into the first paragraph; the second struck whole |
| C03 `;`→`.` + `it`→`It` | partial | 0.67 | **0** | `basis[-;-]{+.+} [-it-]{+It+} survives` |
| C12 phrase → longer phrase | partial | 0.41 | **0** | `[-howsoever -]arising {+directly or indirectly out of or +}in connection` |
| **corpus total redundancy** | | **3.19** | **0.11** | |

0.11 is C07's score and the floor the corpus can reach — a genuine four-swap
sentence where one shared space lands inside a mark. **The corpus is now at that
floor.** This pass alone raises marks 82 → 84, which is the intended direction:
three honest deletions beat one confusing block.

The table above measures **this pass in isolation**, which is how it was
developed. It did not ship alone: ADR-0031 (adjacent transposition) landed in the
same commit and took C08 from 4 marks to 2, so the corpus **as this ADR landed**
reads redundancy **0.11**, marks **82**, and **eight of the nine** renderings the
reviewer wrote down as wanted produced exactly (C03, C08, C11, C12, C16, C23,
C24, C29). `idealMatch` is measured and pinned in the conformance suite, so a
case that reaches the wanted rendering can never quietly fall back off it.

**ADR-0008 is untouched.** The pass only re-partitions the same two texts. It
reconstructs both sides before and after on every call and returns its input
unchanged if they ever differ — the same never-suppress self-check
`slideEditBoundaries` carries.

## Alternatives rejected

- **Drop `diff_cleanupSemantic` entirely.** Measured: the corpus reaches the same
  redundancy but with 88 marks instead of 84, i.e. more fragmentation for no
  reading gain, and it re-introduces the `[-a-] {+b+}` shape ADR-0028 rejected.
- **A bounded replacement for `diff_cleanupSemantic`** that merges only across an
  equality carrying no word. Prototyped and measured: it changes **nothing** on
  the corpus, alone or combined. Deleted rather than kept as untested weight.
- **Accept any candidate that recovers an equality.** Metric-optimal — it drives
  the corpus to 0.00 — but only by splitting C07's one rewritten phrase into two
  marks over the space between them. The reviewer had already graded C07 *good*;
  optimising the metric past the reader is the failure mode the metric exists to
  avoid.

## Open

- **C21** (true paragraph merge) is the one case still differing from the
  reviewer's wanted rendering. It needs an n:1 structural match rather than
  anything in the word diff, tracked in
  the readability workstream. (C08, the
  transposition, was the other outstanding case while this pass was being
  developed; ADR-0031 resolved it in the same commit.)
- The `selfDiffing` guard is not exercised by any corpus case — it is pinned by a
  direct unit test instead (`tests/unit/compare.diff.refine.test.ts`), since it
  guards a failure mode the corpus does not happen to contain rather than one it
  cannot contain.
