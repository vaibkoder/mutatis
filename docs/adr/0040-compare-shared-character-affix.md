# ADR-0040: Compare — recover a shared character affix, and run the refinement to a fixed point

- **Status:** Accepted
- **Date:** 2026-09-12
- **Relates to:** ADR-0008 (categorise, never suppress); ADR-0030 (replace-block refinement — this adds a third branch to its acceptance rule); ADR-0029 + its 2026-09-12 correction (the possessive apostrophe); ADR-0038 (the instrument that made this measurable); feature 08 (`spec/features/08-compare.md` §5.3)

## Context

ADR-0030 gave `refineReplaceBlocks` two branches: re-diff the struck and
inserted sides on real tokens and keep the finer rendering if it recovers a
shared **word**; failing that, pull out the run the deletion ends on and the
insertion begins with (`overlapSplit`).

All three of its guards — `recoversWord`, `selfDiffing`, `overlapSplit` —
compare **token identity** through `keyOf`. None of them can see a shared
*character* affix inside a token.

This is the defect that surfaced on the private real-world test pairs. Instrumented on those fixtures, the
`faith` site reached `refineReplaceBlocks` and was **accepted**:

```
del = "good faith"
ins = "[-good faith-][+sole discretion+]"
candidate: ins:"[-" | eq:"good " | del:"faith" | ins:"faith-][+sole discretion+]"
  → recoversWord = true   selfDiffing = []   ACCEPTED
```

The struck token is `faith`; the inserted token is `faith-][+sole`.
Different keys, so no self-diff is detected, and the reader is shown
`faithfaith`. Where the guards instead **decline** (`Statements`, `Adjustment`,
`therewith`, `procedures`) the block is left coarse, which shows the same
duplication anyway. **Both branches of the acceptance rule produced visible
duplication, and there was no third option.**

Neither yardstick could see it either: `readerTokens` is derived from the
production tokenizer, so the metric also saw two unrelated tokens and scored
**0**. That is why ADR-0038's character yardstick had to land first.

## Decision

### 1. A third branch: recover a shared character affix

When the re-diff and `overlapSplit` are both declined, `affixSplit` fires in one
narrow shape — **one side's text is wholly a prefix or a suffix of the other's**.
The shared run then becomes `eq` and the remainder becomes a single `ins` or
`del`; one of the two marks disappears entirely rather than being split in two.

```
[-faith-]{+faith and fair dealing+}   →   faith{+ and fair dealing+}
[-Purchaser-]{+Purchaser's agent+}    →   Purchaser{+'s agent+}
[-fee-]{+reasonable fee+}             →   {+reasonable +}fee
```

Whole-side containment is a guard, not an implementation convenience. `Vendor`
and `Purchaser` share a final `r`; recovering it would render
`[-Vendo-]{+Purchase+}r`, a span that begins mid-construction, which is machine
residue and costs the reader a second pass (P4).

### 2. The guards that protect a rendering a human already graded good

**The branch fires only when the extra material is new WORDS, not the rest of
the same word, and only when the cut itself falls between words.** That is
*two* conditions, and the first draft of this ADR shipped only the first:

- `addsNewWords` requires the remainder to carry whitespace or a character
  outside the tokenizer's word-run class.
- `splitsBetweenWords` requires that the characters on **both sides of the cut**
  are not both letters-or-digits.

**Correction, before first push — the second guard was missing and the first is
not sufficient on its own.** `addsNewWords` scans the *whole* remainder, so a
single space anywhere licenses the cut. `Loss` → `Losses and Damages` has that
space, and the cut still lands mid-way through `Losses`:

| input | rendering with `addsNewWords` alone |
|---|---|
| `any Loss arising` → `any Losses and Claims arising` | `Loss{+es and Claims+}` |
| `A Director shall` → `A Directors and officers shall` | `Director{+s and officers+}` |
| `the fee` → `the fees of the adviser` | `fee{+s of the adviser+}` |
| `A holder may vote` → `A the Shareholder may vote` | `{+the Share+}holder` |

`Loss{+es …+}` is verbatim the rendering this section says the guard exists to
prevent, reached from the other side; `{+the Share+}holder` is the
begins-mid-construction residue §1's containment rule is justified by. Both were
found by the two spec-conformance review passes on this diff, independently of each
other, and neither the 33-case corpus nor the three private real-world test documents contained the
shape — a scan of every rendered mark boundary in all 36 found **zero**
alphanumeric-on-both-sides cuts, so no published figure in this ADR was affected.

The correct test is on the **pair of characters straddling the cut**, not on the
remainder as a whole: a word is cut in half exactly when both of them continue a
word. Testing one side only would fail the other way — `pay the ` →
`pay the reasonable ` cuts after a space and reads perfectly.

`WORD_BODY` is `\p{L}\p{N}` and is deliberately **narrower** than the
tokenizer's `WORD_CHAR`, which also holds `'`, `-`, `.` and `&`. A boundary test
written on `WORD_CHAR` would refuse `Purchaser{+'s agent+}` and
`faith{+-][+…+}` and withdraw the branch's entire purpose. A reader seeing
`Purchaser{+'s agent+}` reads two things; a reader seeing `Loss{+es+}` reads a
word cut in half.

The two guards are independent and each has its own ablation: only
`addsNewWords` refuses `Purchaser` → `Purchaser's` (an apostrophe is not a
letter, so the boundary test is happy to cut there), and only
`splitsBetweenWords` refuses the four shapes in the table above.

This is the most consequential decision here, and it **deliberately does not
meet the original acceptance criterion for this change**, which asked that
`Statement` → `Statements` render as a suffix insertion. It cannot, because the
corpus already contains that shape as **C02**, graded **good** by the reviewing
lawyer, with the reason written down:

> *"Whole word Loss/Losses swapped, not `Loss` + inserted `es`."*

`Loss{+es+}` is the character-level interleave P5 exists to forbid, and taking
it would be buying a metric against a human verdict on a case that verdict
examined directly. The two shapes are not the same defect:

| | shape | verdict |
|---|---|---|
| `Loss` → `Losses` | one word changing its ending | graded **good** (C02) — leave coarse |
| `faith` → `faith-][+…+]` | a word restated because a clause was appended | the observed defect — recover |

So C32 (`Statement` → `Statements`) stays a whole-word swap and keeps
`charRedundancy` 1.00. **This is recorded as an open question for the reviewer,
not as a defect closed** — see Open, below.

### 3. Atomic values are still indivisible

`splitTokensAt` refuses when the affix boundary falls strictly inside an atomic
token. `30` is wholly a prefix of `30 days`, so without the refusal the branch
would render `30{+ days+}` and cut a period in half. Ablated and watched to
fail.

There was a **second**, belt-and-braces `atomicKindOf` re-check over the short
side. It has been removed as unreachable: every token in a block either arrives
from `tokenizeTagged` already carrying its kind, or from `splitTails`, which only
splits a token `atomicKindOf` has already declined. Measured, not argued —
16,339 `affixSplit` calls over a 20,000-pair fuzz, plus the whole suite including
the three private test documents with the branch turned into a `throw`: it fired **zero**
times, while `splitTokensAt`'s live refusal fired 1,068. A guard that cannot fail
cannot be ablated, and an un-ablatable guard is decoration that reads like
protection.

### 4. The pass runs to a fixed point

One pass is not enough, and the reason is the defect itself: **an accepted
re-diff can leave a new replace block inside its own output.** The `faith` trace
above recovers `good ` and leaves `faith` struck beside `faith-][+sole`
— exactly the shape the affix branch answers, and a single pass never looks at
it. When the branch first landed, the private fixtures still showed `faithfaith`
for precisely this reason.

`refineReplaceBlocks` now re-enters its own output until the rendering stops
changing, capped at four passes. Each iteration applies the same acceptance
rules to a strictly smaller block, so it terminates on its own; the cap guards
against a future rule that oscillates, not an expected limit.

**ADR-0008 is untouched.** The pass still only re-partitions the same two texts,
and the length-preserving self-check now runs over the fixed point against the
original input, returning the un-refined diff unchanged if it ever failed.

## Consequences

Measured, with 7b (the possessive apostrophe) already in place.

**The synthetic corpus does not move**: total redundancy **0.11**, marks **89**,
the same **8 of 9** reviewer ideals, `atomicIntact` and `numbered` green for all
33 cases. **C02 is unchanged**, which is the point of §2.

**The 4,000-pair generalisation sweep is byte-identical**: **66** offenders,
mean redundancy **0.0388**, mean marks **2.3287** over 3,675 pairs. `KNOWN_OFFENDERS`
was not targeted and did not move.

**The private test documents**, `charOffendingBlocks`, with ADR-0038's contamination
caveat attached to every figure:

| | before Phase 0.5 | after 7b | after 7c |
|---|---|---|---|
| pair3 | 13 | 11 | **4** |
| pair4 | 6 | 5 | **3** |

`wholeParaChurn` and `alignmentRate` are unchanged on all three pairs, as they
must be: this is a word diff and they are properties of paragraph pairing.
Reconstruction still passes on all three pairs, both sides.

**What remains, stated plainly.** The seven surviving blocks are two shapes, and
neither is closed here:

- **Five are the annotation shape**, where the struck word sits in the *middle*
  of the insertion because the revised document literally contains
  `[-GAAP-][+the Vendor's discretion+]` as body text. Recovering a word from
  the middle of an insertion means fragmenting that insertion into two marks —
  the general `del(X) eq ins(Y)` protection that **ADR-0034 measured and
  rejected** as "the metric-optimal option and the reader-hostile one" (0.0212
  redundancy, 17 offenders, marks 2.33 → 3.48, C07 and C12 both failing). It is
  not re-opened on the strength of contaminated fixtures.
- **Two are partial affixes** (`24]` against `24-][+four (4)…`), where neither
  side is wholly contained in the other. Handling them needs the same
  mid-span machinery.

## Rejected alternatives

- **Render `Statement` → `Statements` as a suffix insertion**, as the original
  acceptance criterion asked. Rejected: C02 is the identical shape, graded **good**
  by the reviewing lawyer with an explicit note rejecting exactly that
  rendering. Recorded as an open question rather than decided unilaterally.
- **Recover any shared character run, wherever it sits.** This is ADR-0034's
  rejected "protect any `del(X) eq(mid) ins(Y)` window" in a new costume:
  metric-optimal and reader-hostile, fragmenting every ordinary rewrite that
  straddles a shared word.
- **A minimum affix length instead of whole-side containment.** An arbitrary
  knob, tuned against the cases it judges. Whole-side containment is a property
  of the two strings, not a threshold.
- **Recover the affix without the `addsNewWords` guard.** Measured: it breaks
  C02 and C32 and four metric tests, and it is the P5 interleave.
- **Make `keyOf` character-aware instead.** It is the sentinel key shared with
  `diff/word.ts` (ADR-0029); making it fuzzy would let a reference value align
  with a look-alike word, which that ADR added it to prevent.
- **A single pass, accepting that an accepted re-diff leaves a residual block.**
  Measured: the private fixtures still show `faithfaith`. Ablated and watched to
  fail.

## Open

- **C32 / C02 — does a word gaining a suffix want a whole-word swap or a suffix
  insertion?** The reviewer has answered for `Loss` → `Losses` (whole-word) and
  the original acceptance criterion assumed the opposite for `Statement` →
  `Statements`. They are the same shape and cannot both hold. The whole-word
  reading is in force because it is the one a human actually looked at. If the
  reviewer prefers the suffix insertion, the change is to drop `addsNewWords` —
  one guard, one ablation, and C02's grade would need re-recording. Note that
  `splitsBetweenWords` would still stand: dropping `addsNewWords` buys
  `Statement{+s+}` and `Purchaser{+'s+}`, and does **not** re-open
  `Loss{+es and Claims+}`, which is a separate and unambiguous defect.
- The seven remaining blocks in the private documents above, which need clean fixture
  variants before they can be judged at all.
