# ADR-0041: Compare — a large contiguous cut is not a rewrite (contiguity escape from the change-complexity demotion)

- **Status:** Accepted
- **Date:** 2026-09-12
- **Relates to:** ADR-0026 (change-complexity threshold — the rule this narrows); ADR-0008 (categorise, never suppress); ADR-0038 (the real-document instrument that measured it); feature 08 (`spec/features/08-compare.md` §5.3)

## Context

ADR-0026 demotes a matched `replace` paragraph pair to a whole-paragraph
delete + insert when the two paragraphs are too dissimilar to read inline. The
measure is Sørensen–Dice over content tokens, and the threshold is 0.5. The
rule is right, and it is the legal-comparator convention: a near-total rewrite
shown inline is "a few incidental shared words as `equal` islands in a sea of
ins/del", which no reader can use.

**Dice is a set measure, so it reads length asymmetry as dissimilarity.** And
the commonest heavy edit in drafting — cutting a block of text out of a clause
— is nothing *but* length asymmetry. The survivor is unchanged, verbatim, and
contiguous; Dice still scores the pair low, because one side simply has far
more tokens than the other.

This was found by benchmarking the engine on a clean private `v1`/`v1.1` pair of
the same real-world contract (not published) — ordinary drafting edits, no fixture
contamination, with a conventionally styled redline of the same two files
alongside. Two of the pair's sixteen paragraphs were
demoted:

| paragraph | orig tokens | rev tokens | Dice | contiguous survivor |
|---|---|---|---|---|
| two of three sentences cut | 78 | 17 | **0.358** | **100 %** of the shorter side |
| a ~100-word middle cut | 191 | 60 | **0.478** | **96.7 %** of the shorter side |

Both were rendered as a whole paragraph struck beside a whole paragraph
inserted. **A conventional redline keeps both inline**, and so does any reader's
intuition: the second one is a single strike-through with two sentences either
side of it untouched.

The cost is exactly what P1/P2 measure. On this pair, before the fix:
`wholeParaChurn` **2**, `alignmentRate` **0.880**, corpus-scale `redundancy`
**0.8** — the reader re-reads every surviving word of both paragraphs to learn
that nothing happened to them.

## Decision

**A `replace` pair that fails the similarity test is still kept for inline
word-diffing when a large enough share of the SHORTER paragraph survives as one
unbroken run of content tokens in the other.**

- `longestContiguousRun(a, b)` — the longest **contiguous** common run of
  content tokens, by a rolling two-row DP.
- `contiguousShare(a, b)` — that run divided by the token count of the
  **shorter** side.
- `config.replaceContiguityThreshold`, default **0.8**. Values > 1 disable the
  escape; the similarity test and its own threshold are untouched.

Contiguity is the point, and an LCS would be the wrong measure: an LCS happily
stitches a shared word here to a shared word forty tokens later, which is
precisely the scattered shape the demotion exists to catch. Only an unbroken
run reads as a block the eye can skip.

The **shorter** side is the denominator on purpose. A large deletion leaves a
short survivor wholly inside a long original; scoring against the longer side
would report ~0.2 for a rendering that is one clean strike-through — the same
asymmetry Dice already gets wrong, reintroduced.

## Rationale — the threshold is measured, not chosen

The risk of an escape is that it re-admits the salad. So the two populations
were generated and measured, 4,000 pairs of each, from the sweep's own clause
vocabulary:

| population | Dice p50 | contiguous share p50 | p99 | max |
|---|---|---|---|---|
| **genuine rewrites** (two independently generated clauses — demotion is correct) | 0.34 | 0.15 | **0.68** | **0.81** |
| **clean large cuts** (a contiguous prefix or suffix kept — demotion is wrong) | 0.67 | **1.00** | 1.00 | 1.00 |

The separation is almost total. Of the 3,240 genuine rewrites that Dice demotes
today, a 0.8 contiguity escape re-admits **zero**. Of the 1,309 clean cuts that
Dice wrongly demotes, it rescues **all 1,309** — every one of them scores
exactly 1.00, because a contiguous cut leaves a contiguous survivor by
construction.

0.8 sits well above the rewrite population's 99th percentile (0.68) and well
below both paragraphs (0.967 and 1.000). It is loose enough to be robust
to real prose, which repeats boilerplate more than the generator does, and
tight enough that the worst rewrite observed anywhere (0.81) would need to be
*more* self-similar than any of 4,000 to slip through.

## Consequences

Measured on the clean `v1`/`v1.1` pair, before → after:

| | before | after |
|---|---|---|
| `wholeParaChurn` | 2 | **0** |
| `alignmentRate` | 0.880 | **1.000** |
| `redundancy` | 0.8 | **0** |
| changed paragraphs | 7 | **5** |

**Nothing else moved.** The 33-case synthetic corpus is unchanged (total
redundancy 0.11, 89 marks, the same 8-of-9 reviewer ideals), and the 4,000-pair
generalisation sweep still passes with `KNOWN_OFFENDERS` untouched — as it must,
since that sweep exercises `wordDiff` directly and never reaches this pass.

**Where a sentence is removed after a word**, the engine keeps that word untouched
and shows two clean deletions, rather than striking `of Disputes.` and re-inserting
`Disputes` merely to say that a sentence went. The remaining difference in
whole-document mark count against the conventional redline is entirely the two
style-inherited numbering changes the engine cannot yet see (see Open).

**ADR-0008 is untouched.** This changes only which of two renderings of the same
two paragraphs is chosen. A kept pair shows *less* markup than a demoted one, so
the never-suppress direction is the one already argued in ADR-0026: the demoted
form is the one with more markup, and declining to demote never hides anything —
`reconstruction` on both sides of the private pair still passes.

**ADR-0026 is narrowed, not superseded.** Its rule, its measure and its
threshold all stand; a second, independent condition now sits beside the first.

## Rejected alternatives

- **Lower the Dice threshold.** The obvious move and the wrong one. Both
  paragraphs sit at 0.358 and 0.478, so the threshold would have to fall below
  0.36 — which is the *median* of the genuine-rewrite population (0.34). It
  would re-admit roughly half the salad to rescue the cuts. The two shapes are
  not separable on Dice at any threshold; that is why a second measure is
  needed rather than a different number.
- **Replace Dice with the overlap coefficient** (`overlap / min(|A|,|B|)`),
  which is asymmetry-tolerant by construction. It scores 1.0 whenever the
  shorter side's vocabulary is a subset of the longer's — including for a short
  paragraph of entirely new drafting built from common legal words, which is a
  genuine rewrite. Contiguity distinguishes them; set containment cannot.
- **Count the marks the inline diff would actually produce, and demote on
  that.** The direct measure rather than a proxy, and tempting. Rejected on
  layering: this pass runs *before* the word diff precisely so a demoted pair is
  never word-diffed, and inverting that would word-diff every below-threshold
  pair twice — a real cost in the at-scale batch paths (ADR-0037). Contiguity is
  a structural property of the two texts that needs no diff and cannot drift
  when the diff is tuned.
- **Require a minimum absolute run length as well as a share.** Considered for
  the pathological case of a two-token paragraph wholly contained in a long one.
  Measured: that case is a near-total deletion, and `The Vendor[- …198 tokens…-]`
  is the honest rendering of it — two words did survive. Adding a floor would
  demote it to restating both paragraphs, which is worse and less true.
- **Detect the underlying edit as a MOVE instead.** For the first paragraph the
  deletion is really a relocation — the cut sentences reappear at the end of the
  preceding paragraph. Rendering it as `moveFrom`/`moveTo` would improve on this
  engine's rendering. Cross-paragraph move detection is a different
  decision with its own reserved number (ADR-0039) and its own alignment work;
  it is not a reason to leave this pair demoted in the meantime.

## Open

- **Style-inherited numbering changes are still invisible**, and they are the
  whole of the remaining gap to the reference on this pair. Neither document
  carries a single `w:numPr`: the visible clause numbers come from the style
  chain, and the actual edit was a `w:pStyle` change on one paragraph
  (`<subclause-L3>` → `<subclause-L2>`), which renumbered it from `(vi)` to
  `(d)` and pushed the paragraph after it from `(d)` to `(e)`. The two
  paragraphs are text-equal, so they never reach a classifier and no op is
  emitted at all. `ParaRecord.styleId` is captured and consumed nowhere. This is
  the style-chain numbering resolver case, which until now had no reproduction on
  any real document and now has one. It is not fixed here.
