# ADR-0031: Compare — adjacent transposition reads as one relocation, not two replacements

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0028 (edit-boundary anchoring); ADR-0030 (replace-block refinement); feature 08 (`spec/features/08-compare.md` §5.3, §5.4)

## Context

Reordering two words in a list is one editing act. The pipeline rendered it as two:

```
bad   The Vendor shall [-indemnify-]{+defend+}, [-defend-]{+indemnify+} and hold harmless …
```

The reader has to notice, unaided, that the two struck words are exactly the two
inserted words in the other order. This is the classic LCS/Myers blind spot for
adjacent transpositions — the same failure `git diff` and `difflib` show. Myers
maximises the retained common subsequence, so it keeps the separator between the
two words as an equality and reports two independent one-word substitutions.

It is **not** a boundary-placement problem, so ADR-0028's `slideEditBoundaries`
cannot reach it: that pass only ever moves a *single-sided* edit flanked by
equalities, and a transposition is a two-sided replace twice over.

Nor is it a granularity problem. ADR-0030's refinement declines it deliberately:
every finer rendering of a transposition strikes a word in one mark and
re-inserts it in another, which is the P2 failure that makes fragmenting a
replace worse than leaving it coarse. Refinement leaving it alone is correct;
something else has to recognise the shape.

P7 says a move reads as a move — *at paragraph scale and at phrase scale*.
`detectMoves` (§5.4) delivers the paragraph half. This is the phrase half.

## Decision

`diff/transpose.ts → detectTransposition` recognises the shape directly in the
coalesced word diff — `del(A) ins(B) eq(mid) del(B) ins(A)`, the same two spans
changing places around a fixed middle — and re-expresses it as **one**
relocation:

```
good  The Vendor shall {+defend, +}indemnify[-, defend-] and hold harmless …
```

One span moves across the connective; the other is left as the sole survivor and
is never marked at all. It stays an ordinary insert-here / delete-there pair:
P7 "at phrase scale" does not require Word's native `w:moveFrom` / `w:moveTo`,
whose range-start/range-end pair and shared `w:name` buy nothing a reader can see
inside a single sentence.

Guarded narrowly, on **structure alone** — no vocabulary, no semantics, no
equivalence scoring (ADR-0008):

- both swapped spans must be short (≤ 4 non-whitespace tokens). A long-phrase
  swap already reads fine as two clean phrase-for-phrase replacements, and
  re-expressing it moves a whole clause across the page;
- the connective must carry **no word** — whitespace and punctuation only. The
  rewrite duplicates it: `mid` is inserted alongside B and struck alongside B's
  old copy. Duplicating `, ` is invisible, but duplicating ` and the ` would make
  the reader prove two marks identical, which is exactly the P2 failure ADR-0030
  rule 2 refuses to commit. The two ADRs have to agree about P2, so
  `A and the B` → `B and the A` is deliberately left as two clean one-span
  replacements (`by the [-Vendor-]{+Purchaser+} and the [-Purchaser-]{+Vendor+}`);
- the connective must be short (≤ 8 characters), which bounds how far apart the
  spans can sit and still read as *these traded places* rather than two
  coincidentally-matching unrelated edits. The bound is in characters, not
  tokens: the connective carries no word, so its cost to the reader is the
  distance it puts between the spans — `, ` is nothing, the dotted leader of a
  schedule or signature block is too far;
- the quintuple must cover **whole** del/ins runs — the entry before it may not
  be a `del`, the entry after may not be an `ins`. `coalesce` leaves adjacent
  same-op entries in place wherever one of them is atomic, so a five-entry window
  can otherwise be a slice through the middle of a longer span, and "A moved"
  would be claimed for only part of what actually moved;
- the two spans must be verbatim-identical at both occurrences, and must differ
  from each other.

Each span keeps the atomic kind it arrived with, so a swap of two amounts still
classifies as a numerical-value change (P10). The relocated span and the
connective are emitted as separate entries for that reason, and `coalesce` then
joins them wherever neither is atomic — an ordinary word swap becomes literally
one entry. Where the span *is* atomic, `coalesce` correctly refuses to merge
(that would destroy the kind), so the diff keeps two adjacent same-op entries;
the renderer emits adjacent same-op entries as one continuous mark, so the reader
sees one mark either way. The reader-facing outcome does not depend on which
branch applies.

The pass runs last in `wordDiff`, after anchoring and coalescing. Like every
other pass in the chain it is **length-preserving and self-verifying**: it
reconstructs both sides before and after and returns its input unchanged on any
mismatch (ADR-0008).

## Consequences

C08 — the reviewer's transposition case — drops from 4 marks to 2 and now
produces the rendering they wrote down, exactly. Corpus totals: redundancy
unchanged at **0.11** (the floor), marks **84 → 82**, no other case moved, no
invariant violated. Tightening the connective to noise-only and adding the
whole-runs check moved nothing on the corpus: C08's connective is `, `.

**Eight of the nine renderings the reviewer recorded as wanted are now produced
exactly** (C03, C08, C11, C12, C16, C23, C24, C29). The one outstanding is C21, a
true paragraph merge, which needs an n:1 structural match rather than anything in
the word diff.

## Alternatives rejected

- **Word-native `w:moveFrom` / `w:moveTo` for the phrase.** Heavier markup —
  range-start/range-end plus a shared `w:name` — for no visible difference inside
  one sentence, and it invites Word's move-tracking UI to treat an intra-sentence
  reorder as a document-scale move.
- **Let ADR-0030's refinement fragment it.** Measured and rejected there: every
  fine-grained rendering makes the reader prove two words identical across marks
  (P2). The refinement's guard exists precisely so this case reaches here intact.
- **Re-score `slideEditBoundaries` to reach it.** Structurally impossible: a
  two-sided replace has no free boundary to slide.
