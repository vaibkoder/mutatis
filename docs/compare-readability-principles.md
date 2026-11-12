# Compare — redline readability principles

Ten principles for what makes a redline *intuitively readable*, derived from a 22-case
indemnity-amendment corpus and **reviewed by a practising SG / E&W lawyer on
2026-09-04**. The target is conventional-redline-grade readability: a reviewing lawyer should
be able to see what changed without diffing the redline in their head.

These are not aspirations. They are the acceptance criteria the Compare rendering is judged
against, and they are enforced objectively by `tests/unit/compare.readability.test.ts` over
the corpus in `tests/fixtures/readability-corpus.ts`.

---

**P1 — Mark the smallest span that carries the change.**
A one-word edit marks one word. Never widen a mark to swallow text that is word-for-word
identical on both sides — every borrowed word is a word the reader reads twice to find out
nothing happened to it.

**P2 — Never make the reader diff the redline.**
If a struck span and the inserted span beside it share most of their words, the tool has
handed its own job back to the reader. Two honest renderings exist — tighten to the real
differences, or swap the whole paragraph cleanly. The middle ground is the one thing that
must never ship.

**P3 — Unchanged text stays contiguous.**
A mark must not interrupt prose that did not change. When an inserted block repeats a word
that already sits on the far side of the mark, the mark is in the wrong place and should
slide until the untouched text runs unbroken.

**P4 — A mark is a syntactic unit.**
A span should open and close where a lawyer would put a bracket: at a phrase, a clause, a
parenthetical, a sentence. A span that begins mid-construction, or trails off on *before*,
*and*, *of*, reads as machine residue and costs the reader a second pass.

**P5 — Atomic values are indivisible.**
Amounts, dates, periods, percentages, clause numbers and defined terms change whole or not
at all. `S$5,000,000 → S$7,500,000` is one mark; a digit-level interleave of the same edit
is unreadable and, worse, misquotable.

**P6 — Punctuation is not dragged into a borrowed mark.**
A sentence-final full stop that exists on both sides stays outside the change. Punctuation
carries meaning in drafting, so marking it falsely is not cosmetic — it makes the reader
check whether the sentence boundary moved.

**P7 — A move reads as a move.**
Text that has relocated verbatim is marked as a move at both ends, at paragraph scale and at
phrase scale. Rendered as a deletion here plus an insertion there, the reader has to prove
the two blocks are identical before they can dismiss them.

**P8 — Renumbering is a consequence, not an amendment.**
Inserting a limb pushes every label below it. The redline shows one inserted limb; the
relabelling is absorbed or foldable. A cascade presented as substantive edits buries the
single change that caused it.

**P9 — Never lose the clause number.**
The number is the reader's only navigation handle and the thing they quote in the email
back. It has to survive on the paragraphs that changed — those are the ones being discussed.

**P10 — Classify to triage, never to hide.**
Every change is always present and always accept/rejectable. Classification only decides
what a reviewer may fold away on a first pass — cross-references, renumbering, formatting —
so the substantive edits stand alone. Nothing is ever removed from the redline. (This is
ADR-0008 / the project principles restated as a reading principle: it constrains the others, and
no readability gain may be bought by suppressing a change.)

---

## How the principles are measured

`tests/fixtures/readability-metrics.ts` renders each corpus case through the real path
(`extractTextRecords → runComparePipeline → buildRenderModel`) so what is measured is the
run/revision structure the paginated preview and the vector PDF actually paint. Four
numbers come out:

| Metric | Principle | Meaning |
|---|---|---|
| `redundancy` | P1, P2, P3 | 0–1. For each struck span sitting beside an inserted span, the share of the shorter side that is word-for-word identical on both sides. **The primary objective**: it ranked the 22 cases in exactly the order the human reviewer graded them. A paragraph rendered as a whole-for-whole swap is exempt — P2 admits that as the other honest rendering. |
| `marks` | P1, P4, P7 | Count of maximal non-equal segments. Not smaller-is-better: three honest deletions beat one confusing block. Pinned exactly, so a change in rendering shape must be acknowledged. |
| `atomicIntact` | P5 | No segment boundary falls inside an amount, date, period, percentage or clause reference. An invariant, never a tuning target. |
| `numbered` | P9 | Share of the revised text's manual clause numbers still visible. An invariant. |

A reader's idea of a word is the yardstick for `redundancy`, deliberately **not** the
production tokenizer's — the tokenizer's punctuation gluing is one of the things under test.
Word-internal apostrophes and hyphens stay inside the word, because `Purchaser's` and
`Third-Party` are one word each to a reader, so swapping them whole borrows nothing.

P7 (move) and P8 (renumbering) and P10 (triage) are not captured by these four numbers;
they are asserted case-by-case in the existing move-detection, cascade and classification
suites, and by `marks` on the move cases.
