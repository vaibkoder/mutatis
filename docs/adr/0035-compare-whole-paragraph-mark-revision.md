# ADR-0035: Compare — a whole-paragraph change must revise the paragraph mark

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0032 (paragraph merge/split, which already marks the mark at a seam); feature 08 (`spec/features/08-compare.md` §5.5)
- **Workstream:** Tier-3 fix on the redline-readability board, approved by the reviewer 2026-09-04

## Context

`authorBody` emitted a whole-paragraph deletion by striking the runs inside the
paragraph and nothing else:

```xml
<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
  <w:del w:id="1" …><w:r><w:delText>Removed clause</w:delText></w:r></w:del>
</w:p>
```

Word removes a paragraph only when the paragraph **mark** carries the revision.
With the mark unmarked, accepting that deletion strips the text and leaves the
paragraph: **an empty paragraph where the clause was**, one per deleted clause.
A reviewer who accepts a redline of six struck clauses gets six blank
paragraphs to clean up by hand, in a document they were told was accepted.

Two things made this survive to now:

- ADR-0032 marks the paragraph mark at a **merge/split seam**, so the codebase
  already had `markParagraphMark` and the `CT_PPr` placement rules. The plain
  whole-paragraph case simply never called it.
- **No test covered it.** The behaviour change below left the full suite green
  at 1749 tests — proof that nothing asserted the paragraph-mark shape on a
  whole-paragraph delete, not evidence that the change was safe.

### The mirror defect is live too

`emitInsert` wrapped only the runs as well. Its failure mode is the reverse and
equally real: **rejecting** a whole-paragraph insertion strips the added text and
leaves the paragraph, so a reviewer who rejects the other side's redline gets
their own document back plus one blank paragraph per clause the other side
added. Fixing one direction and not the other would leave accept-all and
reject-all disagreeing about which paragraphs exist, which is worse than the
symmetric bug: the two paths would need separate mental models.

## Decision

**Revise the paragraph mark alongside the runs for every whole-paragraph
insertion and deletion** — `<w:pPr><w:rPr><w:del …/></w:rPr></w:pPr>` on a
delete, `…<w:ins …/>…` on an insert — placed by the existing
`markParagraphMark`, i.e. by the same `CT_PPr` rules ADR-0032 established (one
`<w:rPr>`, ordered `rPr → sectPr → pPrChange`, revision first inside
`CT_ParaRPr`, leading `w:pPr` matched by depth-counting).

**One guard, shared by both directions** (`paragraphMarkIsRevisable`).
the project principles: two implementations of the same rule is a defect, and here it
would also be a behavioural one — accept-all and reject-all must agree about
which paragraphs survive. Two shapes keep their mark unrevised, because revising
it destroys something the change was never asked to touch:

- **a mark carrying `w:sectPr`.** The last paragraph of a section holds that
  section's page setup — size, orientation, margins, headers. Removing the mark
  merges this section into the next, so resolving an ordinary clause change
  would silently re-paginate the document.
- **the last paragraph of its part** (`partId`, so a table cell is judged
  separately from the body). There is nothing to merge into: a body must end
  with a paragraph and a table cell must contain at least one, so removing the
  final mark of either yields a package Word may refuse.

For those two the pre-ADR behaviour stands and an empty paragraph is left
behind. That is a smaller defect than a re-paginated document or a file Word
declines to open, and it is bounded: at most one paragraph per part, in the
place a reader is least likely to read as a mistake.

## Consequences

- **A whole-paragraph change now emits two `w:id`s**, one for the runs and one
  for the mark, where the mark is revisable. Ids stay unique and monotonically
  *assigned*; they are **not** ascending in document order, because a mark's
  revision serialises inside `w:pPr`, ahead of the runs whose id was minted
  first. Two tests that counted `<w:ins>`/`w:id` occurrences were measuring
  elements while claiming to measure paragraphs, and are corrected to measure
  paragraphs — `countInsParagraphs` in the text-vs-text acceptance suite was
  the one that caught this.
- **Zero effect on the rendered redline.** RenderCore feeds the paragraph-mark
  `rPr` into its style cascade, so `w:del`/`w:ins` are new nodes there;
  `applyRPr` is a whitelist `switch` with `default: break`, so it ignores them.
  Verified by construction, not by inspection: `buildRenderModel` output is
  byte-identical (`JSON.stringify` equality) for the same body with and without
  the mark revision.
- **`classByRevId` capture is unaffected in kind.** The mark's id is minted
  inside the same op, so under `captureClasses` it is tagged with that op's
  classes like any other revision. Capture remains observational — the authored
  XML is byte-identical whether or not it is on.
- **No suppression** (ADR-0008). The mark revision adds a revision; it removes
  and hides nothing. Both sides still reconstruct exactly.

## Evidence

Five ablations, each failing only its own tests (`compare.output.author.test.ts`,
25 tests):

| ablation | fails |
| --- | --- |
| don't strike the deleted mark (pre-ADR behaviour) | 4 |
| drop the `sectPr` guard | 1 — exactly the `sectPr` test |
| drop the last-paragraph-of-part guard | 5, incl. the table-cell/`document` split and a pre-existing insert test |
| don't mark the inserted mark (pre-ADR behaviour) | 3 |
| — restored — | 0 (25 pass) |

The insert-direction tests were written **before** the fix and observed failing;
the last-of-part insert test passed vacuously at that point (nothing was marked)
and is confirmed discriminating by the third ablation above.

Well-formedness is asserted with `XMLValidator`, not by eye, on a paragraph that
already carries both `pPr` properties and a paragraph-mark `rPr` — the shape
`CT_PPr`'s "one `rPr`, ordered first" rule makes delicate.

Full suite: 1749 → **1758 tests + 1 todo**, green; lint and `tsc --noEmit` clean.
Readability corpus unmoved (its ratchets are asserted in the suite).
