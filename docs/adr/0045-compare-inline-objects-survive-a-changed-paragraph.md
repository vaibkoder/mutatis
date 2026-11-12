# ADR-0045: Compare — what is not text survives a changed paragraph, and an equation is compared

- **Status:** Accepted; amended 2026-10-18 (§ Amendment 1 — Accepted limits 1 closed); Accepted limits 3 and 5 narrowed 2026-10-26 
- **Date:** 2026-10-17 (approved 2026-10-18)
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0028 (numbering prefix restored on the replace path); ADR-0032 (merge/split — not covered here); ADR-0035 (paragraph-mark revisions); ADR-0038 (the readability instrument that measured the first draft); ADR-0042 (text identity anchors, non-text identity annotates); ADR-0043 (insertion before deletion); feature 08 (`spec/features/08-compare.md` §5.0, §5.5)

## Context

A probe on 2026-10-17 (branch `feature/factual-consistency` at `57d38fe`, minimal
synthetic packages through `runPipeline`) found three silent losses in the
single-pair Word-vs-Word path.

1. **An equation was invisible.** Paragraph text was the concatenation of
   `<w:t>` only. An Office Math (OMML) equation keeps its text in `<m:t>`, so a
   change to a formula and nothing else gave `summary.total` **0**, ops
   `['equal', 'equal']`, and the revised formula in `authoredBodyXml` with no
   tracked markup. That is an ADR-0008 suppression: a change the reader can see
   in Word, absent from the redline and the summary.
2. **A changed paragraph lost its structure.** `emitReplace` rebuilds a
   changed paragraph from its Level-2 word diff, and a word diff is text. Every
   footnote reference in the paragraph disappeared, a hyperlink became plain
   words, and a complex REF field became its frozen result — while the
   untouched paragraphs around it, emitted from `rawXml`, kept all of theirs.
   The same rebuild dropped bookmarks, so a REF field elsewhere pointing at an
   edited clause lost its target.
3. **A text box was poured into the sentence.** Text-box text has always been
   part of its host paragraph's compared text (extraction recursed into
   `w:txbxContent`), glued to the last word before it (`"the Price.Boxed
   note"`). On the replace path the box itself went with the rebuild. And Word
   normally writes a text box as `mc:AlternateContent` with a VML
   `mc:Fallback` copy, so every word in the box was read twice.

Two neighbouring findings in the same family were measured and are **not**
fixed here (see Accepted limits 1): a paragraph whose only change is a lost
footnote reference, or a lost inline image, still reports `summary.total` 0.
*(Closed on 2026-10-18 by § Amendment 1, at the maintainer's request.)*

## Decision

### 1. One home for paragraph text, and it reads equations

`preprocess/inline-units.ts → collectText` is now the single extractor;
`extractParaRecords` calls it (`paragraphText`). It reads `<w:t>` **and
`<m:t>`**; it reads an `m:oMath`, a `w:txbxContent` and each paragraph nested in
one as its own words (a space either side, folded by `normaliseText`) — unless
it holds no text, when it reads as nothing, so the empty shell a revision leaves
behind cannot split two words that were never apart; and it skips
`mc:Fallback`, reading the `mc:Choice` once. `<w:instrText>` is still not
read — a field is compared by its result, the text the reader sees.

### 2. The replace author puts objects back around the rebuilt text

`inline-units.ts → paragraphUnits` cuts a paragraph's inline content into
units: **text** (ordinary runs — rebuilt from the word diff, exactly as before),
**object** (a hyperlink, a complex field from `begin` to its matching `end`, a
simple field, a content control, an equation, a run carrying a footnote
reference, a drawing or a text box, anything else that is not an ordinary run),
and **marker** (a bookmark). `output/inline-objects.ts → authorObjectAwareRuns`
maps every word-diff character back to the unit it came from, through
`normaliseTextWithOffsets`, and authors each object by one rule:

- **Its text is untouched, it lines up with one object of the same sort on
  the other side** (each side's `eq` characters meet the other and nothing
  else — a space at the object's edge excepted, since `[the portal ]now` →
  `[the portal] now` changes nothing visible), **the two hold the same
  objects inside** (a census: a link that gained a footnote has the same words
  and is not unchanged), **and neither holds a space that is not its own** —
  one collapsed into a neighbour's space, or trimmed at the paragraph's edge,
  which its XML would carry into a view where nothing collapses or trims it
  (`["Portal" ][ Terms]` beside `["Portal"][Terms]`) → re-emitted **verbatim**
  from the revised paragraph, unmarked.
- **It lines up, its text changed, it is a container (a link, a field, a
  content control, a smart tag), the part to be rebuilt holds nothing but
  ordinary runs on BOTH sides, and its characters form one unbroken stretch of
  the diff** → the revised object's **shell is kept** — its tags, a field's
  `begin…separate` and `end` — and the word diff runs *inside* it, the inner
  runs wearing the first text run's properties so a link still looks like a
  link. A container holding anything else (a footnote inside a link, a
  PAGEREF inside a TOC entry's link, an equation inside a content control) is
  replaced whole instead, because rebuilding its words would drop that content;
  so is a pair a plain-text edit falls *between*, which would otherwise close
  and reopen the shell and emit the link or field twice.
- **Anything else** → **whole**: the revised object as one tracked insertion,
  the original as one tracked deletion, insertion first (ADR-0043). Every
  character of both sides is still on the page, inside one of the two objects.
- **Zero-width objects** (a footnote reference, an image) have no text to line
  up by. When both sides carry the same sequence of sorts they are re-emitted
  from the revised side; otherwise every one of them is marked, inserted or
  deleted — over-marking, never under-marking. So are they when any of them
  holds a space after all (a link of only a space, its space collapsed or
  trimmed away): the revised one would carry that space to wherever it sits.
  And one pair is kept unmarked only where a single point of the diff is its
  place on BOTH sides — after the characters before it on each, before the
  characters after. A footnote that moved has two places, and so does an
  unmoved one whose neighbouring words were replaced on both sides (its
  original place falls inside the struck words, its revised place inside the
  inserted ones); kept there, rejecting every change would misplace it. Such a
  pair is marked: the revised object inserted where it is, the original
  struck where it was — in the original's order beside any kept one in the
  same place, or rejecting would swap two objects that never moved.
- **Bookmarks** are re-emitted from the revised side where they sit, never
  revised.

A whole-object deletion that the walk reaches before the insertion beside it
(a link that became plain text) is **held** until that insertion is emitted, and
its `w:id`s are minted only then. A wrapper's id is minted before the ids of a
text box's own paragraphs inside it. The walk also makes revisions stage 6.5
never saw — a stretch plain on one side and inside a changed object on the
other, a split space — so a deletion can still open a region; at the end, in
every run of consecutive inserted and deleted pieces with nothing unchanged
between, the deletions before the region's last insertion are moved after it
and the ids re-dealt in the new order. A bookmark does not end the run and
keeps its place among the insertions, so the accepted view is untouched — the
bookmark still opens where it did — and only struck text moves past it in the
rejected view, where bookmarks are placed from the revised side anyway. A
kept footnote reference or image DOES end the run (limit 13). Each view reads
the same text either way. So within a paragraph's runs, and inside each
container, revisions read insertion first, except across a kept zero-width
object, and ids open in document order; the sweep checks the order
structurally on every case, since a string test for `</w:del><w:ins ` cannot
see past a bookmark. (The paragraph-mark
revision of an inserted or deleted paragraph keeps its pre-existing id, minted
after its runs.)

A collapsed run of whitespace is credited to a REAL space where it has one,
not to the boundary space read around an equation or text box: credited to the
boundary, the space would be authored by nobody, since the object is emitted
from its own XML (`"r=3%per annum"`). A space that is ONLY a boundary belongs
to no unit at all: the document has no character there, so it is nobody's
edit — credited to the equation, deleting the word beside an unchanged formula
struck and re-inserted the formula. Where a kept object's own edge space lines
up with such a boundary on the other side, the verbatim counterpart lacks it,
so it is struck as plain text beside the object. Anything zero-width at a
container's start is emitted before the container's shell is opened, whichever
side's character opens it. The one-stretch condition for a diff inside a
container is checked until nothing more breaks: replacing one pair whole
re-homes a character it shared with the next pair, which can break that one.

An edge space may line up with a character outside its object (`[the portal
]now` → `[the portal] now`), and the kept object's own space then stands
where the other side's did — but only when that character is next to the
object in the stream. Lined up with a character of another object that does
not correspond, or across a deletion or insertion, the object's XML would
carry the space to its own edge: shown twice in one view, or moved across the
words between. Such a space is split into a deletion and an insertion before
anything is authored, and the objects are paired again, until none is left.

### 3. Every result checks itself, and falls back to exactly the old output

Before it is used, the authored paragraph is flattened both ways
(`flattenTrackedChanges`) and compared with the source paragraphs in two
respects: the **text as the diff reads it** — `normText`, boundary spaces and
all, numbering prefix stripped — and a **census** of
the objects text cannot see — footnote and endnote references, field begins,
links, simple fields, equations, drawings, text boxes, content controls, and
(accept-all only) bookmarks — after dropping any container a whole-object
deletion has emptied. The text is compared in the diff's own model on purpose:
a check stricter than the diff it checks (the document's own spelling was the
third draft) rejects what the diff calls equal — a space added beside an
equation — and every rejection is a paragraph sent back to the lossy rebuild
(limit 12). A space LOST by this author is prevented where it would arise, and
the tests, which do compare the document's spelling, pin it. The check runs on a **scratch id counter**, so a result
that fails mints no id. A pair with **no object or marker on either
side** never entered the module, so an ordinary edited paragraph was authored
**byte for byte** as before. *Amended 2026-10-26 :* every one-to-one
replace now enters it, rebuilt from its source runs first and held to the
reconstruction check (characters and run formatting, `output/reconstruction.ts`);
a plain pair that fails it takes the text rebuild, as before. A paragraph of
plain, unformatted runs still comes out byte for byte as the text rebuild wrote
it (pinned: `pipeline.inline-objects.test.ts`, *an edited paragraph of plain runs
is authored byte for byte as before*). Every other way out — a field that opens or
closes in another paragraph, an object inside a stripped numbering prefix, a
word diff that does not spell both sides, the mapped normaliser disagreeing
with `normaliseText`, the two-way check failing — falls back to the text
rebuild, which is what every paragraph got before this ADR.

### 4. Whole-paragraph insertions and deletions revise inside containers

`emitInsert`/`emitDelete` wrap a paragraph's runs in one `<w:ins>`/`<w:del>`.
`w:hyperlink` and `w:fldSimple` are paragraph content (`EG_PContent`) that
`CT_RunTrackChange` may not contain, and Word records an equation's revision on
its math runs. A paragraph carrying a link, a simple field, an equation or a
text box (`needsStructuredRevision`) is now revised **inside** those
(`reviseParagraphRuns`), the box's own paragraphs included; every other
paragraph keeps the single wrapper it always had. This is reachable from item
1: an equation edit that ADR-0026 demotes is authored on this path.

### 5. A deleted field's instruction is deleted

`toDelText` converts `<w:instrText>` to `<w:delInstrText>` as well as `<w:t>`
to `<w:delText>` (ECMA-376 §17.16.13). Before this, a whole-paragraph deletion
carrying a field struck its result and left its instruction live inside the
`<w:del>`.

## Accepted limits

1. **A change to a zero-text object alone — CLOSED by § Amendment 1
   (2026-10-18).** It was: a paragraph whose only change is a removed footnote
   reference or inline image stayed `equal` and reported `summary.total` 0. It
   is now a reported, marked `property` op. What is still not reported, by
   design or measurement, is listed under the amendment: a picture swapped for
   another in the same place, a link or field put around words that did not
   change (an object WITH text, not zero-text), comments; and the change-set
   and AI exports, which omit every `property` op.
2. **A changed text box or equation is replaced whole.** The revised box (or
   equation) is inserted and the original struck, with the box's own words
   marked inside each; the word diff is not run inside a box or inside math
   structure. Every word is shown; Word renders two boxes until the change is
   resolved.
3. **Merge/split (ADR-0032) with an object still rebuilds from text.** The
   object-aware author handles one-to-one replaces; a `replace` spanning
   several records whose records hold an object keeps the pre-ADR behaviour
   and loses its objects. *Narrowed 2026-10-26 :* a range with no
   object is rebuilt from its source runs (spec 08 §5.5). Moves (`moveFrom`/
   `moveTo`) still wrap a paragraph's whole content, links included, as before.
4. **A non-text change to an object whose text is equal is not reported.** A
   link whose target changed, or a REF field now pointing at another bookmark
   with the same result text, is emitted from the revised side, unmarked.
   This is ADR-0042's rule applied inside a paragraph: text identity anchors.
   Comparing the object's XML instead was rejected below.
5. **Ordinary run formatting in a changed paragraph is lost only on the
   fallback.** *Narrowed 2026-10-26 :* every one-to-one replace is
   first rebuilt from its source runs — kept and inserted characters from the
   revised runs, struck ones from the original runs, each run's `rPr` and the
   documents' own characters kept — and that rebuild is used only when it
   reconstructs both documents exactly, characters and run formatting
   (`output/reconstruction.ts`). A paragraph that fails the check, a shape the
   unit model does not cover, and a merge/split with an object (limit 3) take
   the walk below and then the text rebuild, where text units are rebuilt as
   plain runs and text inside a kept container inherits that container's first
   run properties, as before. Measured on two private real-world pairs: no paragraph
   loses run formatting either way. Until then it was:
   *ordinary run formatting in a changed paragraph is still lost, as it always
   was.*
6. **Not opened in Word.** The shapes authored here — `<w:ins>`/`<w:del>`
   inside `m:oMath` around `m:r`, revisions inside `w:hyperlink` and
   `w:fldSimple`, revisions on a text box's own paragraphs, `w:delInstrText` —
   are what my reading of ECMA-376 permits and what Word is understood to write,
   but no package from this change has been round-tripped through Word desktop
   from this repo. That belongs with the standing live-host smoke tests.
7. **A failed check costs the whole paragraph's objects, and the check is a
   backstop, not a proof.** When the two-way check fails, the paragraph falls
   back to the pre-ADR text rebuild and EVERY object in it is lost — the
   losses this ADR fixes, confined to that one paragraph and never a hidden
   text change. Twice this limit said no input trips the check, from the
   cases its tests happened to construct, and twice a review pass falsified
   it. It is now MEASURED over a generated input space instead: a seeded sweep
   of paragraphs mixing text with every object kind here, with words edited,
   edge spaces moved and objects added, dropped and swapped. Then a review
   pass falsified the measurement too: 0 in 210,000 held for the generator's
   vocabulary, and drafting punctuation — a quoted defined term at a link's
   end — makes the diff split an edge space differently, which tripped the
   check about once in 5,000 replaces. The vocabulary now carries that
   punctuation, and over 410,000 generated pairs the final code trips the
   check on none, and misplaces no footnote reference or drawing in either
   view (checked since review pass 8). That bounds the generator's space, not real documents:
   shapes it does not make — containers nested more than one deep,
   revisions already present, fields other than REF, manual numbers (limit
   11) — are not measured, and the last omission found was in the
   vocabulary, which no list of shapes names. 600 pairs run in the suite.
8. **Tables.** The Markdown table-unit path (`ParaRecord.table`,
   `output/author-table.ts`, ADR-0027) is untouched and remains behind 13b's
   live-host gate. Word-vs-Word table-cell paragraphs are ordinary records and
   take this path like any other paragraph.
9. **Two non-text changes beside a text edit resolve to the revised side,
   unmarked.** (a) Zero-width objects with the same sequence of sorts on both
   sides, each at one place on both, are taken from the revised side — an
   inline image swapped for another of the same shape (same kind, same
   content census) in the same place is shown as the new one, and, the case
   that matters most, two footnote references that swapped their footnotes
   (`Seller¹ shall³` → `Buyer³ shall¹`) are taken from the revised side
   unmarked, so the note under a word changes and the redline does not show
   it; a shape swapped
   for a picture is a different sort and is marked, and so, since review pass
   8, is a footnote that moved or whose neighbouring words were replaced on
   both sides (§2). (b) A container diffed
   inside keeps the REVISED shell, so a link target or REF instruction that
   changed along with its text is not marked. Reject-all does not restore the
   original in either case. Same root as limit 4.
10. **An accepted whole-object deletion leaves an empty shell** where the
    revision had to go inside the object — an `m:oMath` with no runs, a
    `w:hyperlink` or content control with no content. Invisible in print, and
    the census ignores it; not removed.
11. **The fallbacks are the old, lossy rebuild.** A field that opens in one
    paragraph and closes in another (a multi-paragraph TOC) sends that
    paragraph back to the text rebuild, structure and all lost — as before
    this ADR. So does any object holding part of a stripped manual number:
    not only a field producing `4.2`, but a link, field or content control
    that OPENS the paragraph with text the number-stripper reads as a number
    (`4.2 `, `(a) `) — measured at about 6% of replaces when a sweep's
    vocabulary carries such tokens. An equation straight after a number is
    not inside it (the separator is only its boundary) and is kept, the
    number restored without a space the document does not have. Both
    fallbacks pinned (*"FALLS BACK"*). **Maintainer decision, 2026-10-18:
    accepted as is** — a clause number made from a field (or held in a link
    or content control) may become plain text in an edited clause: *"I'm fine
    with clause numbers being made from fields becoming plain text."* The
    text and the change are still shown; only the field structure of the
    number is lost.
12. **A space added or removed directly beside an equation or text box is not
    reported.** The boundary space the diff reads there absorbs it, as §5.2's
    whitespace collapse absorbs a doubled space anywhere else: `x= 1` and
    `x=1` beside a formula are the same text to Compare. The authored redline
    keeps whichever side's spacing the object-aware walk emits; the self-check
    compares in the same model and accepts it (§3).
13. **A deletion the walk makes before a kept footnote reference or image,
    with its insertion after it, reads deletion first** — against ADR-0043.
    `Pay promptly ¹[the Seller]` → `Pay ¹the Seller`: the word diff strikes
    "promptly", the walk inserts the link's words as plain text after the
    footnote. Insertion first would carry "promptly" past the footnote, and
    rejecting every change would then read `Pay ¹promptly the Seller` — a
    footnote that never moved, misplaced in the document a reviewer gets
    back. Both views being right and insertion first cannot both hold here;
    the views win. A bookmark does not have this problem (it is placed from
    the revised side in both views) and is reordered past. Pinned
    (*"RECORDS ITS LIMIT"*).
14. **Reject-all of an unchanged stretch the two documents spell differently
    reads in the revised spelling (measured 2026-10-26; accepted by
    the maintainer).** Where §5.2's fold makes two spellings equal — whitespace
    (a tab, a break, a double space), a non-breaking space, a curly against a
    straight quote or apostrophe, an en or em dash against a hyphen — the word
    diff calls the stretch `eq`, and one authored character cannot carry two
    spellings: the rebuild from source writes the revised one. The
    reconstruction check excuses exactly those characters, and only inside
    unchanged stretches: every excused character the view shows is unrevised
    (outside every insertion and deletion) and proved right by accept-all, and
    one the view lacks OUTRIGHT (a gap where the view has no character at all)
    must have every neighbour so — never beside a struck or inserted character,
    where it may be a deleted one lost. A substitution (a curly quote the view
    shows straight) is excused on the shown character's own proof, beside a
    revision or not. Any other character
    (a non-breaking hyphen included) is a loss and sends the paragraph to the
    fallback. The count takes each side of a substitution (a curly apostrophe
    written straight is 2). Measured on two private real-world pairs: reject-all differs
    in a handful of paragraphs and characters (mostly whitespace), accept-all in none. The
    yardstick prints the count and pins it. Marking these differences instead
    would report changes §5.2 folds away on purpose (corpus C30, *"nobody wants
    a redline that marks every double space"*; the corpus notes that conventional
    tools behave the same, which is not measured here) — a §5.2/§6 decision for the
    maintainer, declined for now.

## Rationale

- **Never-suppress is the governing constraint**, and it decided item 1: an
  equation the reader sees and Compare cannot is the shape ADR-0008 forbids.
- **The fix is additive where it can be.** Equal paragraphs still emit `rawXml`.
  Edited paragraphs with no objects are byte-identical. `normText` changes only
  for paragraphs carrying an equation or a text box, and identically on both
  sides, so alignment between two unchanged copies cannot move.
- **The first review pass (spec-conformance review, 2026-10-17) returned FAIL**
  with three blocking defects, each proven and each passing both the draft's
  own check and its tests: a footnote or PAGEREF nested in an edited link, or
  an equation in an edited content control, was rebuilt away; the real space
  after an unchanged equation or text box was dropped; and a plain-text edit
  inside a link's stretch emitted the link (or a whole REF field) twice. Plus
  text-box ids out of document order and limit 9, unrecorded. All are fixed
  or recorded above, each fix has a test that fails without it, and the test
  round-trip now checks the document's own spelling and the object census
  (the module's own check did too, until pass 3 moved its text to `normText`).
- **The second review pass returned FAIL** on one blocking gap: three
  simple inputs tripped the self-check (limit 7 as first written said none
  could), and the fallback then dropped every object in the paragraph — a
  footnote added inside a text-equal link was in neither view of the redline.
  Two low findings came with it: a footnote or bookmark just before a link was
  carried inside it when the link's diff opened with a deletion, and a space
  moving across a link's edge struck and re-inserted an unchanged link. All
  three fixed, each fix ablated against a test that fails without it.
- **The third review pass returned FAIL** on the same class: whitespace-only
  edits beside an equation or text box still tripped the check, which by then
  compared the document's own spelling — stricter than the diff, so it
  rejected what the diff called equal. And a DrawingML shape and an inline
  picture counted as one sort. The check now reads `normText` (§3, limit 12)
  and a sort includes its content census. Since three passes had each found
  inputs the previous one's tests could not express, the claim was then
  measured rather than re-asserted (limit 7): the sweep found four more
  shapes, all in §2 — a boundary-only space credited to an object, an
  emptied equation read as a word break, a kept object's edge space with no
  counterpart, and the one-stretch condition checked once instead of to a
  fixpoint. Each has a test from the sweep's own case that fails without it.
- **The fourth review pass returned FAIL** on one blocking shape, found by
  sweeping with a wider vocabulary than the suite's: two adjacent links or
  REF fields, the first ending in a quoted term, with a space added between
  them alongside an edit. The space collapsed into the first object's; the
  second, unchanged, was emitted verbatim with its own copy, so reject-all
  read a space the original never had, and both objects were lost to the
  fallback. Verbatim now also requires each object to own every space in its
  XML (§2). Widening the suite's vocabulary to measure the fix found three
  more of the same family at once — a space trimmed at the paragraph's edge,
  a space lined up across two objects that do not correspond, and one lined
  up across a deletion — and the general rule in §2 (split a space no object
  can author) replaced a narrower one written for the third, which then
  ablated to nothing over 150,000 pairs and was removed. The pass also found limit 11 understated, now restated, and
  the equation-after-a-number case fixed. Each has a test that fails without
  it, and the suite's sweep vocabulary now carries the punctuation.
- **The fifth review pass returned FAIL (minor)**, with no fallback left in its
  punctuation-heavy sweep. It found §2's insertion-first claim false where the
  walk revises on its own (86 of 5,604 replaced paragraphs) — the order a
  redline, its preview and its side-by-side panel must agree on (ADR-0043) —
  and a link holding only a space still lost to the fallback. Both fixed
  (above), each ablated against its test; the sweep now asserts the order and
  generates space-only links.
- **The sixth review pass returned FAIL (minor)** on the same claim: a
  bookmark or kept footnote reference between the walk's deletion and its
  insertion ended the reordered run, and the sweep's string check could not
  see it (71 of 5,467 replaced paragraphs in the reviewer's sweep). Zero-width
  pieces were made transparent to the reorder, and the sweep's check
  structural.
- **The seventh review pass returned FAIL (minor)** on that fix: a kept
  footnote reference is zero-width too, and moving struck text past it put a
  footnote that never moved in the wrong place once every change was rejected.
  Only bookmarks are transparent now; for a kept footnote or image the order
  and the original's placement cannot both hold, and the placement wins
  (limit 13).
- **The eighth review pass returned FAIL (minor-to-medium)**, the same harm
  through a shape no reorder could reach: words replaced on both sides of a
  footnote that did not move put its original place inside the struck text,
  and the footnote kept from the revised side was misplaced once every change
  was rejected. The keep-the-revised-one rule dated from the first draft.
  A pair is now kept only at a place common to both sides (§2), which also
  retires limit 9(a)'s moved footnote. The census could not see any of this —
  every count was right — so the sweep now also checks where each footnote
  reference and drawing sits among the characters, in both views.
- **The ninth review pass returned PASS**, with one wording note taken into
  limit 9(a): footnote retargeting is named there.
- **The real-document instrument chose between two drafts** (ADR-0038). The
  first draft replaced every changed object whole and raised a private
  pair's `charOffendingBlocks` ceiling from 22 to 23: 24 renders of a
  four-character field result (a renumbered cross-reference, by its length)
  struck and re-inserted with most of its characters in common. Diffing inside
  a kept field or link brought it back to 22 with `wholeParaChurn` and
  `alignmentRate` unchanged. On both private pairs every paragraph with an object
  took the new path and passed its check; none fell back — measured again after the review fixes, under the stricter check.

## Rejected alternatives

1. **Carry each object in `normText` as a sentinel character** (the
   `PARA_SEAM` device). It would make zero-text objects comparable (limit 1),
   but a sentinel in `normText` leaks into every display consumer — the
   side-by-side panel, the change-list preview, the exports, the AI export — and
   a link's or field's own words must stay comparable as words. The seam can be
   spelled out after authoring because it lives only in `wordDiffs`; a sentinel
   in the record could not be.
2. **Demote any changed paragraph with an object to whole-paragraph delete +
   insert.** It keeps structure trivially, but strikes and re-inserts every
   word of an edited clause because it happens to carry a footnote — the churn
   ADR-0026/0041 and the readability plan exist to remove.
3. **Replace every changed object whole** — the first draft. Measured worse
   (above): a renumbered REF result struck and re-inserted rather than diffed.
4. **Call an object unchanged only when its XML is unchanged.** A REF field's
   bookmark id and a link's relationship id are document-local, and Word
   reassigns them freely (ADR-0042 measured the same of `numId`), so every
   cross-reference in an edited clause would be struck and re-inserted.
5. **Map the word diff onto every original run**, keeping all run formatting.
   That is the general fix behind limit 5, but it changes the authored bytes of
   every edited paragraph in every document and is not what this probe found;
   it deserves its own decision. *Taken on 2026-10-26*, through this
   module's own map and held to a reconstruction check (limit 5).
6. **Take text-box text out of the host paragraph.** Text boxes are not
   records of their own, so their words would then be compared nowhere — a
   change inside a box would be suppressed outright.

## Consequences

- `normText` (and so the hash) of a paragraph with an equation now contains
  the equation's text; of one with a text box, the box's text once, as its own
  words. Both sides change alike. Batch pairing sketches hash these paragraphs
  differently from before, again alike on both sides.
- `toDelText` changes the output of whole-paragraph deletions that carry a
  field; `emitInsert`/`emitDelete` change for paragraphs carrying a link, a
  simple field, an equation or a text box. Nothing else they author moves.
- A changed paragraph with objects may carry more revision elements than
  before (one per kept container it diffs inside), each with its own `w:id`,
  minted in document order.
- Tests: `tests/unit/compare/pipeline.inline-objects.test.ts` (each shape,
  each checked accept-all and reject-all, plus the byte-identity pin and the
  recorded limits), `tests/unit/compare/pipeline.inline-objects.sweep.test.ts`
  (the seeded sweep of limit 7, 600 pairs per run: no object lost or doubled,
  both views reading as their source, no revision wrapping a link or simple
  field, ids in order) and `tests/unit/compare/normalise-offsets.test.ts` (the
  mapped normaliser held to `normaliseText` on 5,000 random strings). Fixtures
  are synthetic (`tests/fixtures/inline-objects-docx.ts`).

## Amendment 1 (2026-10-18) — a footnote or picture added, removed or moved alone is reported

On 2026-10-18 the maintainer asked that the removal of footnotes or pictures be
reported. Under ADR-0008 a change found and not reported is suppression, and
this one was not even found.

### Decision

- **Detected where ADR-0042 detects a non-text change**, stage 2.75
  (`diff/property.ts`): text identity anchors, non-text identity annotates. A
  one-to-one `equal` pair whose ordered list of **zero-text objects** differs
  becomes a `property` op with `PropertyChange.objects`. A zero-text object is a
  footnote or endnote reference, a picture (`w:drawing`, or legacy `w:pict`;
  one per `mc:AlternateContent`, whose fallback is skipped as `collectText`
  skips it), a text box, or an embedded object (`w:object`) — each as a SORT and
  a PLACE (`preprocess/inline-units.ts → zeroTextObjects`). The place is counted
  in the normalised, prefix-stripped text, so a whitespace run written
  differently cannot move it; an object standing in a manual number is coded
  `-2` (before the number) or `-1` (inside it), so renumbering `9.` → `10.` does
  not move it and moving it across the number does. An object inside an
  equation or text box also carries `within`, the innermost one holding it: a
  collapsed space can give "just before the equation" and "just inside its
  start" the same count, and they are different places. Places are found with a
  mark pushed into the text as it is read, because the boundary around an
  equation or box is spliced in after its content. **Never which footnote or
  image**: `w:id` and `r:embed` are document-local, Word reassigns them, and one
  footnote inserted early renumbers every reference after it (pinned: a
  renumbered id in the same place is `equal`). A table record is never read
  (ADR-0027). A regex pre-filter keeps every paragraph without such a tag from
  being parsed.
- **No sentinel in `normText`.** Rejected alternative 1 stands; the objects
  ride on the op, not in the text.
- **`substantiveText`, never `formatting`.** A footnote reference is a run, so
  this is a run inserted or deleted (§6 row 6); §9.8 already makes a footnote
  change substantive. Classed `formatting` it would be hidden by *Hide
  formatting-only*, on by default — ADR-0042's recorded trap. And never a
  cascade: a paragraph whose number moved AND whose footnote went changed in
  its own right, and a cascade member is hidden by *Collapse renumbering*.
- **Marked — the one `property` op that does not author as `equal`.** Three
  tries, most faithful first (`output/author.ts → emitObjectChange`); a
  self-closed `<w:p/>` is opened first, since it has no tags to put runs between:
  1. **the revised paragraph as it is**, every run keeping its formatting, only
     the objects that differ marked: a revised one with no counterpart inserted
     where it sits, an original one struck at its place, a run of text split
     there if need be (`authorObjectChange`). Objects are aligned on sort,
     census and place (longest common subsequence), so one footnote removed of
     three marks that one. Checked by this ADR's own check (both views' text and
     census) **and** a check that each view's zero-text objects sit where they
     sit in its document (`placesAs`). A place on a space inside an object with
     text, with only spaces between it and that object's edge, is the edge — the
     same place once whitespace collapses;
  2. failing that — an object that differs inside a link or text box, an object
     holding a space — the object-aware replace over an all-unchanged word diff,
     which marks the container whole and rebuilds the words, so their run
     formatting goes (limit 5). Held to `placesAs` too: its census cannot see an
     object that MOVED inside a container, and keeps such a container unmarked;
  3. failing that, **one** paragraph: the revised paragraph's properties, its
     runs inserted and the original's struck. Not a deleted paragraph plus an
     inserted one: a mark that cannot carry a revision (a section break,
     ADR-0035) was then authored twice.
  Never the revised paragraph verbatim (the redline silent) nor the text
  rebuild (the objects gone). At one place an inserted object reads before a
  struck one (ADR-0043), including just after a bookmark; a struck one keeps the
  original's order against a kept one beside it.
- **The census counts `w:object` as well** (`CENSUS`). It is now a sort this
  ADR compares, and the replace author's census decides whether a container is
  an unchanged pair: without it an embedded object inside a link was dropped
  from both views of an edited clause, and left unmarked by try 2 here.
- **Surfaces.** The change-list row names the objects (`-[footnote
  reference]`, a moved one as `-[…] +[…]`), before any number or style that
  moved with them; the report prints each side's text with its objects written
  in at their places (`diff/object-edits.ts`, free of the XML stack so a
  component can import it). *Go to change* anchors on the paragraph's words like
  any `property` op.

### Measured

- **Synthetic.** The seeded sweep gains two object-only seeds (200 pairs each:
  objects dropped, moved and added, no word changed), puts a manual number in a
  quarter of its paragraphs, asserts that a text-identical pair whose objects
  differ always has an op, and holds every object op to its census and places
  checks (previously skipped for every text-identical pair). On the object-only
  seeds: 171 and 169 object ops; 148 and 148 spliced, 23 and 20 through the
  replace author — chiefly pairs holding a link of only a space, which the
  splice declines — and 0 and 1 by the last resort (pinned as a ceiling). The
  mixed seeds reach 28, 31 and 22 object ops, none by the last resort.
- **Real documents** (counts only, `MUTATIS_REAL_FIXTURES`). One private pair
  carries no zero-text object: 0 object ops. The other produces **one** object op
  that did not exist before — a footnote reference moved to the end of a
  paragraph whose words did not change — classified `substantiveText`,
  spliced, and reading true in both views. That change was invisible until now.
- **Ablation, one line at a time** — every guard has a test that fails without
  it (23 reverts, final code): detection off (30 fail), authored as `equal`
  (24), class (3), place ignored (9), manual-number code (2) and its use in the
  splice (1), splice off (7), edge snap (2), insertion first (7), insertion
  first after a bookmark (2), kept order (1), places check on try 1 (1, the
  sweep) and on try 2 (2), `w:object` in the census (2), opening `<w:p/>` (1),
  one paragraph in the last resort (3), cascade guard (1), cross-ref guard (1),
  table guard (1), row text (2), report text (2), exact in-box placement (3),
  `within` (1). The bookmark rule was once
  removed as unobservable after a revert broke nothing; the next sweep found
  the shape it guards.
- **Review.** Pass 1 FAIL — an embedded object inside a link reported and left
  unmarked; a move across a manual number not detected; an emptied `<w:p/>`
  authored as a paragraph wrapped in `<w:ins>`; a section break authored twice.
  All four fixed and pinned above. Its other findings (stale descriptions, the
  untested sorts and helpers, the sweep's blind spot for a missed detection) are
  fixed in the same change. Pass 2 FAIL — a footnote moved one character at a
  text box's start, or from just before an equation into it, gave identical
  places (the "one character early" placement, then recorded as harmless, and a
  collapsed space). Both fixed: exact in-box placement, and `within`.
  Pass 3 **PASS**.

### Accepted limits of the amendment

1. **Sort and place only.** A picture swapped for another in the same place is
   the same sort at the same place, and the image lives in a media part behind a
   relationship id; not reported (pinned, *"RECORDS ITS LIMIT"*).
2. **Only these five sorts.** Other run content with no text is not compared: a
   symbol character (`w:sym`, a Wingdings tick box) and a non-breaking hyphen
   (`w:noBreakHyphen`) added or removed alone stay `equal` (pinned for `w:sym`).
   Objects WITH text are out too: a link or field put around words that did not
   change, or a link of only a space dropped where its space trims away, changes
   no text and no zero-text object. Comments and bookmarks are review metadata
   and markers, not content of the clause.
3. **The change-set and AI-readable exports omit it**, as they omit every
   `property` op (rule F1 is before/after text). The redline, the change list,
   the summary and the report carry it.
4. **A paragraph with no words** (a picture alone, emptied) has no anchor, so
   its row does not navigate.
5. **The replace path's alignment is unchanged.** A `replace` op whose
   zero-width objects differ in sequence still marks every one of them
   (§ Decision, "over-marking, never under-marking"); the alignment is the
   splice's only. Taking it to the walk would change the authored bytes of
   edited paragraphs, which this amendment did not set out to do.
6. **Try 3 on a field that closes in the next paragraph** authors two field
   `begin`/`separate` pairs (one inserted, one struck) against one `end`. Each
   view is balanced; how Word displays the markup unresolved is unchecked from
   this repo — limit 6 of the Decision (not opened in Word desktop). The
   sweep's vocabulary never puts an object inside a text box; the in-box
   placement is pinned by direct tests instead.
   And `within` is an ordinal over the paragraph's equations and text-box
   paragraphs, so an EMPTY equation added earlier in the paragraph — itself
   otherwise undetected — shifts it, and a footnote in a later box is reported
   as moved (`-[footnote reference] +[footnote reference]`) and marked by try 3.
   Over-reporting, never hiding.
7. **What a kept or inserted reference points AT is not checked** — limit 9(a)
   for a footnote; the same holds for a picture, whose `r:embed` names a
   relationship of the revised package while the redline is assembled from the
   original's (`writeout.ts`). Body-XML checks cannot see a package's
   relationships. Pre-existing, and not changed here.

### Found in passing, not this amendment's

- On one private real-world pair the accept-all view splits one paragraph in two
  (no zero-text object in it). Identical before this amendment; most likely the
  merge/split authoring of ADR-0032.
- A manual number removed or rewritten in an EDITED clause is emitted from the
  revised side unmarked (ADR-0028), so the rejected view loses it. Found when the
  sweep began generating manual numbers; the sweep keeps its number fixed.
- `emitInsert` wraps a self-closed `<w:p/>` in `<w:ins>` at body level (an
  inserted empty paragraph authors `<w:ins …><w:p/></w:ins>`). This amendment no
  longer routes into it; the defect itself is untouched.
