# ADR-0042: Compare — text identity anchors, non-text identity annotates (style and resolved numbering as comparable properties)

- **Status:** Accepted
- **Date:** 2026-09-12
- **Relates to:** ADR-0008 (categorise, never suppress — this ADR exists because of it); ADR-0026/ADR-0041 (the replace pipeline this must not disturb); ADR-0023 (numbering is detect-and-flag, never auto-convert); feature 08 (`spec/features/08-compare.md` §6 row 8, §9.4, §9.9)
- **Supersedes nothing.** ADR-0006/0023's "never auto-convert numbering" is untouched: this reads numbering to *compare* it and writes none.

## Context

A change that produces **no op at all** is not categorised, not filtered, and not
reported — it is absent. Under ADR-0008 that is suppression by omission, and the
the project principles ranks never-suppress as a hard product invariant.

Benchmarking on a clean private `v1`/`v1.1` pair (a real-world contract pair, not
published) found exactly that. A conventionally styled redline of the same two files
shows two numbering changes the engine reported nothing for:

| paragraph | original | revised |
|---|---|---|
| the outdented sub-clause (`¶13`) | `(vi)` | `(d)` |
| the sub-clause after it (`¶14`) | `(d)` | `(e)` |

(Paragraph indices, not headings. The pair is never committed, and neither is
prose taken from it.)

The cause is a single edit: one paragraph was outdented, `<subclause-L3>` →
`<subclause-L2>` (the real style ids are redacted throughout — they are the
document's own internals, and every argument here needs only that the two belong
to one family and differ by one level). Three facts make it invisible:

1. **The documents contain no `<w:numPr>` at all.** Every clause number is
   inherited from the paragraph style, and `preprocess/pararecord.ts` never
   reads the style chain. `numberingPrefix` is scraped from literal run text, so
   it is `''` for every automatically numbered paragraph in both documents.
2. **`ParaRecord.styleId` is captured and consumed by nothing.** A comment at
   `styles-diff.ts:6-7` claims per-paragraph `styleId` changes are "handled in
   the main body pipeline". They are not.
3. **The paragraphs are text-equal**, so `structuralDiff` emits `equal` and no
   classifier ever runs. `classify/taxonomy.ts` has a `formatting` class, but it
   only labels ops that already exist.

The spec has specified all of this from the beginning and it was simply never
built: §6 row 8 puts *"style id"* inside formatting-only; §9.9 says
*"per-paragraph `styleId` changes are likewise formatting-only"*; §9.4 says
*"the **numbering** delta is handled as a `crossRefRenumber`/formatting signal,
not a content change"*.

## Decision

> **Text identity anchors. Non-text identity annotates.**

**`ParaRecord.hash` stays exactly what it is** — a hash of prefix-stripped
visible text — and continues to drive anchoring and LCS. Two non-text
properties ride alongside it as *comparable properties of a matched pair*:

- **`numbering`** — the paragraph's effective numbering, resolved from direct
  formatting → `pStyle` → `basedOn` → the default paragraph style, **with
  `numId` and `ilvl` taken independently, each from its own nearest
  declaration**: ECMA-376 §17.7.2 merges the style hierarchy leaf by leaf, and
  these are two leaf properties that happen to share a `<w:numPr>` container.
  `numId="0"` at the nearest declaration is the removal sentinel. It carries the
  document-local `(numId, ilvl)` **and a separate comparison `key`**, because
  those are not the same thing — see *The identity that is comparable*, below.

  **This bullet described an ATOMIC `<w:numPr>` for seventeen review passes —
  *"a direct `numPr` if present, otherwise the first `numPr` found walking the
  ancestry"* — and that rule is a defect, not a simplification.** Under it a
  style overriding only `<w:ilvl>` (`basedOn` a numbered base style, changing
  nothing but the level — ordinary template construction) was skipped entirely,
  so every paragraph using it resolved to level 0 **on both sides** and a real
  list re-levelling produced **`summary.total` 0**. Measured at pass 17:
  `Body = <w:numId w:val="3"/>`, `Sub = basedOn Body, <w:ilvl w:val="2"/>`
  resolved to `{numId:'3', ilvl:0}` where the style says 2. That is ADR-0008
  suppression reached through the one chain shape the original acceptance
  sentence (*"a `pStyle` chain that REACHES a `numPr`"*) cannot express. The
  mirror went with it: a direct `<w:numId>` with no `<w:ilvl>` discarded the
  level its style declared and invented a `formatting` op between two documents
  differing only in whether a redundant `<w:ilvl>` is spelled out. The module had
  already documented per-leaf behaviour for a DIRECT `ilvl` and implemented the
  atomic rule for a style's; that asymmetry is what hid it for seventeen passes,
  and the suite stayed green with the old resolver restored.
- **`numberLabel: string | null`** — the **rendered** label (`(vi)`, `(d)`,
  `Clause 1.01`), produced by feeding those identities in document order to the
  **existing** `@shared/numbering` counter machine (`renderListNumbers`).
  the project principles: there is one numbering resolver and one counter machine, and
  RenderCore's painting half uses the same one.

  **`resolveLevel` follows `<w:numStyleLink>` (ECMA-376 §17.9.21), added at pass
  24 as the third ADR-0008 suppression of this work.** When numbering is attached
  to a Word LIST STYLE — how a firm template usually carries house numbering —
  the `<w:abstractNum>` the `<w:num>` points at holds nothing but a delegation,
  and the levels sit on a second one carrying `<w:styleLink>` (§17.9.28). Without
  the hop every level resolved to `undefined`, so `numberLabel` was `null` on both
  sides and the comparison key was the bare level: a redefinition that moves every
  number on the page reported **`summary.total` 0** against 3 for the identical
  edit on a non-delegating list, and an insertion above such a list lost its whole
  §6.1 cascade (total 4 with 1 cascade → 1 with none). The hop is resolved inside
  `numbering.xml` — §17.9.21's round trip goes through a numbering style whose
  `<w:numPr>` names a `<w:num>`, and Word always writes `<w:styleLink>` back on
  that definition, so the module keeps its one-part signature. It is **additive**:
  a definition that declares the level itself keeps it, so no resolution that
  worked before changes, which is what makes it safe for the renderer's gated
  consumers. Residue, stated not hidden: a `numStyleLink` with no matching
  `styleLink` anywhere still resolves to nothing.

  This is pass 22's lesson one mechanism along, and the reason that lesson is
  written as a rule rather than an anecdote: *"style-linked numbering … handled
  upstream/deferred"* was recorded as a **renderer** bound, and a renderer bound
  is not confined to the renderer while the comparison key resolves through the
  same module. It survived twenty-three passes because a re-LEVELLING is still
  caught whatever the level resolves to (`ilvl` is the key's first field), and
  every probe in the design re-levels.
- **`styleId`** — already captured; now read.

A new pass, `detectPropertyChanges`, runs after `structuralDiff` **and after the
merge/split re-pairing (2.25) and the §5.3 dissimilarity demotion (2.5)** —
`pipeline.ts` calls it stage 2.75, so it sees the FINAL set of `equal` pairs and
never examines one a later pass would have re-owned. It converts an `equal` op
whose two paragraphs differ in any of these into a new op type:

- **`OpType 'property'`** — a matched paragraph pair, **identical text**, one or
  more non-text properties differ. Never a `replace`: a `replace` means
  intra-paragraph *word* changes and carries `wordDiffs`, and pushing a
  format-only change through that path would invent a text delta that does not
  exist.

### The identity that is comparable

**`numId` must never be compared across two documents.** A `w:num` instance id is
document-local bookkeeping, and Word reassigns it freely: a Save-As onto another
template, a list pasted in from elsewhere, a counterparty re-creating the file.
The first implementation compared `(numId, ilvl)` directly and, on a document
with **no changes at all** — same body, same styles, same abstract definition,
instance id 3 against 7 — reported **six**. The evidence that nothing had changed
was already computed and not consulted: both sides rendered the same labels.

So the comparable identity is the **level plus the SHAPE of every level
definition the paragraph's number is rendered from** — a statement about the two
documents rather than about Word's internal ids.

Both halves of that took two corrections, and both were the same mistake at
different scales: a key narrower than the thing it claims to identify leaves two
different numberings equal, and a change found is then not reported.

- **The whole definition, not the readable parts of it.** A first cut covered
  `numFmt`/`lvlText`/`isLgl`, so changing `<w:start>` from 1 to 5 left the key
  equal and every clause reported as a *bystander* cascade — "its number moved
  because a sibling did" — when no sibling had moved and the list itself had been
  redefined. So `start` and `lvlRestart` joined the key.
- **The ancestors too, not just the paragraph's own level.** Under `lvlText` `%1.%2` a sub-clause's
  number is rendered from its ANCESTORS' formats as well as its own, so
  redefining level 0 moves every sub-clause's label while the sub-clause's own
  definition is untouched. Measured on a 2×6 list with level 0 redefined
  `decimal` → `upperRoman`: **twelve rows classified `crossRefRenumber`** — each
  asserting a sibling had moved — and because a redefinition moves no counter
  either, no shift and so no cascade to collapse them into. Twelve uncollapsed
  rows carrying a false statement. With the ancestors in the key, those
  paragraphs' identity changed, which is what happened.
- **But only what is actually RENDERED FROM.** Widening to every field of every
  level 0…`ilvl` was too much in the other direction. An ancestor's `lvlText` is
  never substituted into a descendant's label (`applyLvlText` reads the OWN
  level's template) and an ancestor's `isLgl` is never read for it, so re-wording
  a top-level prefix `%1.` to `Clause %1` reported **240 changes on a 40×5
  document where 40 paragraphs changed**. Direction was safe — over-report, never
  suppress — but `summary.total` is §8.2's headline number, and the evidence that
  nothing had moved (identical labels on both sides) was computed and not
  consulted, which is the pass-2 `numId` finding's exact shape.
- **Two fields as BEHAVIOUR, not as written.** `isLgl` renders every substituted
  token `decimal`, so its whole effect folds into the per-level format; stored
  beside it, it reported a list whose formats were already decimal. `lvlRestart`
  at level *k* treats every written value at or above *k* — and an absent element
  — identically, so storing the raw value reported **twenty** changes for adding
  `<w:lvlRestart w:val="1"/>` to a level-1 definition, with both label lists
  byte-identical.
- **But the OWN level's DECLARED `numFmt` is carried as well as its folded one,
  and that is not a contradiction of the rule above.** `render-list-numbers`
  applies `isLgl` to the paragraph's own token as well as to its ancestors and
  **Word does not** — measured against Word's own rendering, `decimalZero` +
  `isLgl` over an `upperRoman` ancestor gives `1.01`, the ancestor converted and
  the own format kept. So for the own token the folded value is not "the
  behaviour"; it is *this build's* behaviour, and it differs from the document's.
  Carrying only it made the key blind exactly where the renderer is blind, and an
  edit to the own level's `numFmt` under `isLgl` produced **no op at all**
  (`summary.total` 0 against 3 for the same edit without `isLgl`) — found at the
  twenty-second review pass, on a shape present three times in the benchmark
  packages. Both values are carried: the folded one because key equality must
  keep forcing `renders` equality, the declared one because Word renders it. The
  declared field is pushed **unconditionally** — pushed only under `isLgl` it
  made the element's presence visible and re-created the false positive the rule
  above exists to prevent.
- **And the sequence, which is not about this paragraph at all.** The counter
  machine walks every paragraph in document order and SKIPS one whose level does
  not resolve, so it neither advances that level nor resets the deeper ones.
  Deleting a level definition therefore moves the counters of the paragraphs
  below it while touching none of their own fields. Measured: removing level 1's
  `<w:lvl>` from a 3-level list ran two `(%3)` sub-points on from `(a)`,`(b)` to
  `(c)`,`(d)`, and those two reported `crossRefRenumber` — a sibling moving —
  with no sibling anywhere. So the key also carries, for every level 0…`ilvl`,
  whether that level has a definition at all.

**Why narrowing this key is legitimate, where the first two attempts were not.**
`PropertyChange.numberLabel` is computed from the two **rendered labels**, never
from the key. **So within this pass's own reach — the one-to-one `equal` pairs
`detectPropertyChanges` examines — a label that MOVED is always reported,
whatever the key says**; for those, the key decides only *which of §9.4's two
sentences* the op gets — *"this paragraph's numbering was changed"*
(`formatting`) or *"its number moved because a sibling's did"*
(`crossRefRenumber`). That asymmetry — misattribution possible, disappearance
not — is the argument every width decision above rests on, and it is scoped
exactly as written: it covers labels that MOVED, **on pairs this pass reaches**,
and nothing else. The clause about reach is not a hedge: Accepted limit **Four**
is the case outside it, and a third review pass found this sentence unqualified
and read it as a guarantee it does not give.

**The scope of that guarantee, stated exactly — an earlier draft of this
paragraph overstated it as "narrowing cannot hide a change", and the paragraph
immediately below contradicted it.** The label comparison protects only changes
the reader can SEE. A definition change that moves no rendered label is detected
by the key and **by nothing else**, so for that class a narrowing — or a key
collision — really does hide it. Falsified by measurement, not reasoned about:
two `numFmt` values this build cannot express, with different `start` values, on
the LEVEL-0 definition of a `%1.%2` list — the `%1` token renders `''` on both
sides and both labels read `.1`; the shipped key reports one `formatting` op, and
with the key's field separator emptied it reports **none**, `summary.total`
1 → 0.

*(That probe was rebuilt at the twenty-third pass. Until then it collided the
paragraph's OWN `numFmt` against its own `start`; the field added at pass 22 —
the own level's DECLARED format — trails every other field, a difference there
has nothing after it to absorb the boundary, and the collision stopped being one.
Both separator tests went on passing, and testing nothing, for one commit. The
collision now lives in an ancestor level's block, which has no trailing field.
There were two separator constants; the second one separated `ilvl` from the
fields and could never have had an ablation — a boundary that appears only at the
head of the string can change an answer only when one side's `ilvl` decimal
string is a proper prefix of the other's, and Word's nine levels are all single
digits. It is gone: the level is the first ordinary field.)*

**And there is no bridging generalisation. Two have been written here and
measurement falsified both; do not write a third.** The second said the key's
bound — levels 0…`ilvl`, "what the renderer reads when rendering this paragraph"
— is a superset of everything that can move this paragraph's label, so staying
inside it loses nothing a reader could see. It is not. Accepted limit **Two**,
75 lines below, is the counterexample and always was: a level definition deleted
DEEPER than a paragraph's own level moves that paragraph's counter, the key does
not carry it, and the label moves anyway — measured, `1.2` → `1.1`, reported and
misattributed. (The renderer's literal read set is wider still: `renderListNumbers`
resolves every deeper level in `counters` while rendering this paragraph, and the
key carries none of them.)

What is actually true is two facts and no rule joining them:

- **On a pair this pass examines, a change that moves a rendered label is always
  reported**, whatever the key carries — by the label comparison, which never
  consults the key. Narrowing can misattribute such a change; it cannot make it
  disappear. That is the whole justification for narrowing, and it is enough.
  **Outside that reach it says nothing** — a paragraph whose words also changed
  pairs as a text op, never reaches `detectPropertyChanges`, and its renumbering
  is not reported at all (Accepted limit Four, now pinned by a test). Three
  successive review passes have falsified an over-general version of this
  sentence; the qualifier is the correction, not decoration.
- **A change that moves no rendered label is reported only if the key carries the
  field that changed.** Narrowing there really does hide it, and a key collision
  does the same — pinned by a test that empties the separator.

Every width decision above is therefore a judgement about which *invisible*
definition changes to catch, made against the recorded costs, and not a
deduction from a safety property. Treat it that way.

**Where the key still exceeds the label, stated rather than hidden.** Two `numFmt`
values this build cannot express (`cardinalText`, say) both render `''`, so
swapping one for the other changes the key and no label; and a definition change
is reported for every paragraph drawing on the changed level, including the
minority whose rendered label happens to be unaffected by it. Both are statements
about the *definition*, which really did change and which Word would show, and
both err towards reporting — the side ADR-0008 wants. **These are exactly the
class the paragraph above says the label comparison does not cover**, which is
why they are the key's job and why the key must not be narrowed below the
renderer's own read set.

`numId` survives on the record because the counter machine needs it to group
counters *within* a document, and for nothing else.

This also settles a second question correctly and for free. When a level's format
is **redefined** (`decimal` → `lowerRoman`), every clause's own number changes.
Under a label-only comparison all of them classified `crossRefRenumber`, whose
documented meaning is *"this paragraph did not change; its number moved because a
sibling did"* — and no sibling moved. Every definition the label is drawn from is
part of the key, so they now classify `formatting`, at every level the
redefinition reaches: the numbering of these paragraphs was changed.

One consequence worth stating, because it is what a reader sees. A list-definition
change produces one `formatting` op **per affected paragraph** and no cascade —
correctly, since none of them is a bystander. §6.1's collapse is for paragraphs
whose numbering did *not* change, and these are not those. What keeps a
redefinition from filling the navigation list is the mechanism the spec already
has for exactly this: *Hide formatting-only* (§8.2), **ON by default**. Every op
still exists and `summary.total` still counts every one, so ADR-0008 holds — the
change is categorised and filtered by a switch the user can see, never dropped.

## Accepted limits

**One. A clause moved between two lists whose definitions are identical does not
change its key.** Cross-document that is indistinguishable from Word having
reassigned the instance id — and choosing the other way means inventing a change
on every Save-As, which is the worse error.

Stated precisely, because an earlier draft of this paragraph overstated it: the
*move* is still reported whenever the labels move, which is the case a user
actually creates. Word's "restart numbering" writes a second `w:num` carrying
`<w:lvlOverride><w:startOverride w:val="1"/>`, which renumbers every clause that
follows; measured on six sub-clauses with the last three moved onto it, that
reports three `crossRefRenumber` ops collapsed into one cascade row —
*"Clause numbering (level 2) shifted (-3): 3"* — with the labels `1.4`,`1.5`,`1.6`
→ `1.1`,`1.2`,`1.3`.

**That measurement was not evidence for the `startOverride` fold when it was
taken, and it is now.** Until pass 25 `render-list-numbers.ts` keyed counters per
`numId`, so the fresh instance's level-1 counter started at 1 with or without the
override and deleting the fold left those labels unchanged — which is why an
earlier draft citing it as evidence was corrected. That keying was itself the
defect (see the correction below): counters belong to the abstract definition, so
without the override those three sub-clauses now continue `1.4`,`1.5`,`1.6` and
the restart is what produces the figure above. **Both mechanisms are load-bearing
and each is pinned:** `resolveLevel`'s fold puts the overridden start in the
comparison KEY, and `startOverrideFor` restarts the shared counter so it reaches
the LABELS. A start the definition does not carry shows the difference between
them — `<w:startOverride w:val="9"/>` gives `1.9`,`1.10`,`1.11`, and because
`start` is a key field that is a list REDEFINITION (`formatting`) rather than a
renumbering, so §6.1 does not collapse it.

*(This paragraph asserted the per-`numId` keying in the present tense for one
pass after the keying was removed, while the test file it cites had been turned
around in the same commit. The most-repeated failure — a change landing in
some of its homes — reaching the ADR, which is the record that outlives the
plan.)*

What is unreported is the residue, and it is narrower than "a move between
identical lists": the move must also leave every label where it was. One
sub-clause moved to a second instance of the same definition does that — it is
the only member of that level's sequence either way, so the label is `1.1` on
both sides and nothing a reader could see has happened to the numbering. A
detection limit is not a suppression; nothing is found and then hidden.

**This paragraph used to say six moved that way are reported, and that was a
fabrication of the renderer's — corrected at pass 25.** `render-list-numbers.ts`
keyed counters per `numId` while Word continues them across `w:num` instances
sharing an `abstractNum`, so moving six sub-clauses to a second instance
"restarted" them at `1.1` and Compare dutifully reported a `-3` cascade over
numbers no reader would ever see. The ADR recorded that over-report as a
curiosity — *"named so the next reader does not cite it as a measurement of
Word"* — and recorded nothing at all about its mirror, which was an ADR-0008
suppression: a clause inserted into one group moved the other group's numbers in
Word and produced **no op at all** (`summary.total` 1 against 3 for the identical
edit on a single instance).

**Counters now belong to the abstract definition** (`counterGroupOf`, delegation
followed so a list style's two routes are one sequence), and a `<w:startOverride>`
restarts the shared counter on the instance's first **TOUCH** of that level
(`takeRestart`) — touch and not *use*, because reading a level as an ancestor
token is as much a use of the instance as sitting at it. That is what the element
is for, and it is what makes Word's restart button and its continue button both
work. Both directions are pinned, in the renderer's own tests and end to end. The
fourth ADR-0008 suppression of this work, and the third reached through a bound
filed as *"the renderer's"*.

*(The word matters because the first version of this fix read "use" as *sitting
at the level*, and pass 26 measured what that costs: an instance restarting at 1
whose first paragraph is a sub-item rendered its ancestor token from the counter
it was restarting away from — `1. 2. 2.1 1. 2.` where Word gives
`1. 2. 1.1 2. 3.` — and an instance that never reaches the overridden level at
all never restarted it. Wrong with the override on every level too, so not the
narrow case it first looked like. Ablated: removing the ancestor half fails 1 of
15; removing the once-only consumption fails 8.)*

**Two. A level definition deleted or added DEEPER than a paragraph's own level
moves that paragraph's number without moving its key.** The key covers what the
renderer reads and writes when rendering *this* paragraph — levels 0…`ilvl`.
Levels below that are touched by other paragraphs' render calls: the counter
machine initialises every shallower counter when it renders a deep paragraph, and
skips a paragraph whose own level does not resolve. Measured on a level-0 head, a
level-2 point and a level-1 sub-clause in that order, with level 2's `<w:lvl>`
removed: the sub-clause moves `1.2` → `1.1`, is reported, and is reported as
`crossRefRenumber` — a bystander of a sibling that did not move. At twenty such
groups it is 20 `formatting` ops for the level-2 points plus one collapsed
cascade row for the sub-clauses.

Both halves of that are deliberate. Reporting it is required. Classing it as a
bystander is wrong, and the alternative — widening the key to every level the
scheme defines — marks every member of a list whenever a definition no member
renders is edited, which is the 240-vs-40 failure above in a different costume. A
per-paragraph key cannot hold a sequence's state; §6.1's cascade row is where the
sequence is visible. Pinned by *"RECORDS its limit: a level deleted DEEPER than
the paragraph reads as a sibling"* in `compare.diff.property.test.ts` — change
that expectation only together with this section.

**Three. `w:pPrChange` is not reverted by `rejectAll`.** `preprocess/flatten.ts`
unwinds `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`, and `quarantine.ts` drops those
plus the comment markers. `w:pPrChange` — Word's own record of a tracked
PARAGRAPH-PROPERTY change — is in neither set, so under `flattenRevised:
'rejectAll'` a paragraph's TEXT reverts and its `w:pStyle`/`w:numPr` do not.
Measured on a tracked outdent `C2` → `C1`: under `rejectAll` the revised record
still reads `styleId C1`, label `1.`, and the comparison reports one `property`
op — a re-levelling the user asked to reject.

It is recorded here rather than fixed because this ADR is what made it
observable, not what caused it. Before ADR-0042 `styleId` was captured and read
by nothing and numbering was not resolved at all, so a surviving `pPr` changed no
output. Spec §5.0 names only the four revision wrappers, so this is a limit and
not a violation — and the DEFAULT (`acceptAll` on both sides) is unaffected;
`rejectAll` is a panel choice. Fixing it means teaching `flatten.ts` to swap a
paragraph's `pPr` for the one nested inside its `w:pPrChange`, which is a
behaviour change to the module every comparison path runs through. Pinned by
*"a tracked paragraph-property change under rejectAll"* in
`compare.diff.property.test.ts`, which asserts the surviving level explicitly, so
the behaviour is known rather than discovered.

**Four, of a different family.** `detectPropertyChanges` examines `equal`
ops only, so a property change riding on a paragraph that ALSO moved or ALSO had
its words edited is not reported as a property change. A clause relocated
verbatim *and* promoted a level reports the move and not the promotion; a clause
whose text and style both changed reports the text. The paragraph is never
silent — it has an op either way — but the property half of what happened to it
is missing from that op. Extending `replace` and `move` to carry a
`propertyChange` alongside their existing evidence is a real design question (one
op, two kinds of change, one `classification`) and deserves its own decision
rather than an addition at the end of this one. Recorded here rather than found
later.

**Pinned by a test since the seventeenth review pass** — one clause inserted and
one of the clauses below it also reworded: the untouched clause's `2.` → `3.`
reports as a `property` op, and the reworded clause's `3.` → `4.` is reported by
nothing at all (`compare.diff.property.test.ts`, *"a property change riding on a
text change"*). It was the only one of these four limits with no test, and it is
the one most likely to be "fixed" silently — which is why the test asserts the
limit rather than the fix.

Classification follows the spec's own two answers:

| what differs | class | why |
|---|---|---|
| `styleId`, or the numbering **identity** (level + level-definition shape) | `formatting` | §6 row 8, §9.9 — the paragraph itself was restyled or re-levelled |
| only the rendered **label**, identity unchanged | `crossRefRenumber` | §9.4 — the paragraph did not change; its number moved because a sibling did. That is a renumbering cascade, and §6.1's collapse machinery already keys on this class |

That split is what keeps a long document readable. Insert one clause and every
later label at that level shifts; if all of them classified `formatting` they
would bury the one real edit, which is the "malpractice-grade pain point" §6.1
exists to prevent. On the benchmark pair the split reports the outdent as
`formatting` and the knock-on `(d)`→`(e)` as a collapsible cascade.

**"Collapsible" had to be made true, and the first implementation was not.**
`computeCascades` derives a cascade member from `op.wordDiffs`, and a `property`
op has none — its two paragraphs are word-for-word identical. Measured on a
synthetic list of twelve sub-clauses with one inserted at the head: **twelve
uncollapsed rows and zero cascades**, i.e. this ADR causing the exact defect it
cites §6.1 to avoid. So a pure cascade now carries its own evidence
(`PropertyChange.cascade`) and `computeCascades` has a branch for it. Two
consequences follow from what a cascade needs:

- **The counter path, not the label — and not the own-level counter either.**
  §6.1 groups by a uniform *shift*, and `(d)` → `(e)` cannot be subtracted.
  `renderListNumbers` has always computed the counters behind the token; it now
  returns them all, for levels 0…`ilvl` (`ListNumber.path`), and the shift is the
  shallowest level at which two paths disagree — **among the levels the label
  actually shows** (`ParaNumbering.renders`), not among every level of the path.
  See *"…and only among the levels the label prints"* below for why the second
  half is not a refinement but a correction.

  Returning only the paragraph's *own* counter was the first attempt, and it is
  wrong in the most common house style there is. Under `lvlText` `%1.%2` — "1.1",
  "1.2" — a sub-clause renumbers when its PARENT does and its own counter never
  moves: insert one top-level clause and `1.1` → `2.1` while the level-1 counter
  holds at 1. Measured on a six-sub-clause probe: **shift 0 on all six, zero
  cascades, seven uncollapsed rows**, the inserted clause's own row left alone in
  a group of one below the threshold — this ADR causing, in the dominant SG/E&W
  numbering style, the exact §6.1 flood it cites §6.1 to avoid. Reading the first
  differing level fixes both halves at once: all seven rows report `(level 0, +1)`
  and collapse into the single fact they are. Pinned by
  `tests/unit/compare.diff.property.test.ts` → *"a cascade carried by an ANCESTOR
  level"*, which fails on the own-level derivation and passes on this one.
- **…and only among the levels the label PRINTS.** Reading the first differing
  level of the whole path is the *other* half of the same mistake. §6.1's row
  names a level and a shift and prints up to three example labels beneath it, so
  the two have to agree — and they do not when the level that moved is one the
  template never shows. Measured on a clause list split by a new top-level
  clause, sub-level `lvlText` `(%2)`:

  ```
  O: 1. | (a) | (b) | (c) | (d) | (e)
  R: 1. | (a) | (b) | 2. | (a) | (b) | (c)
  header: "Clause numbering (level 1) shifted (+1) — 3 paragraphs"
  rows:   (c) → (a)   (d) → (b)   (e) → (c)
  ```

  Three rows that each moved **minus two**, under a header claiming plus one: the
  `+1` is the parent clause's counter, and `(%2)` renders none of it. The two
  templates disagree and only `%1.%2` had been probed — the ancestor fix above
  was written with `%1.%2` fixtures exclusively, and under those the `1 → 2` is
  on the page, so header and rows agreed by accident.

  So `ParaNumbering.renders` carries the levels the label actually shows, and the
  scan runs over those, shallowest first. Two filters, each a measured
  contradiction: the `lvlText` must SUBSTITUTE the level (`%N`), and the level's
  effective format must RENDER something — a substituted ancestor at
  `numFmt="none"`, or at a format this build cannot express, occupies its slot
  and prints nothing, and produced the same `+1`-over-`−2` header. Both filters
  read the definition only and never a counter, which is what lets one side's
  answer stand for the pair: equal identity ⇒ equal template ⇒ equal formats.
  (Stated residue: letter and roman formats also render `''` at a counter of zero
  or less. `<w:start w:val="0"/>` empties only that level's FIRST item, so a
  header that contradicts its rows needs `start` at `-1` or below — constructed,
  and it does produce one. A predicate over the definition
  cannot see a counter, and moving it to render time would let the two sides
  disagree — the one property the caller relies on.) Pinned by *"the cascade
  header agrees with the labels beneath it"*, which asserts `(%2)`, `%1.%2` and
  the `numFmt="none"` ancestor together, so none can be satisfied by breaking
  another.
- **The level is part of the group.** Two levels cascading at once are two
  cascades, so the token kind carries the level (`Clause numbering (level 2)`).

`CascadeAnnotation` gains `source: 'crossReference' | 'numbering'` because these
are two different sentences to a reader: one is the document *pointing at* a
clause whose number moved, the other is the clause numbers themselves moving.
Reporting the second as "references renumbered" would claim a reference the
document does not contain.

## Rationale

**The authored document does not change, by construction.** `emitEqual` already
*"prefer[s] the revised record (it carries the current numbering)"*, so the
redline has always carried the revised style and accepting it has always given
the right formatting. There is no `emitProperty`: `author.ts` falls
`case 'equal': case 'property':` through to the SAME `emitEqual` call, which is
stronger than two functions that agree, because they cannot drift apart. This
change is therefore **purely additive to `ops`**: the authored XML is
byte-identical before and after, which is both the acceptance test and the
reason the risk is low.

*(Amended 2026-10-18 by ADR-0045 § Amendment 1. The two cases no longer fall
through to one call: `case 'property'` sends a property op whose zero-text
objects differ — a footnote or picture added, removed or moved, which is a
change to the CONTENT — to `emitObjectChange`, which marks it, and every other
property op to the same `emitEqual` as before. So the paragraph above still
holds, byte for byte, for every property this ADR defines; it no longer holds
for `property` ops as a whole.)*

**Why a new `OpType` rather than a flag on `equal`.** The summary counts by
classification over ops, and navigation targets ops; a change that is not an op
cannot be counted or navigated to. `OpType`'s consumers end in an
`exhaustive: never` switch, so adding a member makes the compiler name every
site that must handle it — two of them.

**Why `hash` must not absorb numbering.** The obvious implementation — fold the
number into the hash — is a trap. Every paragraph whose number shifted would
stop matching its counterpart, the structural diff would degrade into
whole-paragraph delete + insert, and the result would be issue 1 of the parity
brief, caused deliberately. Making a change visible must never make the
alignment worse.

## Consequences

Measured on the benchmark pair, 2026-09-12.

- **Two ops appear that did not exist before** — `¶13` `formatting`
  (`<subclause-L3>`→`2`, `(vi)`→`(d)`) and `¶14` `crossRefRenumber`
  (`(d)`→`(e)`) — which is exactly the two numbering changes shown above. `summary.total` **5 → 7**
  (measured, not recalled: the figure first recorded here, 6, was wrong).
- *(Amended by ADR-0045 § Amendment 1, 2026-10-18: a `property` op whose
  zero-text objects differ — a footnote or picture added, removed or moved — is
  marked, not authored as `equal`. Every other `property` op still is, and the
  byte-identity below still holds for them.)*
- **The authored redline is byte-identical**: authoring the same ops with the
  property ops mapped back to `equal` reproduces `authoredBodyXml` to the byte.
  Re-verified after the cascade and key fixes below by comparing the whole
  pipeline's output on the benchmark pair before and after: identical. The
  comparison is only meaningful for fixed inputs: the author string
  (the `SG_PROFILE` default `"Mutatis Compare"`), the date and the id start are
  all inputs to the authored bytes, so a bare length is not a usable pin. Marks
  stay at 20; no markup was added.
- **Nothing else moved.** Synthetic corpus 0.11 redundancy / 89 marks / 8-of-9
  ideals; 4,000-pair sweep 66 offenders / 0.0388 / 2.3287; private pair
  reconstruction true both sides, `wholeParaChurn` 0, `alignmentRate` 1,
  `charOffendingBlocks` 5. Every figure byte-identical to the pre-change reading.
- **A renumbering cascade of `property` ops collapses like any other**: twelve
  shifted sub-clauses become one expandable row, every member tagged
  `collapsedInto`, not one op dropped. Measured for both shapes a cascade takes —
  twelve `(a)`→`(b)` siblings whose own counter moved, and seven `1.1`→`2.1`
  rows carried by an ancestor's, which collapse 7-of-7 into one `level 1` row. Below `renumberThreshold` (3) they stay
  as individual rows — the same rule cross-references follow, and the reason the
  benchmark pair's single `(d)`→`(e)` remains a visible row rather than
  disappearing into a cascade of one.
- **One instrument change.** `alignmentRate` enumerates op types to decide what
  counts as matched, so it read the new type as *unmatched* and fell to 0.94
  while the pairing was provably identical. `property` joined that list. This is
  a measure learning about a new type, not a ratchet being moved — and it is
  recorded here because it is exactly the kind of edit that should never pass
  unexplained.
- **`preprocessSide` gains an optional numbering context** (`stylesXml`,
  `numberingXml`). Optional because the Text and PDF paths have neither; they
  pass NO context and behave exactly as before — `ParaRecord.numbering` is null
  and no `property` op can be emitted, so a source that cannot know never reports
  a change it cannot see. **An earlier draft said "absent and null compare
  alike"; they do not, and the difference is worth stating.** A context built
  from two null parts still resolves a DIRECT `<w:numPr>` — measured,
  `{ numId: '3', ilvl: 1 }` with `label: null` — because reading a direct
  `<w:numPr>` needs neither part. That is the better behaviour, not a defect: a
  re-levelling is still detected in a package whose `numbering.xml` is missing.
  It is safe because every call site builds contexts for the two sides
  identically, so neither shape can produce a one-sided difference.
- **RenderCore is deliberately not wired to the new resolver**, though it has the
  same gap at `parse-body.ts:147`. Painting a label is the renderer's half, RenderCore
  Phase 2 is gated on a live-host smoke test, and the `isLgl` defect below would
  ship with the labels. The resolver takes a structural `StyleNumberingSource`
  that `parseStyles`' existing `StyleTable` already satisfies, so wiring it later
  is an argument, not a rewrite.
- **The AI-readable export does not carry `property` ops.** Rule F1 is
  before/after *text*, and a property change's before and after are the same
  string; representing it honestly needs a field `ChangeEntry` does not have, and
  that schema is feature 11's published contract. The change is still in `ops`,
  in the summary's formatting count and in the exported report. Recorded as
  feature 11's to add.
- **Field identity and section properties are dropped** on the benchmark pair.
  Section properties measure **exactly zero**: one `sectPr` per side,
  byte-identical, with the same footer references and no header parts. Field
  identity measures **zero changes worth reporting, which is not the same
  statement and must not be written as one**: both sides carry the same
  field characters, and the sole delta is a formatting-preservation switch on one
  field, which changes no displayed text and which a field-identity comparison
  would report as a change to a clause the author never touched. Building it
  to report that is worse than not building it. The delta is named here
  so that the reason this item was dropped is not lost. They were real on the superseded fixtures; they
  are not built speculatively.

## Rejected alternatives

- **Fold numbering into `hash`.** See Rationale — it converts a visibility gain
  into an alignment catastrophe, and it is the single most tempting wrong turn
  here, which is why it is written down.
- **Emit a `replace` with an all-`eq` word diff.** Cheaper — no new `OpType`,
  no switch changes. Rejected because it lies in the data model: every
  downstream consumer that sees `replace` is entitled to assume a text delta,
  and `demoteDissimilarReplaces`, `refineReplaceBlocks` and the classifier all
  reason about `wordDiffs`. A format change is not a degenerate text change.
- **Leave the op as `equal` and hang a `propertyChange` annotation off it.**
  Rejected for the reason a new type was chosen: the summary counts ops and the
  navigation list targets ops, so an annotated `equal` would still be
  unreportable and unnavigable — the defect, restated.
- **Report every rendered-label change as `formatting`.** Simpler, and it would
  report every label change individually. Rejected: it makes a
  one-clause insertion produce one op per following clause, which §6.1 names as
  the worst failure mode in the product. The identity/label split costs one
  comparison and keeps the cascade collapsible.
- **Compare `(numId, ilvl)` directly.** The obvious reading of "numbering
  identity", and wrong: measured at six reported changes on an unchanged document
  whose only difference was Word's own instance id. Rejected on that measurement.
- **Key on part of the level definition** — `numFmt`, `lvlText` and `isLgl`, the
  three that decide what the label looks like. Rejected on measurement too, and
  it is the same mistake one attribute along: `<w:start>` 1 → 5 leaves those
  three equal, so a redefined list reported as four bystanders shifted +4 with no
  sibling anywhere. The key carries every field that can move the rendered label —
  see the Decision above for the full rule, which turned out to be neither "part
  of the definition" nor "the whole definition".
- **Diff the numbering by comparing `numberingPrefix`.** It is already on
  `ParaRecord` and needs no new resolution. Rejected because it is scraped from
  literal run text, so it is `''` for exactly the automatically-numbered
  paragraphs this ADR is about — it can only see manually typed numbers, which
  are the case ADR-0023 handles elsewhere.
- **Write a second numbering resolver inside Compare.** Rejected on
  the project principles. `@shared/numbering` already parses `numbering.xml`, resolves
  `(numId, ilvl)` through `lvlOverride`, and runs the counter machine including
  `lvlRestart`. Only the style-chain hop was missing, and it is added there so
  RenderCore's painting half gets it too.

## Open

- **Direct formatting is still invisible — §6 row 8 has four limbs and this ADR
  gave it two.** Row 8 is *"entirely within `<w:rPr>` / `<w:pPr>` (font, bold,
  italic, colour, spacing, indent, style id)"*, and §10 T2 asks for a bolded
  heading to be classified `formatting` and hidden by the default filter.
  `comparePropertyPair` compares `styleId` and the numbering key, so the STYLE-ID
  and NUMBERING limbs now produce an op and the `<w:rPr>` and direct-`<w:pPr>`
  limbs produce none at all. Measured at the twentieth review pass: a bold-only
  change and an indent-only change each give `summary.total` 0, `ops` `[equal]`,
  `byCategory.formatting` 0 — so T2's *"hidden from the change list"* is
  satisfied only in the sense that there is nothing to hide. The gap predates
  this ADR and measures identically at the base commit; what this ADR changed is
  that half of row 8 now works, which makes the other half look like a bug rather
  than an absence. Building it is a fifth property class: direct formatting
  differs on almost every paragraph of a real redline (Word rewrites `<w:rPr>`
  freely), so it needs its own noise answer before it can be reported at all, and
  under ADR-0008 the answer cannot be "hide it". That is a decision, not an
  addition to this one. **Recorded so it is owned; the original evidence log
  carries the measurement.**
- **`numbering.xml` is never diffed as a PART, and `styles.xml` is.** §9.9 gives a
  style-definition change its own summary note; a list-definition change has no
  counterpart. So a change to a list definition reaches the user ONLY through the
  fields `numberingKey` carries per paragraph, and everything outside that key
  produces no output at all. Measured at the twenty-sixth review pass, on the two
  bounds this work had predicted would bind Compare: a bullet glyph written in
  `w:lvlText` is compared and reports two `formatting` ops, but the SAME glyph
  changed through the level's `<w:rPr><w:rFonts>` — Symbol to Wingdings, `lvlText`
  untouched, which is how Word's bullet picker writes it — gives `summary.total` 0
  and `ops` `[equal, equal]`. The same is true of a level's `w:pPr`/`w:ind` (the
  hanging gutter that the original acceptance criterion names), `w:suff`, and any
  level no paragraph currently uses.

  It is a detection LIMIT and not a suppression — nothing is found and then
  hidden, because nothing reads those bytes — and it is the same shape as the row-8
  gap above: a class of change the pipeline has no input for. It is recorded here
  because three of this work's four ADR-0008 suppressions were bounds filed
  elsewhere as harmless, and the honest way to stop that pattern is to write the
  measurement down where the next reader will look rather than to leave the
  asymmetry with `styles.xml` for them to rediscover. Closing it is a part-level
  `diffNumbering` beside `styles-diff.ts` plus a §9.9-style note — a decision with
  a spec change behind it, not an addition to this one.
- **The batch panels do not diff the auxiliary parts at all.** `diffPartsIfEnabled`
  is called by the single-pair Word driver only, so a footnote content change —
  live at the default config, and now counted into `summary.total` — is in the
  single-pair panel's total and absent from a batch row's `changeCount`. The
  divergence predates this ADR (neither counted them before) but this ADR widened
  it, and it is the same "one product, two answers depending on the panel"
  objection this ADR raised for numbering at the batch seams. Wiring it in is a
  behaviour and cost change to a completed piece of work (the at-scale batch work): an extra part diff per
  pair at N up to a hundred, and a changed `changeCount` for rows whose artefacts
  are already counted against a ceiling. That is a separate decision, not an
  addition here.
- **The navigation half of §9.8.** A footnote content change is substantive and
  compared by default, and `mergePartsIntoSummary` now counts it into
  `summary.total` and `byCategory` — but the part's ops stay out of the navigable
  change list, so the reader can see that a footnote changed and cannot click
  through to it. The reason is concrete rather than an oversight: `goToChange`
  resolves an op's range against the BODY paragraph arrays, and a part op's range
  indexes its own part's paragraphs — the same integer meaning a different
  paragraph — so merging them without a part-aware seam would scroll the reader
  to an unrelated clause instead of to the footnote. Closing it means carrying
  per-part records through `CompareResult` and giving navigation a part
  dimension. That changes a public shape and a contract, so it is a decision of
  its own and not an addition at the end of this one.
- **§9.9's "filterable" half is unmet.** `diffStyles` returns a synthetic
  formatting-only `DiffOp` and `driver.ts` consumes only its `note`, so a
  `styles.xml` change is shown on the summary panel but is not counted into
  `summary.total`/`byCategory.formatting` and *Hide formatting-only* does not
  reach it — §9.9 asks for both. Nothing is suppressed: the change is detected
  and reported, in a note the reader always sees. An earlier task routed
  the two discarded diffs into `ChangeSummary`, which it did; the filter half was
  never in its acceptance. Closing it runs into the same wall as the §9.8 item
  below — the op has `origRange`/`revRange` of `null`, because a style definition
  is not at a paragraph, so counting it into the headline adds a number with no
  row to click. Gating the note on the filter alone is a one-line change and is
  the cheaper half; both are behaviour changes and belong to a decision, not to a
  review pass.
- **The parts diff never runs `detectPropertyChanges`.** `diffPartXml`
  (`parts-diff.ts`) calls `structuralDiff` alone and is given no numbering
  context, so a `styleId` or numbering change on a FOOTNOTE paragraph — compared
  by default under §7.5/§9.8 — produces no op, while the identical change in the
  body produces a `property` op. It is a detection limit rather than a
  suppression (the parts diff never looks, so nothing is found and hidden), but
  it is the body/parts twin of the batch divergence this ADR closed, inside a
  single comparison. What it costs is smaller than an earlier draft of this entry
  claimed: that draft deferred it partly on *"deciding what a property op's range
  means in a part"*, but this diff already answered that for every other part op
  — `mergePartsIntoSummary` counts them into `total`/`byCategory` and leaves them
  out of the navigable list, and a footnote `property` op would simply inherit
  that decision. What genuinely remains is a numbering context per part (footnote
  numbering is its own counter sequence) and a behaviour change to a comparison
  that is on by default. Recorded here so it is owned, on its real cost.
- **`isLgl` is applied one level too far** in `render-list-numbers`'s token
  loop (`level.isLgl ? 'decimal' : …` covers the level's own token as well as
  the inherited ones). On the benchmark fixture that renders `<unit> 1.1` where
  Word renders `<unit> 1.01` (the template's literal prefix word is the real
  document's and is redacted; the arithmetic is the evidence). **It DID affect
  this ADR, and this bullet said it could not for twenty-two review passes.**
  `numberingKey` mirrored the renderer's fold, so the key and the label went
  blind together and an edit to the own level's `numFmt` under `isLgl` produced
  no op at all — `summary.total` 0 where the same edit without `isLgl` gives 3,
  on a shape present three times in the benchmark packages' own `numbering.xml`.
  The key now carries the own level's DECLARED format as well as the folded one;
  see *The identity that is comparable*. The renderer's fold is untouched and
  remains the renderer's — this changed what is compared, not what is painted.
- **`pPrChange` markup** — the Word-native representation of a paragraph-property
  revision, which would give the reader a "Formatted:" balloon in the document
  rather than a summary entry alone. That is painting, and painting is a separate piece of work.
