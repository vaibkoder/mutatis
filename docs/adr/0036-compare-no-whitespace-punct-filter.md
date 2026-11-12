# ADR-0036: Compare — remove the "hide whitespace/punctuation" filter rather than repair it

- **Status:** Accepted
- **Date:** 2026-09-04
- **Relates to:** ADR-0008 (Compare categorises, never suppresses); ADR-0002 (SG + E&W jurisdiction); feature 08 (`spec/features/08-compare.md` §6, §8.2)
- **Workstream:** Tier-3 fix on the redline-readability board, approved by the reviewer 2026-09-04

## Context

The Compare panel shipped three display filters. One of them —
*Hide whitespace/punctuation*, default ON — had never hidden anything.

`ChangeList.applyFilters` read the **primary** classification:

```ts
if (flags.hideWhitespacePunct && op.classification === 'whitespacePunct') return false;
```

That comparison can never be true, and the proof is two lines of `taxonomy.ts`:

- `isWhitespacePunct` returns `false` unless `hasTextDelta(op)`;
- `isSubstantiveText(op)` **is** `hasTextDelta(op)`.

So matching `whitespacePunct` (§6 priority 7) implies matching `substantiveText`
(priority 6), which always wins the priority resolution. `substantiveText`'s
implementation comment says "any **remaining** run-text insert/delete/replace"
— it was meant as a fall-through, but as written it is a catch-all sitting one
rank above the class it was supposed to fall through to.

Measured on the real pipeline, not inferred:

| edit | op emitted | `classification` | `classes` |
| --- | --- | --- | --- |
| comma removed | replace | `substantiveText` | `substantiveText`, `whitespacePunct` |
| semicolon → comma | replace | `substantiveText` | `substantiveText`, `whitespacePunct` |
| double space | **none** | — | — |
| curly → straight quote | **none** | — | — |

The last two rows also correct a claim made earlier in this workstream — that
whitespace and curly-quote edits are reported as `formatting` and hidden by the
working toggle. They are not reported **at all**: §5.0 normalisation absorbs
them before the diff, so the paragraphs compare equal and no op exists.

## Decision

**Remove the toggle** — the signal, the checkbox, the `FilterFlags` field, the
`filtersActive` envelope field, and the filter line itself. Compare now has two
display filters (*Hide formatting-only*, *Collapse renumbering cascades*), and
the spec says so.

**Rewiring the filter to read `op.classes` was considered and rejected.** It is
the obvious "fix", and it is the wrong one twice over:

1. An op that matches `whitespacePunct` **also** carries `substantiveText` —
   always, by the implication above. Hiding on the secondary class would
   therefore hide an op the classifier has called substantive. That is the
   spirit of ADR-0008 pointed the wrong way: a filter is navigation, and a
   navigation aid that removes substantive changes from the only list a reviewer
   scans is a suppression mechanism in all but name.
2. Punctuation is substantive in the drafting this product serves (ADR-0002).
   A comma governs the scope of a proviso; the class of dispute that turns on
   one is ordinary, not exotic. A default-ON filter that hid comma changes would
   be a defect with a much worse failure mode than the dead toggle it replaced.

The `whitespacePunct` **class** stays. It is real, it is recorded in
`op.classes`, and it is available to reporting. What is removed is the one
consumer that read it from the wrong field.

## Consequences

- **The exported report envelope changes shape.** `ChangeSummary.filtersActive`
  loses a field. The export builder already ignored every filter (the ADR-0008
  wall in `report-model.ts`), so no exported *content* moves — only the record
  of which filters were on.
- **The `byCategory` Whitespace/punct bucket is permanently `0`,** and this ADR
  does not change that. The counts are keyed on `classification` precisely so
  they partition `total`; counting `classes` instead would make the buckets sum
  to more than the total and would move a published figure. The in-document
  summary annotation omits zero buckets, so it is unaffected; the on-screen
  SummaryPanel renders all eight categories, so it shows a permanently empty
  Whitespace/punct row. **Left as-is deliberately and surfaced to the reviewer
  as a separate decision** rather than folded into an approved fix.
- **Two further surfaces are keyed on `classification` and so are unreachable
  today.** `CLASS_LABEL`/`CLASS_LABELS`' `WS` tag, and the
  `.dp-compare__change-row--whitespacePunct` CSS selector. The label maps are
  typed `Record<ChangeClass, string>`, so their entries are compile-enforced and
  must stay; the CSS selector is genuinely removable but shares a declaration
  with `--formatting` and is now commented to say why it is there, because
  whether it stays dead depends on the open `byCategory` decision above.
- **Two test fixtures were asserting a shape the classifier cannot produce.**
  Both hand-built `classification: 'whitespacePunct'` ops — one in
  `compare.ui.panel.test.tsx` to prove the filter hid them, one in
  `export.report-model.test.ts` to prove the export kept them. The first was
  pinning fiction (which is how the dead filter kept looking covered); the
  second was asserting a `ReportRow` field that cannot hold that value. Both now
  use the real shape — substantive primary, `whitespacePunct` secondary — and
  the panel test asserts the correct behaviour: **a punctuation change stays
  visible.**
- One further latent defect surfaced while repairing those: the report-model
  wall test compared `applyFilters(...).length` against `report.rows.length`,
  but `applyFilters` sees the `equal` op and `rows` never does. It held only
  because two ops were being filtered rather than one. Now compared like with
  like.

## Evidence

`compare.classify.taxonomy.test.ts` gains two tests that pin the unreachability
directly, so the toggle cannot be re-added on the assumption that it works:

- the punctuation delta records the class but is classified `substantiveText`;
- no op that matches `whitespacePunct` can be classified as it, across a
  punctuation swap, a punctuation deletion and a punctuation insertion, each
  given the class's own most-favourable precondition (aggressive normalisation
  identical on both sides).

Ablation: reordering `PRIORITY_ORDER` so `whitespacePunct` outranks
`substantiveText` fails both, and two unrelated priority tests besides —
restored, all 17 pass.

Full suite 1758 → **1760 tests + 1 todo**, green; lint and `tsc --noEmit` clean.
