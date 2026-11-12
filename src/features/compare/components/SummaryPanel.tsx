/**
 * Feature 08 — Compare: the change-summary panel (§8.2).
 *
 * Shows the total change count (ALWAYS the full total — filters never reduce
 * it, per ADR-0008), per-category counts, and dedicated moved + spelling-variant
 * lines. Renumbering cascades render as collapsed one-liners.
 */
import type { ChangeClass, ChangeSummary } from '../model.js';

interface SummaryPanelProps {
  summary: ChangeSummary;
}

const CLASS_LABELS: Record<ChangeClass, string> = {
  structural: 'Structural',
  obligationLanguage: 'Obligation language',
  numericalValue: 'Numerical value',
  definedTerm: 'Defined-term',
  crossRefRenumber: 'Cross-reference renumbering',
  substantiveText: 'Substantive text',
  whitespacePunct: 'Whitespace / punctuation',
  formatting: 'Formatting-only',
};

const CLASS_ORDER: ChangeClass[] = [
  'structural',
  'obligationLanguage',
  'numericalValue',
  'definedTerm',
  'crossRefRenumber',
  'substantiveText',
  'whitespacePunct',
  'formatting',
];

export function SummaryPanel({ summary }: SummaryPanelProps) {
  return (
    <section class="dp-compare__summary" aria-label="Change summary">
      {/* Total is ALWAYS the full count — never reduced by a filter (ADR-0008). */}
      <h3 class="dp-compare__summary-total">
        Total changes: <span class="dp-compare__summary-total-count">{summary.total}</span>
      </h3>

      <ul class="dp-compare__summary-categories">
        {CLASS_ORDER.map(cls => {
          const count = summary.byCategory[cls] ?? 0;
          return (
            <li
              key={cls}
              class={
                'dp-compare__summary-category' +
                (count === 0 ? ' dp-compare__summary-category--zero' : '')
              }
            >
              <span class="dp-compare__summary-category-label">{CLASS_LABELS[cls]}</span>
              <span class="dp-compare__summary-category-count">{count}</span>
            </li>
          );
        })}
      </ul>

      <ul class="dp-compare__summary-special">
        <li class="dp-compare__summary-moved">
          Moved paragraphs: <strong>{summary.movedParagraphs}</strong>
        </li>
        <li class="dp-compare__summary-spelling">
          en-GB/en-US spelling changes: <strong>{summary.spellingVariantCount}</strong>
        </li>
      </ul>

      {/*
        Parts diffed outside the body (§9.7–§9.9). Rendered only when there is
        something to say: at the default config headers and footers are not
        processed at all, so a permanent "0 header changes" line would advertise
        a comparison that did not happen.
      */}
      {(summary.partChanges.headers > 0 ||
        summary.partChanges.footers > 0 ||
        summary.partChanges.footnotes > 0 ||
        summary.styleDefinitionNote !== null) && (
        <div class="dp-compare__summary-parts">
          <h4 class="dp-compare__summary-parts-title">Other parts</h4>
          {/*
            These counts are a BREAKDOWN of changes already in the total above —
            `mergePartsIntoSummary` adds every part op to `total` and
            `byCategory` — not a separate tally to add to it. Said out loud
            because the block sits under the category list, which is exactly
            where a reader would expect an addendum. Spec §8.2.

            Gated on there being a COUNT, because the fourth line of this block
            is not one. `driver.ts` keeps `diffStyles`' note and discards its
            op, so a `styles.xml` change is reported here and is NOT in `total`
            (ADR-0042 § Open — §9.9's counted/filterable half is unmet). Printed
            unconditionally, this caption said "included in the total above"
            over a panel reading `Total changes: 0`, which is the surface
            asserting what the code does not do — the defect this caption was
            added to fix, one line along.
          */}
          {(summary.partChanges.headers > 0 ||
            summary.partChanges.footers > 0 ||
            summary.partChanges.footnotes > 0) && (
            <p class="dp-compare__summary-parts-note">Included in the total above.</p>
          )}
          <ul>
            {summary.partChanges.headers > 0 && (
              <li class="dp-compare__summary-part">
                Header changes: <strong>{summary.partChanges.headers}</strong>
              </li>
            )}
            {summary.partChanges.footers > 0 && (
              <li class="dp-compare__summary-part">
                Footer changes: <strong>{summary.partChanges.footers}</strong>
              </li>
            )}
            {summary.partChanges.footnotes > 0 && (
              <li class="dp-compare__summary-part">
                Footnote changes: <strong>{summary.partChanges.footnotes}</strong>
              </li>
            )}
            {summary.styleDefinitionNote !== null && (
              <li class="dp-compare__summary-part dp-compare__summary-part--styles">
                {summary.styleDefinitionNote}{' '}
                {/*
                  Said on the row itself, not left to the caption above, because
                  this row can be the ONLY one in the block. A style definition
                  is not at a paragraph, so its op has no range to navigate to
                  and is not counted (§9.9, ADR-0042 § Open). Nothing is hidden —
                  the reader always sees this line — but it is not in the total.
                */}
                <span class="dp-compare__summary-part-uncounted">— not counted above</span>
              </li>
            )}
          </ul>
        </div>
      )}

      {summary.renumberingCascades.length > 0 && (
        <div class="dp-compare__summary-cascades">
          <h4 class="dp-compare__summary-cascades-title">Renumbering cascades</h4>
          <ul>
            {summary.renumberingCascades.map(c => {
              const sign = c.shift > 0 ? `+${c.shift}` : `${c.shift}`;
              return (
                <li key={c.id} class="dp-compare__summary-cascade">
                  {/*
                    Two sentences, because they are two facts. A cross-reference
                    cascade is the document POINTING at a clause whose number
                    moved; a numbering cascade is the clause numbers themselves
                    moving (ADR-0042). Saying "references renumbered" of the
                    second would claim a reference the document does not contain.
                  */}
                  {c.source === 'numbering'
                    ? `${c.tokenKind} shifted (${sign}): ${c.affectedCount}`
                    : `${c.tokenKind} references renumbered (${sign}): ${c.affectedCount}`}
                  {c.exampleRefs.length > 0 && (
                    <span class="dp-compare__summary-cascade-examples">
                      {' '}— example: {c.exampleRefs.join('; ')}
                    </span>
                  )}
                </li>
              );
            })}
          </ul>
        </div>
      )}
    </section>
  );
}
