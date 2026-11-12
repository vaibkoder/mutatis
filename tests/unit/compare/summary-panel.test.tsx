/**
 * Feature 08 — the change-summary panel's two conditional blocks (§8.2).
 *
 * Both shipped unrendered by any test. They are the ONLY place two facts reach
 * the reader, so a silent regression in either is indistinguishable from the
 * change never having been found — which under ADR-0008 is the thing Compare
 * must never do:
 *
 *   - "Other parts" (§9.7–§9.9) — header/footer/footnote counts and the style
 *     DEFINITION note. Rendered only when there is something to say, because at
 *     the default config headers and footers are not processed at all and a
 *     permanent "0 header changes" line would advertise a comparison that did
 *     not happen.
 *   - The `source` sentence on a cascade (ADR-0042) — a cross-reference cascade
 *     is the document POINTING AT a clause whose number moved; a numbering
 *     cascade is the clause numbers themselves moving. Saying "references
 *     renumbered" of the second claims a reference the document does not contain.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { render } from 'preact';

import { SummaryPanel } from '@features/compare/components/SummaryPanel.js';
import type { CascadeAnnotation, ChangeSummary } from '@features/compare/model.js';

let host: HTMLDivElement | null = null;

afterEach(() => {
  if (host) render(null, host);
  host?.remove();
  host = null;
});

function mount(summary: ChangeSummary): string {
  host = document.createElement('div');
  document.body.appendChild(host);
  render(<SummaryPanel summary={summary} />, host);
  return host.textContent ?? '';
}

function summaryOf(over: Partial<ChangeSummary> = {}): ChangeSummary {
  return {
    total: 0,
    byCategory: {
      structural: 0, obligationLanguage: 0, numericalValue: 0, definedTerm: 0,
      crossRefRenumber: 0, substantiveText: 0, whitespacePunct: 0, formatting: 0,
    },
    renumberingCascades: [],
    movedParagraphs: 0,
    spellingVariantCount: 0,
    quarantinedCommentsOriginal: 0,
    quarantinedCommentsRevised: 0,
    partChanges: { headers: 0, footers: 0, footnotes: 0 },
    styleDefinitionNote: null,
    filtersActive: { hideFormattingOnly: false, collapseRenumbering: false },
    ...over,
  };
}

describe('the "Other parts" block (§9.7–§9.9)', () => {
  it('is absent entirely when nothing outside the body changed', () => {
    const text = mount(summaryOf());
    expect(host!.querySelector('.dp-compare__summary-parts')).toBeNull();
    // Not "Header changes: 0" — the default config does not diff headers at all.
    expect(text).not.toContain('Header changes');
  });

  it('reports each part that changed, and only those', () => {
    const text = mount(summaryOf({ partChanges: { headers: 2, footers: 0, footnotes: 5 } }));

    expect(host!.querySelector('.dp-compare__summary-parts')).not.toBeNull();
    expect(text).toContain('Header changes');
    expect(text).toContain('Footnote changes');
    expect(text).not.toContain('Footer changes');
    const items = [...host!.querySelectorAll('.dp-compare__summary-part')];
    expect(items.map((li) => li.textContent)).toEqual([
      'Header changes: 2',
      'Footnote changes: 5',
    ]);
  });

  it('says the counts are already in the total, because the block sits where an addendum would', () => {
    // Spec §8.2. `mergePartsIntoSummary` adds every part op to `total` AND to
    // `byCategory`, so this block is a BREAKDOWN, never a second tally — and it
    // renders directly under the category list, which is exactly where a reader
    // would expect something to add on. The caption was added in review to
    // close that, and went in with no test: a later edit deleted the line and the
    // whole compare suite stayed green, which is this plan's own recurring
    // proof that a surface was untested.
    const text = mount(summaryOf({ partChanges: { headers: 0, footers: 0, footnotes: 5 } }));

    expect(host!.querySelector('.dp-compare__summary-parts-note')).not.toBeNull();
    expect(text).toContain('Included in the total above.');
    // The counted rows carry no uncounted marker.
    expect(text).not.toContain('not counted above');
  });

  it('shows the style DEFINITION note on its own, with no part counts at all', () => {
    // §9.9 — styles.xml itself differs. Not gated by excludeHeaders, so this is
    // the case where the block must appear with every part count at zero.
    const text = mount(summaryOf({ styleDefinitionNote: '3 style definitions differ' }));

    expect(text).toContain('Other parts');
    expect(text).toContain('3 style definitions differ');
    expect(text).not.toContain('Header changes');

    // …and it is NOT in the headline total. `driver.ts` keeps `diffStyles`' note
    // and discards its op, so this row is reported and uncounted (§9.9's
    // counted half is unmet — ADR-0042 § Open). This test already mounted the
    // contradicting case for two passes and asserted neither half of it: with
    // `total` 0 on screen, the caption above read "Included in the total
    // above." Found in review.
    expect(text).not.toContain('Included in the total above');
    expect(text).toContain('not counted above');
  });
});

describe('a cascade says which KIND of renumbering it was (ADR-0042)', () => {
  const cascade = (over: Partial<CascadeAnnotation> = {}): CascadeAnnotation => ({
    id: 'c1',
    tokenKind: 'Clause',
    shift: 1,
    affectedCount: 12,
    exampleRefs: ['Clause 4.2 → 4.3'],
    source: 'crossReference',
    ...over,
  });

  it('says "references renumbered" for a CROSS-REFERENCE cascade', () => {
    const text = mount(summaryOf({ renumberingCascades: [cascade()] }));
    expect(text).toContain('Clause references renumbered (+1): 12');
    expect(text).toContain('example: Clause 4.2 → 4.3');
  });

  it('says "shifted" for a NUMBERING cascade — the clause numbers themselves moved', () => {
    const text = mount(
      summaryOf({
        renumberingCascades: [
          cascade({
            source: 'numbering',
            tokenKind: 'Clause numbering (level 1)',
            affectedCount: 7,
            exampleRefs: ['1.1 → 2.1'],
          }),
        ],
      }),
    );

    expect(text).toContain('Clause numbering (level 1) shifted (+1): 7');
    // The claim this sentence exists to avoid: these paragraphs contain no
    // reference to anything. Their own numbers moved.
    expect(text).not.toContain('references renumbered');
  });

  it('renders a negative shift with its sign', () => {
    const text = mount(summaryOf({ renumberingCascades: [cascade({ shift: -2 })] }));
    expect(text).toContain('(-2)');
  });
});
