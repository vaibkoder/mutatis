/**
 * apps/web — feature-grouped findings view (ADR-0022, spec/features/12 §5 phase 3b).
 *
 * Renders the standalone Proof results: cross-reference, manually-typed clause
 * numbers (the "numbering detect" / static-numbering flag), defined terms, and
 * factual consistency — grouped exactly as the add-in's `FindingsList`, reusing
 * the shared category labels/ordering (`findings-grouping.ts`) and the host-neutral
 * `FindingRow`. It works off the pure `analyse*` returns (plain arrays + the two
 * banner-bearing analyses); there is no global findings store on web.
 *
 * Deliberately omits the add-in's Assign-source picker (a write/apply path → export
 * edit, deferred) and the findings filter box (add-in nicety). No `Office.*`/`Word.*`
 * token — stays out of the host bundle.
 */
import type { Finding } from '@shared/findings/index.js';
import type { DefinedTermsAnalysis } from '@features/defined-terms/index.js';
import type { CrossReferenceAnalysis } from '@features/cross-reference/index.js';
import { FindingRow } from '@taskpane/components/FindingRow.js';
import { IconChecklist } from './icons.js';
import {
  categoryLabel,
  groupByCategory,
  splitFindingsByFeature,
} from '@taskpane/components/findings-grouping.js';

interface WebFindingsProps {
  findings: ReadonlyArray<Finding>;
  /** Defined-terms analysis — drives the no-definitions / cross-ref banners. */
  defined: DefinedTermsAnalysis | null;
  /** Cross-reference analysis — drives the REF-field-unavailable banner. */
  xref: CrossReferenceAnalysis | null;
}

function CountBadge({ n }: { n: number }) {
  return (
    <span class="dp-findings-group__count" aria-label={`${n} issue${n === 1 ? '' : 's'}`}>
      {n}
    </span>
  );
}

function FindingItems({ findings }: { findings: ReadonlyArray<Finding> }) {
  return (
    <>
      {findings.map(f => (
        <li key={f.id}>
          <FindingRow finding={f} />
        </li>
      ))}
    </>
  );
}

export function WebFindings({ findings, defined, xref }: WebFindingsProps) {
  const grouped = groupByCategory(findings);
  const { xrefOrdered, staticNum, definedTermsOrdered, other } = splitFindingsByFeature(grouped);

  const total = findings.length;

  return (
    <div class="dp-web-findings">
      {/* Same section-header idiom as the document pane: icon, small uppercase
          label, count chip, 1px rule. The icon repeats the label, so it is
          decorative; the count keeps its own accessible name. */}
      <div class="dp-web__pane-head">
        <IconChecklist />
        <h2 class="dp-web__findings-heading">
          Findings
          <span
            class="dp-web__findings-count"
            aria-label={`${total} finding${total === 1 ? '' : 's'}`}
          >
            {total}
          </span>
        </h2>
      </div>

      <div class="dp-web__pane-body">
      {/* Defined-terms banner — mirrors the add-in (FindingsList §banner). */}
      {defined && (defined.noDefinitions || defined.crossRefCount > 0 || defined.crossDocCount > 0) && (
        <div class="dp-defined-terms-banner" role="status">
          {defined.noDefinitions && (
            <p class="dp-defined-terms-banner__empty">No definitions detected.</p>
          )}
          {defined.crossRefCount > 0 && (
            <p class="dp-defined-terms-banner__crossref">
              {defined.crossRefCount} term(s) defined by cross-reference within this document (treated as defined).
            </p>
          )}
          {defined.crossDocCount > 0 && (
            <p class="dp-defined-terms-banner__crossdoc">
              {defined.crossDocCount} term(s) defined by cross-reference to external documents (not checked).
            </p>
          )}
        </div>
      )}

      {/* REF-field detection is never available off-host — always degraded on web. */}
      {xref?.fieldsApiUnavailable && (
        <div class="dp-capability-banner" role="status">
          REF-field detection is not available outside Word — all cross-references are shown as manually typed.
        </div>
      )}

      {total === 0 && (
        <p class="dp-web__status" role="status">No issues found.</p>
      )}

      {/* Feature A — Cross-references. MUST render before Feature B. */}
      {xrefOrdered.length > 0 && (
        <section class="dp-feature-group" aria-label="Cross-references">
          <h3 class="dp-feature-group__heading">Cross-references</h3>
          {xrefOrdered.map(({ cat, findings: f }) => (
            <section key={cat} class="dp-findings-group">
              <h4 class="dp-findings-group__heading">
                {categoryLabel(cat)}
                <CountBadge n={f.length} />
              </h4>
              <ul class="dp-findings-group__list" aria-label={categoryLabel(cat)}>
                <FindingItems findings={f} />
              </ul>
            </section>
          ))}
        </section>
      )}

      {/* Feature B — Manually-typed clause numbers (the static-numbering flag). */}
      {staticNum.length > 0 && (
        <section class="dp-feature-group" aria-label="Manually-typed clause numbers">
          <h3 class="dp-feature-group__heading">Manually-typed clause numbers</h3>
          <ul class="dp-findings-group__list" aria-label="Manually-typed clause numbers">
            <FindingItems findings={staticNum} />
          </ul>
        </section>
      )}

      {/* Feature 01 — Defined terms. */}
      {definedTermsOrdered.length > 0 && (
        <section class="dp-feature-group" aria-label="Defined terms">
          <h3 class="dp-feature-group__heading">Defined terms</h3>
          {definedTermsOrdered.map(({ cat, findings: f }) => (
            <section key={cat} class="dp-findings-group">
              <h4 class="dp-findings-group__heading">
                {categoryLabel(cat)}
                <CountBadge n={f.length} />
              </h4>
              <ul class="dp-findings-group__list" aria-label={categoryLabel(cat)}>
                <FindingItems findings={f} />
              </ul>
            </section>
          ))}
        </section>
      )}

      {/* Other findings (factual consistency). */}
      {other.map(({ cat, findings: f }) => (
        <section key={cat} class="dp-findings-group">
          <h4 class="dp-findings-group__heading">
            {categoryLabel(cat)}
            <CountBadge n={f.length} />
          </h4>
          <ul class="dp-findings-group__list" aria-label={categoryLabel(cat)}>
            <FindingItems findings={f} />
          </ul>
        </section>
      ))}
      </div>
    </div>
  );
}
