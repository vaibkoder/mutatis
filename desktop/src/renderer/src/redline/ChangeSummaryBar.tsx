import type { ChangeSummary } from '@features/compare/model.js';

/**
 * The count strip above a redline.
 *
 * Every category is shown, including zeroes suppressed only when the whole
 * category is empty — this is a summary OF the change set, never a filter ON it.
 * The document below always shows everything.
 */

const LABELS: Record<string, string> = {
  structural: 'Structural',
  obligationLanguage: 'Obligation',
  numericalValue: 'Numeric',
  definedTerm: 'Defined term',
  crossRefRenumber: 'Cross-ref',
  substantiveText: 'Substantive',
  whitespacePunct: 'Whitespace',
  formatting: 'Formatting',
};

export function ChangeSummaryBar({
  summary,
  quarantinedOriginal,
  quarantinedRevised,
}: {
  summary: ChangeSummary;
  quarantinedOriginal: number;
  quarantinedRevised: number;
}) {
  const categories = Object.entries(summary.byCategory).filter(([, n]) => n > 0);
  const quarantined = quarantinedOriginal + quarantinedRevised;

  const { headers, footers, footnotes } = summary.partChanges;
  const outsideBody = headers + footers + footnotes;
  const outsideBodyParts = [
    headers > 0 ? `${headers} header` : null,
    footers > 0 ? `${footers} footer` : null,
    footnotes > 0 ? `${footnotes} footnote` : null,
  ].filter((part): part is string => part !== null);

  return (
    <div className="flex h-9 shrink-0 flex-wrap items-center gap-x-3 gap-y-1 border-b border-border bg-surface-raised px-4 text-[11px]">
      <span className="font-semibold text-text">
        {summary.total} {summary.total === 1 ? 'change' : 'changes'}
      </span>

      {categories.map(([key, count]) => (
        <span key={key} className="text-text-muted">
          {LABELS[key] ?? key} <span className="tabular-nums text-text">{count}</span>
        </span>
      ))}

      {summary.movedParagraphs > 0 ? (
        <span className="text-[var(--redline-move)]">
          {summary.movedParagraphs} moved
        </span>
      ) : null}

      {summary.renumberingCascades.length > 0 ? (
        <span className="text-text-muted">
          {summary.renumberingCascades.length} renumbering{' '}
          {summary.renumberingCascades.length === 1 ? 'cascade' : 'cascades'}
        </span>
      ) : null}

      {summary.spellingVariantCount > 0 ? (
        <span className="text-text-muted">
          {summary.spellingVariantCount} spelling{' '}
          {summary.spellingVariantCount === 1 ? 'variant' : 'variants'}
        </span>
      ) : null}

      {/* Changes OUTSIDE the body. The engine counts these into `total` but
          deliberately keeps them out of the navigable change list, because a
          header op's range indexes a different part and jumping to it would
          land on the wrong paragraph. So the count above can exceed what the
          document below can show you, and headers and footers are not drawn
          there at all — the renderer takes the body and the footnotes. Saying
          nothing would leave that arithmetic looking like a bug, or worse, let
          a real change go unnoticed. */}
      {outsideBody > 0 ? (
        <span
          className="text-text-muted"
          title={
            'Counted in the total above. These live outside the document body, so they are ' +
            'not in the change list and cannot be navigated to; header and footer changes ' +
            'are not drawn in the redline below.'
          }
        >
          {outsideBodyParts.join(', ')} outside the body
        </span>
      ) : null}

      {/* A change to styles.xml itself, not to any one paragraph — so it can
          appear nowhere in the document below. Upstream builds it as a note
          meant for exactly this strip. */}
      {summary.styleDefinitionNote !== null ? (
        <span className="text-text-muted" title={summary.styleDefinitionNote}>
          style definitions changed
        </span>
      ) : null}

      {/* Quarantined comments are a DISCLOSURE, not a detail. They are content
          the engine deliberately set aside, and the user has to know it existed. */}
      {quarantined > 0 ? (
        <span
          className="ml-auto text-text-faint"
          title="Pre-existing tracked changes and comments were set aside before diffing so they could not be mistaken for this comparison's own changes."
        >
          {quarantined} pre-existing revision{quarantined === 1 ? '' : 's'} set aside
        </span>
      ) : null}
    </div>
  );
}
