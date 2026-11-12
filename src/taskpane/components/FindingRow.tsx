import { useState } from 'preact/hooks';
import type { Finding } from '@shared/findings/index.js';
import { getActiveHost, hasActiveHost } from '@shared/host/index.js';

interface FindingRowProps {
  finding: Finding;
}

const SEVERITY_ICONS: Record<string, string> = {
  error: '✕',
  warning: '⚠',
  info: 'ℹ',
};

const SEVERITY_LABELS: Record<string, string> = {
  error: 'Critical',
  warning: 'Warning',
  info: 'Info',
};

// Categories that use amber "Check" styling instead of the standard info blue.
const AMBIGUOUS_CATEGORIES = new Set([
  'factual-consistency/ambiguous-currency',
  'factual-consistency/century-ambiguous',
]);

export function FindingRow({ finding }: FindingRowProps) {
  const [showSource, setShowSource] = useState(false);
  const [occIdx, setOccIdx] = useState(0);

  const occs = finding.occurrences ?? [];
  const hasMulti = occs.length > 1;

  async function navigateTo(i: number) {
    const target = occs[i] ?? finding.location;
    // Navigation goes through the active host (ADR-0022): the office adapter
    // re-resolves and selects the live Range (the deferred-search + double-sync
    // sequence now lives in OfficeHostPort.reveal); a web host scrolls/highlights
    // in its own view. Best-effort: no host yet, or a stale location, is a no-op.
    if (!hasActiveHost()) return;
    try {
      await getActiveHost().reveal(target);
    } catch {
      // Navigation is best-effort; never let it surface as an unhandled rejection.
    }
  }

  async function handleNavigate() {
    await navigateTo(hasMulti ? occIdx : 0);
  }

  function handleNext(e: MouseEvent) {
    e.stopPropagation();
    const next = (occIdx + 1) % occs.length;
    setOccIdx(next);
    void navigateTo(next);
  }

  async function handleApply(e: MouseEvent) {
    e.stopPropagation();
    if (!finding.suggestedAction) return;
    try {
      // Lazy-load the Office bridge so the apply (write) path tree-splits OUT of
      // the eager bundle. FindingRow is reused by the standalone web app's
      // WebFindings, so a static `wordRun` import would drag Office JS (Word.run)
      // into the web chunk and break the confidentiality invariant (ADR-0001/0022).
      // The web app never surfaces apply; the add-in pays only a one-off dynamic
      // import on first click. Mirrors the cross-reference fields.js lazy-load.
      const { wordRun } = await import('@shared/office/index.js');
      await wordRun(async ctx => {
        await finding.suggestedAction!.apply(ctx);
        await ctx.sync();
      });
    } catch {
      // Fix failed; leave the finding in place.
    }
  }

  function handleSourceToggle(e: MouseEvent) {
    e.stopPropagation();
    setShowSource(v => !v);
  }

  const icon = SEVERITY_ICONS[finding.severity] ?? 'ℹ';
  const sevLabel = SEVERITY_LABELS[finding.severity] ?? finding.severity;
  const isAmbiguous = AMBIGUOUS_CATEGORIES.has(finding.category);
  const rowClass = [
    'dp-finding-row',
    isAmbiguous ? 'dp-finding-row--check' : `dp-finding-row--${finding.severity}`,
    showSource ? 'dp-finding-row--expanded' : '',
  ].filter(Boolean).join(' ');

  const locationEl = finding.sourceSnippet ? (
    <button
      class="dp-finding-row__location dp-finding-row__location--src"
      onClick={handleSourceToggle}
      type="button"
      aria-expanded={showSource ? 'true' : 'false'}
    >
      {finding.location.label}
    </button>
  ) : (
    <span class="dp-finding-row__location">{finding.location.label}</span>
  );

  return (
    <>
      <button
        class={rowClass}
        onClick={handleNavigate}
        aria-label={`${sevLabel}: ${finding.title} at ${finding.location.label}`}
        type="button"
      >
        <span class="dp-finding-row__icon" aria-hidden="true">{icon}</span>
        <span class="dp-finding-row__body">
          <span class="dp-finding-row__title">{finding.title}</span>
          {finding.detail && (
            <span class="dp-finding-row__detail">{finding.detail}</span>
          )}
          {locationEl}
          {finding.occurrences && (
            <span class="dp-finding-row__count" aria-label={`used ${finding.occurrences.length} times`}>
              {finding.occurrences.length}×
            </span>
          )}
        </span>
        {finding.suggestedAction && (
          <button
            class="dp-finding-row__action"
            onClick={handleApply}
            aria-label={finding.suggestedAction.label}
            type="button"
          >
            {finding.suggestedAction.label}
          </button>
        )}
      </button>
      {hasMulti && (
        <button
          class="dp-finding-row__next"
          onClick={handleNext}
          aria-label={`Go to next occurrence (${occIdx + 1} of ${occs.length})`}
          type="button"
        >
          Next ▸ <span class="dp-finding-row__next-pos">{occIdx + 1}/{occs.length}</span>
        </button>
      )}
      {finding.sourceSnippet && showSource && (
        <div class="dp-source-snippet" role="region" aria-label={`Definition of ${finding.sourceSnippet.label}`}>
          <p class="dp-source-snippet__text">{finding.sourceSnippet.text}</p>
        </div>
      )}
    </>
  );
}
