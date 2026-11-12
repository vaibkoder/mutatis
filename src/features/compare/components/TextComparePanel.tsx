/**
 * Feature 12 (Compare — additional sources) T4: the Text-vs-text panel (ADR-0018).
 *
 * Two plain-text areas (Original / Modified) → an eagerly-generated tracked-changes
 * redline that the user opens on demand. Reuses the feature-08 SummaryPanel,
 * FilterToggles and ChangeList unchanged.
 *
 * ADR-0008: the filter toggles affect ONLY the navigation list. The summary total
 * and the generated redline carry every change — nothing is suppressed.
 */
import { useState } from 'preact/hooks';
import type { TextCompareDriverState } from '../ui-types.js';
import type { DiffOp } from '../model.js';
import { SummaryPanel } from './SummaryPanel.js';
import { FilterToggles } from './FilterToggles.js';
import { ChangeList, cascadeLabelMap } from './ChangeList.js';
import { RedlinePreview } from './RedlinePreview.js';
import { CopyButton } from '@features/ai-export/index.js';
import { buildChangeSet } from '../change-set-export.js';

interface TextComparePanelProps {
  driver: TextCompareDriverState;
}

export function TextComparePanel({ driver }: TextComparePanelProps) {
  const status = driver.status.value;
  const running = status === 'running';
  const done = status === 'done';
  const ops: DiffOp[] = driver.ops.value ?? [];
  const summary = driver.summary.value;
  // A run is allowed whenever we are not already running — the driver validates
  // empty/oversize input and surfaces a friendly message (TT6/TT7).
  const canRun = !running;

  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  /**
   * Export the read-only rendered-redline SNAPSHOT (ADR-0020, feature 16): render
   * the surfaced render source (the FULL change set — never the filtered list,
   * ADR-0008) to a programmatic vector PDF via RenderCore and download it. The
   * editable .docx ("Open redline" above) stays the PRIMARY artefact and is
   * untouched. Layout runs off the main thread (worker); the render path is
   * lazy-imported so it is code-split out of the main bundle (D1).
   */
  async function onExportSnapshot(): Promise<void> {
    const src = driver.renderSource.value;
    if (!src || exporting) return;
    setExporting(true);
    setExportError(null);
    try {
      const { renderRedlinePdf, downloadPdf, redlinePdfFilename } = await import('../render/index.js');
      const bytes = await renderRedlinePdf(src);
      downloadPdf(bytes, redlinePdfFilename(new Date()));
    } catch (err: unknown) {
      setExportError(err instanceof Error ? err.message : String(err));
    } finally {
      setExporting(false);
    }
  }

  return (
    <div class="dp-compare__text" lang="en-GB">
      <p class="dp-compare__text-intro">
        Paste two versions of a passage as plain text. Comparing builds a tracked-changes
        Word document; it opens only when you choose to open it — nothing is sent anywhere.
      </p>

      {/* ── Two text areas (Original / Modified) ───────────────────────── */}
      <div class="dp-compare__text-areas">
        <label class="dp-compare__text-field">
          <span class="dp-compare__text-label">Original (base)</span>
          <textarea
            class="dp-compare__text-input"
            rows={8}
            value={driver.originalText.value}
            disabled={running}
            onInput={e => {
              driver.originalText.value = (e.target as HTMLTextAreaElement).value;
            }}
          />
        </label>
        <label class="dp-compare__text-field">
          <span class="dp-compare__text-label">Modified (compared)</span>
          <textarea
            class="dp-compare__text-input"
            rows={8}
            value={driver.revisedText.value}
            disabled={running}
            onInput={e => {
              driver.revisedText.value = (e.target as HTMLTextAreaElement).value;
            }}
          />
        </label>
      </div>

      {/* ── Run / Cancel ───────────────────────────────────────────────── */}
      <div class="dp-compare__run-controls">
        <button
          type="button"
          class="dp-compare__run-btn"
          disabled={!canRun}
          aria-busy={running ? 'true' : 'false'}
          onClick={() => driver.runCompare()}
        >
          {running ? `Comparing… ${driver.progress.value}%` : 'Compare text'}
        </button>
        {running && (
          <button
            type="button"
            class="dp-compare__cancel-btn"
            onClick={() => driver.cancelCompare()}
          >
            Cancel
          </button>
        )}
      </div>

      {/* ── Error / message ────────────────────────────────────────────── */}
      {status === 'error' && driver.errorMessage.value && (
        <p class="dp-compare__error" role="alert">
          {driver.errorMessage.value}
        </p>
      )}

      {/* ── Results: open-redline + summary + filters + change list ─────── */}
      {done && summary && (
        <div class="dp-compare__results">
          <div class="dp-compare__text-redline">
            <button
              type="button"
              class="dp-compare__open-redline-btn"
              disabled={!driver.redlineReady.value}
              onClick={() => driver.openRedline()}
            >
              Open redline (.docx)
            </button>
            <p class="dp-compare__text-redline-note">
              The redline has been generated and is held in memory. It opens as a new
              tracked-changes document only when you click — with native accept/reject.
            </p>
            {/* Read-only rendered-redline snapshot (ADR-0020) — additional to the
                editable .docx above, never a replacement. Full change set,
                rendered as a vector PDF by RenderCore (feature 16). */}
            <div class="dp-compare__export">
              <button
                type="button"
                class="dp-compare__export-btn"
                disabled={!driver.renderSource.value || exporting}
                aria-busy={exporting ? 'true' : 'false'}
                onClick={() => void onExportSnapshot()}
              >
                {exporting ? 'Rendering PDF…' : 'Export rendered redline (snapshot)'}
              </button>
              {exportError && (
                <p class="dp-compare__error" role="alert">
                  {exportError}
                </p>
              )}
            </div>
          </div>

          {/* Read-only in-pane preview of the same redline — RenderCore's
              paginated layout (feature 16, upgrade of the feature-14 inline
              preview), painted from the same render source. The .docx above stays
              the editable artefact; nothing is suppressed (ADR-0008). */}
          <RedlinePreview src={driver.renderSource.value} />

          {/* Optional detail — at-a-glance breakdown, filters, and full change
              list collapsed by default to keep the common case uncluttered. The
              count stays visible; nothing is suppressed (ADR-0008). */}
          <details class="dp-compare__details">
            <summary class="dp-compare__details-summary">
              {summary.total} {summary.total === 1 ? 'change' : 'changes'} — breakdown, change list &amp; AI export
            </summary>
            <div class="dp-compare__details-body">
              <SummaryPanel summary={summary} />
              {/* Copy-for-AI: full change set (before/after, never markup — feature 11,
                  ADR-0021). Mirrors the Word-vs-Word panel; reuses the shared builder. */}
              <CopyButton
                build={() => buildChangeSet(ops, driver.originalParas.value, driver.revisedParas.value)}
                fileLabel="change-set"
              />
              <FilterToggles driver={driver} />
              <ChangeList
                ops={ops}
                filters={{
                  hideFormattingOnly: driver.hideFormattingOnly.value,
                  collapseRenumbering: driver.collapseRenumbering.value,
                }}
                cascadeLabels={cascadeLabelMap(summary.renumberingCascades)}
                onGoToChange={id => driver.goToChange(id)}
              />
            </div>
          </details>
        </div>
      )}
    </div>
  );
}
