/**
 * Feature 08 — Compare: the task-pane panel (C6). Spec §8.
 *
 * The panel accepts the driver as a prop for testability; it defaults to a stub
 * driver. Wave 2 / I1 will inject the real driver (a live `CompareDriverState`).
 *
 * ADR-0008: the filter toggles here affect ONLY the navigation list (see
 * ChangeList.applyFilters). The summary total and the authored redline document
 * are never altered by a filter.
 */
import { useState } from 'preact/hooks';
import type { ChangeSummary, DiffOp } from './model.js';
import {
  type CompareDriverState,
  type TextCompareDriverState,
  type PdfCompareDriverState,
  makeStubDriverState,
  makeStubTextDriverState,
  makeStubPdfDriverState,
} from './ui-types.js';
import { SlotPicker } from './components/SlotPicker.js';
import { FilterToggles } from './components/FilterToggles.js';
import { SummaryPanel } from './components/SummaryPanel.js';
import { ChangeList, cascadeLabelMap } from './components/ChangeList.js';
import { RedlinePreview } from './components/RedlinePreview.js';
import { TextComparePanel } from './components/TextComparePanel.js';
import { PdfComparePanel } from './components/PdfComparePanel.js';
import { CopyButton } from '@features/ai-export/index.js';
import { buildChangeSet } from './change-set-export.js';

interface ComparePanelProps {
  /** Optional: defaults to a stub. Wave 2 / I1 injects the real driver. */
  driver?: CompareDriverState;
  /** Text-vs-text driver (feature 12, ADR-0018). Defaults to a stub. */
  textDriver?: TextCompareDriverState;
  /** PDF-vs-PDF driver (feature 12). Defaults to a stub. */
  pdfDriver?: PdfCompareDriverState;
  /**
   * Web target (12f): hide "Use current document" in the Word-vs-Word slots —
   * the standalone web app has no active host document, so both sides upload.
   * The add-in omits this (defaults false), keeping its behaviour unchanged.
   */
  uploadOnly?: boolean;
}

/** Which comparison input mode is active (feature 12 adds Text-vs-text + PDF). */
type CompareMode = 'docx' | 'text' | 'pdf';

export function ComparePanel({
  driver = makeStubDriverState(),
  textDriver = makeStubTextDriverState(),
  pdfDriver = makeStubPdfDriverState(),
  uploadOnly = false,
}: ComparePanelProps) {
  const [mode, setMode] = useState<CompareMode>('docx');

  const status = driver.status.value;
  const running = status === 'running';
  const done = status === 'done';
  const ops: DiffOp[] = driver.ops.value ?? [];
  const summary = driver.summary.value;
  const canRun = driver.originalReady.value && driver.revisedReady.value && !running;

  return (
    <div class="dp-compare" lang="en-GB">
      {/* ── 0. Input-mode toggle (Word-vs-Word | Text-vs-text) ─────────── */}
      <div class="dp-compare__mode-toggle" role="tablist" aria-label="Comparison input mode">
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'docx' ? 'true' : 'false'}
          class={`dp-compare__mode-btn${mode === 'docx' ? ' dp-compare__mode-btn--active' : ''}`}
          onClick={() => setMode('docx')}
        >
          Word documents
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'text' ? 'true' : 'false'}
          class={`dp-compare__mode-btn${mode === 'text' ? ' dp-compare__mode-btn--active' : ''}`}
          onClick={() => setMode('text')}
        >
          Pasted text
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'pdf' ? 'true' : 'false'}
          class={`dp-compare__mode-btn${mode === 'pdf' ? ' dp-compare__mode-btn--active' : ''}`}
          onClick={() => setMode('pdf')}
        >
          PDFs
        </button>
      </div>

      {mode === 'text' && <TextComparePanel driver={textDriver} />}
      {mode === 'pdf' && <PdfComparePanel driver={pdfDriver} />}
      {mode === 'docx' && (
        <DocxCompareBody
          driver={driver}
          ops={ops}
          summary={summary}
          status={status}
          running={running}
          done={done}
          canRun={canRun}
          uploadOnly={uploadOnly}
        />
      )}
    </div>
  );
}

interface DocxCompareBodyProps {
  driver: CompareDriverState;
  ops: DiffOp[];
  summary: ChangeSummary | null;
  status: string;
  running: boolean;
  done: boolean;
  canRun: boolean;
  uploadOnly: boolean;
}

function DocxCompareBody({ driver, ops, summary, status, running, done, canRun, uploadOnly }: DocxCompareBodyProps) {
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  /**
   * Export the read-only rendered-redline SNAPSHOT (ADR-0020, feature 16): render
   * the surfaced render source (the FULL change set — never the filtered list,
   * ADR-0008) to a programmatic vector PDF via RenderCore and download it. The
   * host-doc tracked-changes redline stays PRIMARY and untouched — this is an
   * additional read-only artefact, computed off the main thread (worker). The
   * render path (engine + clone-font bytes + worker) is lazy-imported here so it
   * is code-split out of the main bundle — non-export flows pay nothing (D1).
   */
  async function onExportSnapshot(): Promise<void> {
    const src = driver.renderSource.value;
    if (!src || exporting) return;
    setExporting(true);
    setExportError(null);
    try {
      const { renderRedlinePdf, downloadPdf, redlinePdfFilename } = await import('./render/index.js');
      const bytes = await renderRedlinePdf(src);
      downloadPdf(bytes, redlinePdfFilename(new Date()));
    } catch (err: unknown) {
      setExportError(err instanceof Error ? err.message : String(err));
    } finally {
      setExporting(false);
    }
  }

  return (
    <>
      {/* ── 1. Two source slots with direction (§8.1) ──────────────────── */}
      <div class="dp-compare__slots">
        <SlotPicker
          side="original"
          label="Original (base)"
          driver={driver}
          ready={driver.originalReady.value}
          uploadOnly={uploadOnly}
        />
        <div class="dp-compare__direction" aria-hidden="true">
          →
        </div>
        <SlotPicker
          side="revised"
          label="Revised (compared)"
          driver={driver}
          ready={driver.revisedReady.value}
          uploadOnly={uploadOnly}
        />
      </div>

      {/* ── 2. Per-side tracked-change warnings (§8.1) ─────────────────── */}
      {driver.originalHasRevisions.value && (
        <FlattenNotice
          label="Original"
          value={driver.flattenOriginal.value}
          onChange={v => {
            driver.flattenOriginal.value = v;
          }}
        />
      )}
      {driver.revisedHasRevisions.value && (
        <FlattenNotice
          label="Revised"
          value={driver.flattenRevised.value}
          onChange={v => {
            driver.flattenRevised.value = v;
          }}
        />
      )}

      {/* ── 3. Run / Cancel controls (§8.1) ────────────────────────────── */}
      <div class="dp-compare__run-controls">
        <button
          type="button"
          class="dp-compare__run-btn"
          disabled={!canRun}
          aria-busy={running ? 'true' : 'false'}
          onClick={() => driver.runCompare()}
        >
          {running ? `Comparing… ${driver.progress.value}%` : 'Run Compare'}
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

      {/* ── 4. Error display ───────────────────────────────────────────── */}
      {status === 'error' && driver.errorMessage.value && (
        <p class="dp-compare__error" role="alert">
          {driver.errorMessage.value}
        </p>
      )}

      {/* ── 5. Results: the 90% path — preview + download up front; the
             summary / filters / change list / AI export are demoted into a
             single collapsed disclosure so the common case stays uncluttered.
             ADR-0008: nothing is suppressed — every change is still in the
             redline and one click away in the details. ─────────────────── */}
      {done && summary && (
        <div class="dp-compare__results">
          {/* Read-only in-pane preview of the rendered redline — RenderCore's
              paginated layout, painted from the same render source that backs
              the PDF download. This is the headline result. */}
          <RedlinePreview src={driver.renderSource.value} />

          {/* Primary download. Read-only rendered-redline snapshot (ADR-0020),
              full change set, vector PDF via RenderCore (feature 16). In the
              Word host the live tracked-changes redline stays PRIMARY; this is
              an additional artefact — on web it is the deliverable. */}
          <div class="dp-compare__export">
            <button
              type="button"
              class="dp-compare__export-btn dp-compare__export-btn--primary"
              disabled={!driver.renderSource.value || exporting}
              aria-busy={exporting ? 'true' : 'false'}
              onClick={() => void onExportSnapshot()}
            >
              {exporting ? 'Rendering PDF…' : 'Download redline (PDF)'}
            </button>
            {exportError && (
              <p class="dp-compare__error" role="alert">
                {exportError}
              </p>
            )}
          </div>

          {/* Everything below is optional detail (the at-a-glance breakdown,
              filters, full change list, and AI export). Collapsed by default —
              the count stays visible in the summary. */}
          <details class="dp-compare__details">
            <summary class="dp-compare__details-summary">
              {summary.total} {summary.total === 1 ? 'change' : 'changes'} — breakdown, change list &amp; AI export
            </summary>
            <div class="dp-compare__details-body">
              <SummaryPanel summary={summary} />
              {/* Copy-for-AI: full change set (before/after, never markup — feature 11). */}
              <CopyButton
                build={() => buildChangeSet(
                  ops, driver.originalParas.value, driver.revisedParas.value, driver.parts.value,
                )}
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
                locations={driver.parts.value}
                onGoToChange={id => driver.goToChange(id)}
              />
            </div>
          </details>
        </div>
      )}
    </>
  );
}

interface FlattenNoticeProps {
  label: string;
  value: 'acceptAll' | 'rejectAll';
  onChange: (v: 'acceptAll' | 'rejectAll') => void;
}

function FlattenNotice({ label, value, onChange }: FlattenNoticeProps) {
  return (
    <div class="dp-compare__flatten-notice" role="note">
      <em>{label} contains tracked changes — they will be accepted before comparison.</em>
      <label class="dp-compare__flatten-select">
        Flatten:
        <select
          value={value}
          onChange={e => onChange((e.target as HTMLSelectElement).value as 'acceptAll' | 'rejectAll')}
        >
          <option value="acceptAll">Accept all</option>
          <option value="rejectAll">Reject all</option>
        </select>
      </label>
    </div>
  );
}


// Module-level stub driver for the default export, so the panel renders
// standalone today. Wave 2 / I1 swaps this for the live driver from ./driver.js.
const stubDriver = makeStubDriverState();

/** Default export: ComparePanel bound to a (stub) driver. I1 injects the real one. */
export default function ComparePanelDefault() {
  return <ComparePanel driver={stubDriver} />;
}
