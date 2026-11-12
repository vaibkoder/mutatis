/**
 * PDF-vs-PDF compare panel.
 *
 * SAFETY, NOT POLISH. The dangerous failure mode this guards against is a lawyer
 * trusting a "no differences" result on an execution copy whose signature block,
 * seal or stamp was never visually checked. So the panel:
 *   - frames results as "candidate differences", never a redline;
 *   - shows a first-class, ALWAYS-VISIBLE "what is NOT compared" disclosure
 *     (images/signatures/seals/stamps; visual formatting; scanned pages); and
 *   - surfaces the per-run image-presence notice ("N pages contain image content
 *     not text-compared", D3) so blind spots are named, never silently passed.
 *
 * A scanned/image-only PDF surfaces the driver's NoExtractableTextError message
 * verbatim (P7) — an explicit "cannot compare", not a misleading empty result.
 * Reuses the shared SummaryPanel / ChangeList / SideBySideDiff via PdfFindings.
 */
import type { PdfCompareDriverState, PdfImageNotice } from '../ui-types.js';
import {
  PDF_NOT_COMPARED_TITLE,
  PDF_NOT_COMPARED_ITEMS,
  pdfNotComparedLines,
  imageNoticeLine,
  imageNoticeLines,
} from '../export/safety-copy.js';
import { buildCompareReport } from '../export/report-model.js';
import { exportReport } from '../export/print-view.js';
import { FilterToggles } from './FilterToggles.js';
import { PdfFindings } from './PdfFindings.js';

interface PdfComparePanelProps {
  driver: PdfCompareDriverState;
}

export function PdfComparePanel({ driver }: PdfComparePanelProps) {
  const status = driver.status.value;
  const running = status === 'running';
  const done = status === 'done';
  const ops = driver.ops.value ?? [];
  const summary = driver.summary.value;
  const canRun = driver.originalReady.value && driver.revisedReady.value && !running;

  /**
   * Build the candidate-differences report from the FULL `ops` (never the
   * filtered list — ADR-0008) + the shared disclosures + the per-run image
   * notice, then print it (A5). The generation date is formatted en-GB here in
   * the handler so the report-model stays pure.
   */
  function onExport(): void {
    if (!summary) return;
    const notice = driver.imageNotice.value;
    const report = buildCompareReport({
      mode: 'pdf',
      ops,
      summary,
      originalParas: driver.originalParas.value,
      revisedParas: driver.revisedParas.value,
      disclosures: pdfNotComparedLines(),
      ...(notice ? { imageNotice: imageNoticeLines(notice) } : {}),
      generatedAt: new Date().toLocaleDateString('en-GB', {
        day: 'numeric',
        month: 'long',
        year: 'numeric',
      }),
    });
    exportReport(report);
  }

  return (
    <div class="dp-compare__pdf" lang="en-GB">
      <p class="dp-compare__pdf-intro">
        Upload two PDFs to find <strong>candidate differences</strong> in their text. This is
        not a redline — PDFs have no tracked-changes target — and the comparison reads text
        only. Nothing is uploaded anywhere; the PDFs are read on your machine.
      </p>

      {/* ── First-class "what is NOT compared" disclosure (ALWAYS visible) ── */}
      <NotComparedDisclosure />

      {/* ── Two PDF upload slots ───────────────────────────────────────── */}
      <div class="dp-compare__pdf-slots">
        <PdfSlot
          side="original"
          label="Original (base)"
          ready={driver.originalReady.value}
          disabled={running}
          onPick={file => driver.uploadPdf('original', file)}
        />
        <div class="dp-compare__direction" aria-hidden="true">→</div>
        <PdfSlot
          side="revised"
          label="Revised (compared)"
          ready={driver.revisedReady.value}
          disabled={running}
          onPick={file => driver.uploadPdf('revised', file)}
        />
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
          {running ? `Comparing… ${driver.progress.value}%` : 'Compare PDFs'}
        </button>
        {running && (
          <button type="button" class="dp-compare__cancel-btn" onClick={() => driver.cancelCompare()}>
            Cancel
          </button>
        )}
      </div>

      {/* ── Error (scanned/image-only → clear unsupported message, P7) ──── */}
      {status === 'error' && driver.errorMessage.value && (
        <p class="dp-compare__error" role="alert">
          {driver.errorMessage.value}
        </p>
      )}

      {/* ── Results: image notice + summary + filters + findings ───────── */}
      {done && summary && (
        <div class="dp-compare__results">
          {driver.imageNotice.value && <ImageNotice notice={driver.imageNotice.value} />}
          <FilterToggles driver={driver} />
          <PdfFindings
            ops={ops}
            summary={summary}
            filters={{
              hideFormattingOnly: driver.hideFormattingOnly.value,
              collapseRenumbering: driver.collapseRenumbering.value,
            }}
            originalParas={driver.originalParas.value}
            revisedParas={driver.revisedParas.value}
          />
          {/* ── Export the FULL candidate-differences report (ADR-0008/0020) ── */}
          <div class="dp-compare__export">
            <button
              type="button"
              class="dp-compare__export-btn"
              onClick={onExport}
            >
              Export candidate-differences report
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

/**
 * The always-visible blind-spot disclosure. Safety-critical copy (D2), now
 * sourced from the shared safety-copy module (A3) so the panel and the exported
 * report can never drift.
 */
function NotComparedDisclosure() {
  return (
    <section class="dp-compare__pdf-not-compared" role="note" aria-label="What is not compared">
      <h3 class="dp-compare__pdf-not-compared-title">{PDF_NOT_COMPARED_TITLE}</h3>
      <ul class="dp-compare__pdf-not-compared-list">
        {PDF_NOT_COMPARED_ITEMS.map(item => (
          <li key={item.rest}>
            {item.lead && <strong>{item.lead}</strong>}
            {item.lead ? ` ${item.rest}` : item.rest}
          </li>
        ))}
      </ul>
    </section>
  );
}

/** Per-run image-presence notice — names pages with image content not text-compared. */
function ImageNotice({ notice }: { notice: PdfImageNotice }) {
  const lines = [
    imageNoticeLine('Original', notice.original),
    imageNoticeLine('Revised', notice.revised),
  ].filter((l): l is string => l !== null);
  if (lines.length === 0) return null;
  return (
    <section class="dp-compare__pdf-image-notice" role="note" aria-label="Image content notice">
      <p class="dp-compare__pdf-image-notice-lead">
        Some pages contain image content that was <strong>not text-compared</strong> — review
        them visually:
      </p>
      <ul class="dp-compare__pdf-image-notice-list">
        {lines.map(line => (
          <li key={line}>{line}</li>
        ))}
      </ul>
    </section>
  );
}

interface PdfSlotProps {
  side: 'original' | 'revised';
  label: string;
  ready: boolean;
  disabled: boolean;
  onPick: (file: File) => void;
}

function PdfSlot({ side, label, ready, disabled, onPick }: PdfSlotProps) {
  // The file <input> is nested inside its <label> with NO `for`/`id` — clicking
  // the label is the ONLY activation. A redundant `for`+nesting double-fires the
  // click in WebKit (Word-for-Mac task pane) and the native picker never opens;
  // this mirrors the proven .docx SlotPicker pattern.
  return (
    <div class={`dp-compare__pdf-slot dp-compare__pdf-slot--${side}`}>
      <span class="dp-compare__pdf-slot-label">{label}</span>
      <label class="dp-compare__upload">
        {ready ? 'Choose a different PDF…' : 'Choose PDF…'}
        <input
          class="dp-compare__pdf-slot-input"
          type="file"
          accept="application/pdf,.pdf"
          disabled={disabled}
          onChange={e => {
            const input = e.target as HTMLInputElement;
            const file = input.files?.[0];
            if (file) onPick(file);
          }}
        />
      </label>
      <span class="dp-compare__pdf-slot-status" role="status">
        {ready ? 'PDF selected' : 'No PDF selected'}
      </span>
    </div>
  );
}
