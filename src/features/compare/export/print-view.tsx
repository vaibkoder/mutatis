/**
 * Compare export — print render target (Tier 1: browser print-to-PDF via `window.print()`).
 *
 * `ExportPrintRoot` renders a {@link CompareReport} into a self-contained
 * `.dp-export-print` subtree; the `@media print` block in shell.css hides every
 * other element so the OS "Save as PDF" dialog captures ONLY this report (title,
 * disclosures, image notice, and EVERY change row). `exportReport()` mounts that
 * subtree into a hidden host on `document.body` and calls `window.print()`.
 *
 * Confidentiality (ADR-0001/0003): `window.print()` renders the local DOM — zero
 * network, no cloud render service. ADR-0008: the rows come straight from
 * `buildCompareReport`, which is filter-independent.
 */
import { render } from 'preact';
import type { CompareReport, ReportRow } from './report-model.js';

interface ExportPrintRootProps {
  report: CompareReport;
}

/** The print-only DOM subtree. Pure presentational — renders the full report. */
export function ExportPrintRoot({ report }: ExportPrintRootProps) {
  return (
    <div class="dp-export-print" lang="en-GB">
      <header class="dp-export-print__header">
        <h1 class="dp-export-print__title">{report.title}</h1>
        <p class="dp-export-print__mode">{report.modeLabel}</p>
        <p class="dp-export-print__generated">{report.generatedNote}</p>
      </header>

      {report.disclosures.length > 0 && (
        <section class="dp-export-print__disclosures" aria-label="What is not compared">
          <h2 class="dp-export-print__section-title">What is not compared</h2>
          <ul class="dp-export-print__disclosure-list">
            {report.disclosures.map(line => (
              <li key={line}>{line}</li>
            ))}
          </ul>
        </section>
      )}

      {report.imageNotice && report.imageNotice.length > 0 && (
        <section class="dp-export-print__image-notice" aria-label="Image content notice">
          <h2 class="dp-export-print__section-title">Image content not text-compared</h2>
          <ul class="dp-export-print__image-notice-list">
            {report.imageNotice.map(line => (
              <li key={line}>{line}</li>
            ))}
          </ul>
        </section>
      )}

      <section class="dp-export-print__changes" aria-label="Changes">
        <h2 class="dp-export-print__section-title">
          Changes ({report.changeCount})
        </h2>
        {report.rows.length === 0 ? (
          <p class="dp-export-print__no-changes">No candidate differences found.</p>
        ) : (
          <ol class="dp-export-print__change-list">
            {report.rows.map(row => (
              <PrintRow key={row.id} row={row} />
            ))}
          </ol>
        )}
      </section>
    </div>
  );
}

/**
 * Reader-facing names for the op types. `property` is the one that needs it: a
 * printed report handed to a lawyer said "property", which names the
 * implementation rather than the change. The other four already read as English
 * and are spelled out here so the column is consistent and the exhaustive map
 * fails the build if a sixth type is ever added without a name (ADR-0042 added
 * the fifth, and this column was the surface that was missed).
 */
const TYPE_LABELS: Record<ReportRow['type'], string> = {
  equal: 'unchanged',
  insert: 'inserted',
  delete: 'deleted',
  replace: 'amended',
  move: 'moved',
  property: 'formatting / numbering',
};

function PrintRow({ row }: { row: ReportRow }) {
  return (
    <li class="dp-export-print__row">
      <p class="dp-export-print__row-meta">
        <span class="dp-export-print__row-tag">{row.classTag}</span>
        <span class="dp-export-print__row-type">{TYPE_LABELS[row.type]}</span>
        <span class="dp-export-print__row-loc">{row.location}</span>
        {row.lowConfidence && <span class="dp-export-print__row-badge">low confidence</span>}
      </p>
      {row.original && (
        <p class="dp-export-print__row-original">
          <span class="dp-export-print__row-side">Original:</span> <del>{row.original}</del>
        </p>
      )}
      {row.revised && (
        <p class="dp-export-print__row-revised">
          <span class="dp-export-print__row-side">Revised:</span> <ins>{row.revised}</ins>
        </p>
      )}
      {row.moveTarget && (
        <p class="dp-export-print__row-move">Moved — paired change: {row.moveTarget}</p>
      )}
    </li>
  );
}

const PRINT_HOST_ID = 'dp-export-print-host';

/** Get (or lazily create) the hidden host the print subtree mounts into. */
function ensurePrintHost(): HTMLElement {
  let host = document.getElementById(PRINT_HOST_ID);
  if (!host) {
    host = document.createElement('div');
    host.id = PRINT_HOST_ID;
    document.body.appendChild(host);
  }
  return host;
}

/**
 * Render the report into the hidden print host and open the OS print dialog.
 * One export entry point regardless of mode (A4 contract). Tier 1: `window.print()`.
 */
export function exportReport(report: CompareReport): void {
  const host = ensurePrintHost();
  render(<ExportPrintRoot report={report} />, host);
  window.print();
}
