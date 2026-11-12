/**
 * apps/web — the standalone Compare surface (ADR-0022, spec/features/12 §5 phase 4).
 *
 * Renders the UNCHANGED `ComparePanel` (feature 08 / 10) wired to the LIVE driver
 * states, so the web app gets all three Compare modes without Word:
 *   - Word-vs-Word (.docx upload, both sides) → redline `.docx` download.
 *   - Text-vs-text (paste) → inline preview + redline `.docx` download.
 *   - PDF-vs-PDF (upload) → in-pane candidate differences (no redline).
 *
 * Host coupling flows entirely through the active `HostPort` (12f Group A): the
 * web entry installs `WebHostPort`, so Compare's redline write becomes a Blob
 * download (`WebHostPort.writeBack`) and Word-vs-Word `goToChange` is a clean
 * no-op (`WebHostPort.findAndReveal` is omitted). Nothing here touches Office JS.
 *
 * Two web-specific wirings:
 *   1. Inject the LIVE `compareDriverState`/`textCompareDriverState`/
 *      `pdfCompareDriverState` — the panels default to no-op STUB drivers, so
 *      without this upload/compare silently does nothing (memory:
 *      compare-pdf-driver-shell-wiring-gap). Mirrors the add-in's Shell.tsx.
 *   2. `uploadOnly` — the web target has no active host document, so the
 *      Word-vs-Word slots hide "Use current document" and upload both sides.
 */
import {
  ComparePanel,
  compareDriverState,
  textCompareDriverState,
  pdfCompareDriverState,
} from '@features/compare/index.js';

export function WebCompare() {
  return (
    <div class="dp-web__compare" aria-label="Compare two documents">
      <ComparePanel
        driver={compareDriverState}
        textDriver={textCompareDriverState}
        pdfDriver={pdfCompareDriverState}
        uploadOnly
      />
    </div>
  );
}
