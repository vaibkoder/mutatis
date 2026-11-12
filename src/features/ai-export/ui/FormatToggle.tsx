/**
 * JSON | MD segmented control (feature 11 — spec §6.1).
 *
 * Switches the shared, persisted export format. One instance per copy button;
 * all instances reflect the same `exportFormat` signal.
 */
import { exportFormat, setExportFormat } from './prefs.js';

export function FormatToggle({ idBase }: { idBase: string }) {
  const fmt = exportFormat.value;
  return (
    <div class="dp-ai-export__format" role="group" aria-label="Export format">
      <button
        type="button"
        id={`${idBase}-json`}
        class="dp-ai-export__format-opt"
        aria-pressed={fmt === 'json' ? 'true' : 'false'}
        onClick={() => setExportFormat('json')}
      >
        JSON
      </button>
      <button
        type="button"
        id={`${idBase}-md`}
        class="dp-ai-export__format-opt"
        aria-pressed={fmt === 'markdown' ? 'true' : 'false'}
        onClick={() => setExportFormat('markdown')}
      >
        MD
      </button>
    </div>
  );
}
