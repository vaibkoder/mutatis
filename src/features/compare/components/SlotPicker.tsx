/**
 * Feature 08 — Compare: a single source slot (Original or Revised). §8.1.
 *
 * "Use current document" (office: getFileAsync via the host) or "Upload .docx"
 * (→ File API). On the standalone web target there is no active host document,
 * so `uploadOnly` hides the "Use current document" button (12f). Pure
 * presentation: it calls back into the driver; it holds no state itself.
 */
import type { CompareDriverState } from '../ui-types.js';

interface SlotPickerProps {
  side: 'original' | 'revised';
  /** Human label for the slot, e.g. "Original (base)". */
  label: string;
  driver: CompareDriverState;
  ready: boolean;
  /** Web target: hide "Use current document" — no active host doc (12f). */
  uploadOnly?: boolean;
}

export function SlotPicker({ side, label, driver, ready, uploadOnly = false }: SlotPickerProps) {
  return (
    <div class={`dp-compare__slot dp-compare__slot--${side}`}>
      <div class="dp-compare__slot-label">{label}</div>
      <div class="dp-compare__slot-controls">
        {!uploadOnly && (
          <button
            type="button"
            class="dp-compare__use-current"
            onClick={() => driver.useCurrentDocument(side)}
          >
            Use current document
          </button>
        )}
        <label class="dp-compare__upload">
          Upload .docx
          <input
            type="file"
            accept=".docx"
            class="dp-compare__file-input"
            onChange={e => {
              const input = e.target as HTMLInputElement;
              const file = input.files?.[0];
              if (file) driver.uploadDocx(side, file);
            }}
          />
        </label>
      </div>
      <div class="dp-compare__slot-status" role="status">
        {ready ? 'Source selected' : 'No source selected'}
      </div>
    </div>
  );
}
