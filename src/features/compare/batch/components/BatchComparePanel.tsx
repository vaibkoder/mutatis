/**
 * The one-to-many batch panel (ADR-0037; spec §5.4).
 *
 * A template slot, a variant queue, a two-level progress readout, and one row
 * per pair. The row list is the UI face of the batch index and inherits its
 * rule: every pair the batch was asked to run appears, whatever became of it. A
 * failed or unrunnable pair is shown with its reason, never hidden — a list that
 * quietly omitted its failures would look complete and be wrong.
 *
 * `driver` defaults to a no-op stub so the panel renders in isolation. Any
 * surface mounting it MUST inject the live driver; forgetting to is silent
 * (`web/WebCompare.tsx:15-22` records the previous occasion).
 */
import { signal } from '@preact/signals';

import type { BatchDriverState } from '../ui-types.js';
import { makeStubBatchDriverState } from '../ui-types.js';
import { isDownloadable } from '../artefacts.js';
import { BatchRowItem, BatchRunFootnotes } from './BatchRowList.js';

export interface BatchComparePanelProps {
  driver?: BatchDriverState;
  /** Open one pair's full change list. Wired by the mounting surface. */
  onOpenRow?: (rowId: string) => void;
}

function filesOf(event: Event): File[] {
  const input = event.currentTarget as HTMLInputElement;
  return Array.from(input.files ?? []);
}

export function BatchComparePanel({ driver, onOpenRow }: BatchComparePanelProps) {
  const state = driver ?? makeStubBatchDriverState(signal);

  const template = state.template.value;
  const variants = state.variants.value;
  const rows = state.rows.value;
  const status = state.status.value;
  const running = status === 'running';
  const anyDownloadable = rows.some(isDownloadable);
  const finished = state.index.value;

  return (
    <section class="dp-batch" aria-label="Compare one template against many documents">
      <p class="dp-batch__intro">
        Compare one template against every version you drop in. Each pair produces its
        own redline — nothing leaves your machine.
      </p>

      <div class="dp-batch__slots">
        <label class="dp-batch__slot">
          <span class="dp-batch__slot-title">Template</span>
          <input
            type="file"
            accept=".docx"
            // Locked once rows exist for the same reason the versions are: the
            // batch's rows were produced against THIS template, and swapping it
            // would make every later re-run compare against a different one.
            disabled={running || rows.length > 0}
            onChange={e => {
              const [file] = filesOf(e);
              if (file) state.setTemplate(file);
            }}
          />
          <span class="dp-batch__slot-value">{template?.label ?? 'No template chosen'}</span>
        </label>

        <label class="dp-batch__slot">
          <span class="dp-batch__slot-title">Versions to compare</span>
          <input
            type="file"
            accept=".docx"
            multiple
            // Also locked once a batch has produced rows: the queue list is
            // hidden then, so anything added would sit in the driver with
            // nothing on screen to show or remove it.
            disabled={running || rows.length > 0}
            onChange={e => state.addVariants(filesOf(e))}
          />
          <span class="dp-batch__slot-value">
            {variants.length} of {state.maxVariants}
          </span>
        </label>
      </div>

      {variants.length > 0 && !running && rows.length === 0 && (
        <ul class="dp-batch__queue">
          {variants.map(v => (
            <li key={v.id} class="dp-batch__queue-item">
              <span>{v.label}</span>
              <button
                type="button"
                class="dp-batch__link"
                onClick={() => state.removeVariant(v.id)}
                aria-label={`Remove ${v.label}`}
              >
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}

      <div class="dp-batch__actions">
        <button
          type="button"
          class="dp-batch__run"
          disabled={running || template === null || variants.length === 0}
          onClick={() => { void state.runBatch(); }}
        >
          Compare {variants.length > 0 ? `${variants.length} version${variants.length === 1 ? '' : 's'}` : 'versions'}
        </button>

        {running && (
          <button type="button" class="dp-batch__link" onClick={() => state.cancelBatch()}>
            Cancel
          </button>
        )}

        {rows.length > 0 && !running && (
          <button
            type="button"
            class="dp-batch__link"
            onClick={() => state.clearVariants()}
          >
            Clear batch
          </button>
        )}

        {anyDownloadable && !running && (
          <button
            type="button"
            class="dp-batch__link"
            onClick={() => { void state.downloadAll(); }}
          >
            Download all redlines
          </button>
        )}
      </div>

      {running && (
        <div class="dp-batch__progress">
          <progress max={100} value={state.progress.value} />
          <span>
            {rows.filter(r => r.status !== 'queued' && r.status !== 'running').length} of{' '}
            {variants.length} done
          </span>
        </div>
      )}

      {state.errorMessage.value !== null && (
        <p class="dp-batch__error" role="alert">{state.errorMessage.value}</p>
      )}

      {rows.length > 0 && (
        <ul class="dp-batch__rows">
          {rows.map(row => (
            <BatchRowItem
              key={row.id}
              row={row}
              {...(onOpenRow !== undefined ? { onOpen: onOpenRow } : {})}
              onDownload={id => { void state.downloadRow(id); }}
            />
          ))}
        </ul>
      )}

      {state.index.value?.disclosures.map(text => (
        <p key={text} class="dp-batch__disclosure">{text}</p>
      ))}

      {finished !== null && <BatchRunFootnotes index={finished} />}
    </section>
  );
}
