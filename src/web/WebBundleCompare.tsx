/**
 * The many-to-many surface (ADR-0037; spec §6).
 *
 * Two bundles in, a pairing grid, and — only once that grid is confirmed — a
 * batch of redlines. The run button is gated on `confirmed`, and so is the
 * driver behind it: the disabled attribute is the courtesy, `runConfirmed`'s
 * own refusal is the invariant. Both, because a gate that exists only in a
 * component is a gate that the next component forgets.
 *
 * The batch panel is deliberately NOT reused here. It is built around one
 * template and N variants; this surface has two bundles and a proposal, and
 * the row list is the only part they share.
 *
 * Opening a row re-runs that single pair's pipeline (spec §2.1 capability 8) —
 * never the artefact path, which packages a redline and on web means triggering
 * a download. "Show me what changed" must not save a file.
 */
import { useMemo } from 'preact/hooks';

import { createBundleDriver } from '@features/compare/batch/bundle-driver.js';
import type { BundleDriverState } from '@features/compare/batch/bundle-driver.js';
import { PairingGrid } from '@features/compare/batch/components/PairingGrid.js';
import { isDownloadable } from '@features/compare/batch/artefacts.js';
import {
  PairDrilldown,
  usePairDrilldown,
} from '@features/compare/batch/components/PairDrilldown.js';
import {
  BatchRowItem,
  BatchRunFootnotes,
} from '@features/compare/batch/components/BatchRowList.js';
import type { BundleSide } from '@features/compare/batch/pairing/sketch.js';

function filesOf(event: Event): File[] {
  const input = event.currentTarget as HTMLInputElement;
  return Array.from(input.files ?? []);
}

function BundleSlot({ side, title, driver }: {
  side: BundleSide;
  title: string;
  driver: BundleDriverState;
}) {
  const docs = side === 'a' ? driver.sideA.value : driver.sideB.value;
  const locked = driver.busy.value || driver.status.value === 'running';

  return (
    <label class="dp-batch__slot">
      <span class="dp-batch__slot-title">{title}</span>
      <input
        type="file"
        accept=".docx"
        multiple
        disabled={locked}
        onChange={e => driver.addDocuments(side, filesOf(e))}
      />
      <span class="dp-batch__slot-value">
        {docs.length} of {driver.maxBundleDocuments}
      </span>
      {docs.length > 0 && (
        <ul class="dp-batch__queue">
          {docs.map(d => (
            <li key={d.id} class="dp-batch__queue-item">
              <span>{d.label}</span>
              <button
                type="button"
                class="dp-batch__link"
                disabled={locked}
                onClick={() => driver.removeDocument(side, d.id)}
                aria-label={`Remove ${d.label}`}
              >
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
    </label>
  );
}

export function WebBundleCompare({ driver }: { driver?: BundleDriverState }) {
  // One driver for the life of the surface. Re-creating it per render would
  // drop the bundles — and the confirmation — on the first unrelated change.
  const state = useMemo(() => driver ?? createBundleDriver(), [driver]);

  // Opening a row re-runs that ONE pair's pipeline. The batch drops every
  // pair's ops the moment it completes — retaining N of them is what breaks
  // first (spec §4.3) — so the change list is recovered, not remembered. It is
  // deliberately not the artefact path: packaging a redline on web means
  // triggering a download, and "show me what changed" must not save a file.
  const drilldown = usePairDrilldown(
    rowId => state.openPair(rowId),
    'That comparison is no longer available. Run the comparisons again to reopen it.',
  );

  const running = state.status.value === 'running';
  const rows = state.rows.value;
  const bothSides = state.sideA.value.length > 0 && state.sideB.value.length > 0;

  return (
    <div class="dp-web__bundle" aria-label="Compare two bundles of documents">
      <p class="dp-batch__intro">
        Compare one set of documents against another. Nothing is compared until you
        have checked the proposed pairing — nothing leaves your machine.
      </p>

      <div class="dp-batch__slots">
        <BundleSlot side="a" title="Earlier bundle" driver={state} />
        <BundleSlot side="b" title="Later bundle" driver={state} />
      </div>

      <div class="dp-batch__actions">
        <button
          type="button"
          class="dp-batch__run"
          disabled={!bothSides || state.busy.value || running}
          onClick={() => { void state.proposePairs(); }}
        >
          Propose a pairing
        </button>

        {rows.length > 0 && !running && (
          <button type="button" class="dp-batch__link" onClick={() => state.clearBundles()}>
            Start again
          </button>
        )}

        {rows.some(isDownloadable) && !running && (
          <button
            type="button"
            class="dp-batch__link"
            onClick={() => { void state.downloadAll(); }}
          >
            Download all redlines
          </button>
        )}
      </div>

      {state.errorMessage.value !== null && (
        <p class="dp-batch__error" role="alert">{state.errorMessage.value}</p>
      )}

      <PairingGrid state={state} />

      {state.proposal.value !== null && (
        <div class="dp-batch__actions">
          <button
            type="button"
            class="dp-batch__run"
            // Gated on the confirmation. `runConfirmed` refuses too — this is
            // the courtesy, that is the invariant.
            disabled={!state.confirmed.value || running}
            onClick={() => { void state.runConfirmed(); }}
          >
            Run comparisons
          </button>

          {running && (
            <button type="button" class="dp-batch__link" onClick={() => state.cancelRun()}>
              Cancel
            </button>
          )}
        </div>
      )}

      {running && (
        <div class="dp-batch__progress">
          <progress max={100} value={state.progress.value} />
          <span>
            {rows.filter(r => r.status !== 'queued' && r.status !== 'running').length} of{' '}
            {rows.length} done
          </span>
        </div>
      )}

      {rows.length > 0 && (
        <ul class="dp-batch__rows">
          {rows.map(row => (
            <BatchRowItem
              key={row.id}
              row={row}
              onOpen={(id: string) => drilldown.open(id)}
              onDownload={(id: string) => { void state.downloadRow(id); }}
            />
          ))}
        </ul>
      )}

      {state.index.value?.disclosures.map(text => (
        <p key={text} class="dp-batch__disclosure">{text}</p>
      ))}

      {/* §5.5's cancellation notice, §8 case 7's flatten statement and §10's
          limitations. Shared with the one-to-many panel rather than written
          again: the hand-rolled copy this replaced had all three missing. */}
      {state.index.value !== null && <BatchRunFootnotes index={state.index.value} />}

      <PairDrilldown state={drilldown} />
    </div>
  );
}
