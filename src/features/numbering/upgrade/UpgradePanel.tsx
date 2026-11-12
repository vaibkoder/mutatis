/**
 * Feature 06 — Numbering upgrade (ADR-0023): task-pane section (M11).
 *
 * A two-step, human-gated power-up rendered inside the Numbering tab:
 *   Step 1 — assisted removal: delete hand-typed clause numbers across the
 *            current selection AS TRACKED CHANGES.
 *   Step 2 — re-apply: turn the cleared selection into a real Word list using the
 *            scheme selected above, then review cross-references (§5.3).
 *
 * Presentational only — all orchestration lives in upgrade-driver.ts.
 */

import { useEffect } from 'preact/hooks';
import {
  upgradeRunning,
  upgradeError,
  trackingUnavailable,
  lastRemovalReport,
  sequenceAnomalies,
  whitespaceOverride,
  removalQueue,
  activateUpgrade,
  setOverride,
  runRemoval,
  runReapply,
} from './upgrade-driver.js';
import type { WhitespaceOverride } from './token-span.js';

const OVERRIDE_OPTIONS: { value: WhitespaceOverride; label: string }[] = [
  { value: 'auto',                  label: 'Auto (tidy to flush)' },
  { value: 'collapse-to-one-space', label: 'Keep one space' },
  { value: 'strip-to-flush',        label: 'Strip all whitespace' },
  { value: 'keep-one-tab',          label: 'Keep one tab' },
];

export function UpgradePanel() {
  useEffect(() => { void activateUpgrade(); }, []);

  const running = upgradeRunning.value;
  const disabled = trackingUnavailable.value;
  const report = lastRemovalReport.value;
  const anomalies = sequenceAnomalies.value;
  const err = upgradeError.value;
  const candidates = removalQueue.value.length;

  return (
    <section class="dp-num-upgrade" aria-label="Numbering upgrade">
      <header class="dp-num-upgrade__header">
        <h2 class="dp-num-upgrade__title">Numbering upgrade</h2>
        <p class="dp-num-upgrade__intro">
          Convert hand-typed clause numbers into a real Word list. Removal is
          recorded as tracked changes; you choose which paragraphs to re-number.
        </p>
      </header>

      {disabled && (
        <p class="dp-num-upgrade__banner" role="note">
          Track Changes can’t be enabled on this host — assisted removal is
          unavailable. Numbers must be removed manually.
        </p>
      )}

      {candidates > 0 && (
        <p class="dp-num-upgrade__count" role="status">
          {candidates} hand-typed number{candidates === 1 ? '' : 's'} detected in the document.
        </p>
      )}

      <div class="dp-num-upgrade__whitespace">
        <label class="dp-num-upgrade__ws-label" for="dp-num-ws-override">
          Whitespace after removal
        </label>
        <select
          id="dp-num-ws-override"
          value={whitespaceOverride.value}
          disabled={running}
          onChange={e => void setOverride(
            (e.currentTarget as HTMLSelectElement).value as WhitespaceOverride,
          )}
        >
          {OVERRIDE_OPTIONS.map(o => (
            <option key={o.value} value={o.value}>{o.label}</option>
          ))}
        </select>
      </div>

      <div class="dp-num-upgrade__steps">
        <div class="dp-num-upgrade__step">
          <span class="dp-num-upgrade__step-num">1</span>
          <button
            type="button"
            class="dp-num-btn dp-num-btn--primary"
            disabled={running || disabled}
            aria-busy={running ? 'true' : 'false'}
            onClick={() => void runRemoval()}
            title="Delete leading hand-typed numbers across the current selection as tracked changes"
          >
            {running ? 'Working…' : 'Remove numbers (tracked)'}
          </button>
        </div>

        <div class="dp-num-upgrade__step">
          <span class="dp-num-upgrade__step-num">2</span>
          <button
            type="button"
            class="dp-num-btn"
            disabled={running}
            onClick={() => void runReapply()}
            title="Turn the selected cleared paragraphs into a numbered list using the scheme above"
          >
            Number selected paragraphs
          </button>
        </div>
      </div>

      {err && <p class="dp-num-error" role="alert">{err}</p>}

      {report && !report.disabled && (
        <p class="dp-num-upgrade__report" role="status">
          Removed {report.deleted} of {report.total} paragraph(s).
          {report.skipped > 0 && ` ${report.skipped} had no leading number.`}
          {report.notFound > 0 && ` ${report.notFound} could not be matched automatically — remove manually.`}
        </p>
      )}

      {anomalies.length > 0 && (
        <section class="dp-num-upgrade__anomalies" aria-label="Numbering anomalies">
          <h3 class="dp-num-upgrade__anomalies-title">
            Sequence anomalies ({anomalies.length})
          </h3>
          <p class="dp-num-upgrade__anomalies-note">
            Surfaced, not changed — you decide. Re-applying a dynamic list dissolves a typo;
            keep the gap if the number was deliberately reserved.
          </p>
          <ul class="dp-num-upgrade__anomalies-list">
            {anomalies.map((a, i) => (
              <li key={i} class="dp-num-upgrade__anomaly">
                <code class="dp-num-upgrade__anomaly-token">{a.token}</code>
                {' '}
                <span class="dp-num-upgrade__anomaly-kind">{a.kind}</span>
                {' after '}
                <code>{a.previousToken}</code>
                <span class="dp-num-upgrade__anomaly-readings">
                  Typo: {a.readings.typo} · Reserved: {a.readings.reserved}
                </span>
              </li>
            ))}
          </ul>
        </section>
      )}
    </section>
  );
}
