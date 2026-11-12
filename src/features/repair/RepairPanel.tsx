/**
 * Feature 05 — Document Repair: task-pane panel.
 *
 * Spec §7.1 layout:
 *   1. Profile selector.
 *   2. Trust banner (always visible — client-side guarantee).
 *   3. Guard warnings (court-formatting / commercial-on-court).
 *   4. Error display.
 *   5. Scan button (pre-scan) OR findings checklist + Apply button (post-scan).
 *   6. Apply-phase result summary and re-scan button.
 */

import type { RepairFinding } from './types.js';
import {
  repairProfile,
  repairScanReport,
  repairApplyReport,
  repairBusy,
  repairError,
  repairChecked,
  repairGroupedFindings,
  repairNormalisationWarning,
  repairCommercialOnCourtWarning,
  repairConfirmedSectionBreaks,
  setProfile,
  runScan,
  runApply,
  toggleFinding,
  confirmSectionBreak,
  confirmProtectedBreak,
} from './driver.js';
import { listProfileIds, getProfile } from './profiles.js';

// ---------------------------------------------------------------------------
// FindingRow
// ---------------------------------------------------------------------------

interface FindingRowProps {
  finding: RepairFinding;
  checked: boolean;
  confirmed: boolean;
  phase: 'scan' | 'apply';
}

function FindingRow({ finding, checked, confirmed, phase }: FindingRowProps) {
  const isDestructive = finding.destructive || finding.requiresConfirm;
  const resultBadge =
    phase === 'apply' && finding.result != null ? ` [${finding.result}]` : '';

  return (
    <div
      class={`dp-repair__finding${isDestructive ? ' dp-repair__finding--destructive' : ''}`}
    >
      <label class="dp-repair__finding-label">
        <input
          type="checkbox"
          checked={isDestructive ? confirmed : checked}
          disabled={phase === 'apply' || isDestructive}
          onChange={() => {
            if (!isDestructive) toggleFinding(finding.id);
          }}
        />
        {finding.label}
        {finding.count != null ? ` (${finding.count})` : ''}
        {isDestructive && (
          <span class="dp-repair__lock" aria-label="requires confirmation">
            {'🔒' /* 🔒 */}
          </span>
        )}
        {resultBadge && <span class="dp-repair__result">{resultBadge}</span>}
      </label>

      {finding.detail && (
        <p class="dp-repair__finding-detail">{finding.detail}</p>
      )}

      {isDestructive && !confirmed && phase === 'scan' && (
        <button
          type="button"
          class={`dp-repair__confirm-btn${finding.protected ? ' dp-repair__confirm-btn--protected' : ''}`}
          onClick={() => {
            if (finding.protected) {
              // Hard rule (spec §5.7 / §7.2): orientation-break removal requires
              // a dedicated second confirmation warning that the Schedule reverts
              // to portrait. Only on explicit accept do we route through the
              // separate protected channel.
              const ok = globalThis.confirm?.(
                'This section break changes page orientation (e.g. a landscape ' +
                  'Schedule). Removing it reverts that content to portrait and ' +
                  'cannot be selectively undone. Remove it anyway?',
              );
              if (ok) confirmProtectedBreak(finding.id);
            } else {
              confirmSectionBreak(finding.id);
            }
          }}
        >
          {finding.protected
            ? 'Confirm (changes orientation)'
            : 'Confirm'}
        </button>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// RepairPanel (root)
// ---------------------------------------------------------------------------

export function RepairPanel() {
  const profile  = repairProfile.value;
  const busy     = repairBusy.value;
  const err      = repairError.value;
  const scanReport  = repairScanReport.value;
  const applyReport = repairApplyReport.value;
  const report   = scanReport ?? applyReport;
  const grouped  = repairGroupedFindings.value;
  const checked  = repairChecked.value;
  const confirmed = repairConfirmedSectionBreaks.value;
  const normWarn  = repairNormalisationWarning.value;
  const commWarn  = repairCommercialOnCourtWarning.value;

  return (
    <div class="dp-repair" lang="en-GB">
      {/* ── 1. Profile selector ─────────────────────────────────────── */}
      <div class="dp-repair__header">
        <label class="dp-repair__profile-label">
          Profile:
          <select
            class="dp-repair__profile-select"
            value={profile.id}
            onChange={e => setProfile((e.target as HTMLSelectElement).value)}
            disabled={busy}
          >
            {listProfileIds().map(id => {
              const p = getProfile(id)!;
              return (
                <option key={id} value={id}>
                  {p.name}
                </option>
              );
            })}
          </select>
        </label>
      </div>

      {/* ── 2. Trust banner — always visible ────────────────────────── */}
      <aside class="dp-repair__trust-banner" role="note">
        All repairs run on your device. Nothing is uploaded or shared.
      </aside>

      {/* ── 3. Guard warnings ───────────────────────────────────────── */}
      {normWarn && <p class="dp-repair__warning">{normWarn}</p>}
      {commWarn && (
        <p class="dp-repair__warning dp-repair__warning--commercial">
          {commWarn}
        </p>
      )}

      {/* ── 4. Error display ────────────────────────────────────────── */}
      {err && (
        <p class="dp-repair__error" role="alert">
          {err}
        </p>
      )}

      {/* ── 5a. Scan button (pre-scan only) ─────────────────────────── */}
      {!report && (
        <button
          type="button"
          class="dp-repair__scan-btn"
          onClick={() => void runScan(document.title || 'Document')}
          disabled={busy}
          aria-busy={busy ? 'true' : 'false'}
        >
          {busy ? 'Scanning…' : 'Scan document'}
        </button>
      )}

      {/* ── 5b. Findings checklist ──────────────────────────────────── */}
      {grouped && (
        <div class="dp-repair__findings">
          {grouped.confidentiality.length > 0 && (
            <section class="dp-repair__group">
              <h3 class="dp-repair__group-title dp-repair__group-title--confidentiality">
                Confidentiality
              </h3>
              {grouped.confidentiality.map(f => (
                <FindingRow
                  key={f.id}
                  finding={f}
                  checked={checked.has(f.id)}
                  confirmed={confirmed.has(f.id)}
                  phase={report?.phase ?? 'scan'}
                />
              ))}
            </section>
          )}

          {grouped.stabilisation.length > 0 && (
            <section class="dp-repair__group">
              <h3 class="dp-repair__group-title dp-repair__group-title--stabilisation">
                Stabilisation
              </h3>
              {grouped.stabilisation.map(f => (
                <FindingRow
                  key={f.id}
                  finding={f}
                  checked={checked.has(f.id)}
                  confirmed={confirmed.has(f.id)}
                  phase={report?.phase ?? 'scan'}
                />
              ))}
            </section>
          )}

          {grouped.info.length > 0 && (
            <section class="dp-repair__group">
              <h3 class="dp-repair__group-title">Info</h3>
              {grouped.info.map(f => (
                <FindingRow
                  key={f.id}
                  finding={f}
                  checked={checked.has(f.id)}
                  confirmed={confirmed.has(f.id)}
                  phase={report?.phase ?? 'scan'}
                />
              ))}
            </section>
          )}
        </div>
      )}

      {/* ── 5c. Apply button (scan phase only) ──────────────────────── */}
      {scanReport && !busy && (
        <div class="dp-repair__actions">
          <button
            type="button"
            class="dp-repair__apply-btn"
            onClick={() => void runApply()}
          >
            Apply selected
          </button>
        </div>
      )}

      {/* ── 6a. Apply result summary ─────────────────────────────────── */}
      {applyReport && (
        <p class="dp-repair__apply-summary" role="status">
          Repair complete.{' '}
          {applyReport.findings.filter(f => f.result === 'done').length}{' '}
          action(s) applied.
        </p>
      )}

      {/* ── 6b. Re-scan after apply ──────────────────────────────────── */}
      {applyReport && (
        <button
          type="button"
          class="dp-repair__rescan-btn"
          onClick={() => void runScan(document.title || 'Document')}
          disabled={busy}
        >
          Scan again
        </button>
      )}
    </div>
  );
}
