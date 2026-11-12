/**
 * Feature 07-A — Table of Contents: task-pane sub-panel (A6).
 *
 * Replaces the stub created by the S scaffold. Implements the full TOC UI
 * per spec §A.7:
 *
 *   1. F9 banner   — shown when wordApi15 is absent.
 *   2. Profile selector — firm profiles from config.
 *   3. Insert TOC — at cursor or top of document.
 *   4. Make scoped TOC from selection — enabled only on non-empty selection.
 *   5. Generate / Regenerate Index of Schedules.
 *   6. Refresh this TOC / Refresh all TOCs.
 *   7. Registered-TOC list with go-to action.
 *   8. Error + info banner display.
 *
 * All Office JS work is delegated to driver.ts (signals) and toc-service.ts.
 * This component is pure presentation — no direct Office JS calls.
 *
 * Lang attribute: en-GB throughout (no US-isms in labels).
 */

import { useEffect, useState } from 'preact/hooks';
import {
  // Signals
  activeProfile,
  availableProfiles,
  registry,
  busy,
  error,
  showF9Banner,
  banners,
  // Actions
  activateToc,
  selectProfile,
  runInsertTocAtCursor,
  runInsertTocAtTop,
  runInsertScopedToc,
  runGenerateScheduleIndex,
  runRefreshOne,
  runRefreshAll,
  runNavigateToToc,
} from './driver.js';

// ---------------------------------------------------------------------------
// TocPanel (root)
// ---------------------------------------------------------------------------

export function TocPanel() {
  const isBusy    = busy.value;
  const err       = error.value;
  const f9Banner  = showF9Banner.value;
  const infoBanners = banners.value;
  const profile   = activeProfile.value;
  const profiles  = availableProfiles.value;
  const tocList   = registry.value;

  // Local state for the scoped-TOC label input.
  const [scopedLabel, setScopedLabel] = useState('Scoped TOC');

  // Activate on mount.
  useEffect(() => {
    void activateToc();
  }, []);

  return (
    <div class="dp-toc" lang="en-GB">

      {/* ── 1. F9 banner — shown when WordApi 1.5 is absent (A-T5) ── */}
      {f9Banner && (
        <aside class="dp-toc__f9-banner" role="alert">
          <strong>Update fields manually:</strong> Your version of Word does not
          support automatic table of contents insertion. After inserting, press{' '}
          <kbd>F9</kbd> or right-click the field and choose{' '}
          <em>Update Field</em> to populate the TOC.
        </aside>
      )}

      {/* ── 2. Error display ─────────────────────────────────────────── */}
      {err && (
        <p class="dp-toc__error" role="alert">
          {err}
        </p>
      )}

      {/* ── 3. Info banners ──────────────────────────────────────────── */}
      {infoBanners.map((msg, i) => (
        <p key={i} class="dp-toc__banner">
          {msg}
        </p>
      ))}

      {/* ── 4. Profile selector ──────────────────────────────────────── */}
      <section class="dp-toc__section dp-toc__profile">
        <label class="dp-toc__label" for="dp-toc-profile">
          TOC profile:
        </label>
        <select
          id="dp-toc-profile"
          class="dp-toc__select"
          value={profile?.id ?? ''}
          onChange={e => void selectProfile((e.target as HTMLSelectElement).value)}
          disabled={isBusy}
        >
          {profiles.map(p => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>
      </section>

      {/* ── 5. Insert TOC ────────────────────────────────────────────── */}
      <section class="dp-toc__section">
        <h3 class="dp-toc__section-heading">Insert Table of Contents</h3>
        <div class="dp-toc__row">
          <button
            type="button"
            class="dp-toc__btn dp-toc__btn--primary"
            onClick={() => void runInsertTocAtCursor()}
            disabled={isBusy}
          >
            Insert at cursor
          </button>
          <button
            type="button"
            class="dp-toc__btn"
            onClick={() => void runInsertTocAtTop()}
            disabled={isBusy}
          >
            Insert at top of document
          </button>
        </div>
      </section>

      {/* ── 6. Scoped (multiple) TOC ─────────────────────────────────── */}
      <section class="dp-toc__section">
        <h3 class="dp-toc__section-heading">Scoped TOC from selection</h3>
        <p class="dp-toc__hint">
          Select a range in the document, enter a label, then click{' '}
          <em>Make scoped TOC</em>. The TOC will cover only the selected section.
        </p>
        <div class="dp-toc__row">
          <label class="dp-toc__label dp-toc__label--inline" for="dp-toc-scope-label">
            Label:
          </label>
          <input
            id="dp-toc-scope-label"
            type="text"
            class="dp-toc__input"
            value={scopedLabel}
            onInput={e => setScopedLabel((e.target as HTMLInputElement).value)}
            disabled={isBusy}
            maxLength={60}
          />
        </div>
        <button
          type="button"
          class="dp-toc__btn"
          onClick={() => void runInsertScopedToc(scopedLabel.trim() || 'Scoped TOC')}
          disabled={isBusy}
          aria-label="Make scoped TOC from current selection"
        >
          Make scoped TOC from selection
        </button>
      </section>

      {/* ── 7. Index of Schedules ────────────────────────────────────── */}
      <section class="dp-toc__section">
        <h3 class="dp-toc__section-heading">Index of Schedules</h3>
        <button
          type="button"
          class="dp-toc__btn"
          onClick={() => void runGenerateScheduleIndex()}
          disabled={isBusy}
        >
          Generate / Regenerate Index of Schedules
        </button>
      </section>

      {/* ── 8. Refresh ───────────────────────────────────────────────── */}
      <section class="dp-toc__section">
        <h3 class="dp-toc__section-heading">Refresh</h3>
        <div class="dp-toc__row">
          <button
            type="button"
            class="dp-toc__btn"
            onClick={() => void runRefreshAll()}
            disabled={isBusy}
          >
            Refresh all TOCs
          </button>
        </div>
      </section>

      {/* ── 9. Registered TOC list ───────────────────────────────────── */}
      {tocList.length > 0 && (
        <section class="dp-toc__section">
          <h3 class="dp-toc__section-heading">Inserted TOCs</h3>
          <ul class="dp-toc__list" role="list">
            {tocList.map((entry, idx) => (
              <li key={idx} class="dp-toc__list-item">
                <span class="dp-toc__list-label">
                  {entry.scopeLabel}
                  {entry.bookmark && (
                    <span class="dp-toc__list-badge"> (scoped)</span>
                  )}
                </span>
                <div class="dp-toc__list-actions">
                  {entry.fieldId && (
                    <button
                      type="button"
                      class="dp-toc__btn dp-toc__btn--small"
                      onClick={() => void runNavigateToToc(entry.fieldId)}
                      disabled={isBusy}
                      aria-label={`Go to ${entry.scopeLabel}`}
                    >
                      Go to
                    </button>
                  )}
                  <button
                    type="button"
                    class="dp-toc__btn dp-toc__btn--small"
                    onClick={() => void runRefreshOne(entry.fieldId)}
                    disabled={isBusy || !entry.fieldId}
                    aria-label={`Refresh ${entry.scopeLabel}`}
                    title={!entry.fieldId ? 'Press F9 in Word to refresh this TOC' : undefined}
                  >
                    Refresh
                  </button>
                </div>
              </li>
            ))}
          </ul>
        </section>
      )}

      {/* ── 10. Busy overlay ─────────────────────────────────────────── */}
      {isBusy && (
        <div class="dp-toc__busy" role="status" aria-live="polite">
          Working…
        </div>
      )}
    </div>
  );
}
