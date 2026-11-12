/**
 * Feature 06 — Numbering upgrade (ADR-0023): driver for assisted static-number
 * removal (Step 1) and human-driven re-apply (Step 2) — M11.
 *
 * Orchestration only: signals + Word.run wiring. The numbering logic lives in
 * removal.ts / reapply.ts; the cross-reference review and apply report reuse the
 * parent numbering driver's signals so the shipped review panel surfaces results.
 */

import { signal } from '@preact/signals';
import { wordRun } from '@shared/office/index.js';
import { getActiveProfile } from '@shared/config/index.js';
import { findingsByCategory } from '@shared/findings/index.js';
import type { Finding } from '@shared/findings/types.js';
import { removeLeadingNumbers } from './removal.js';
import type { RemovalReport } from './removal.js';
import { detectSequenceAnomalies } from './sequence-anomaly.js';
import type { SequenceAnomaly } from './sequence-anomaly.js';
import { startListAndApplyOnSelection } from './reapply.js';
import { getWhitespaceOverride, setWhitespaceOverride } from './whitespace-pref.js';
import type { WhitespaceOverride } from './token-span.js';
import { selectedScheme, reviewState, lastApplyReport } from '../driver.js';

const MANUALLY_TYPED = 'static-numbering/manually-typed';

// ---------------------------------------------------------------------------
// Signals
// ---------------------------------------------------------------------------

/** Whether a removal/re-apply operation is in progress. */
export const upgradeRunning = signal(false);

/** Error message for the last failed upgrade operation. */
export const upgradeError = signal<string | null>(null);

/**
 * True when the host refused to enable Track Changes, so Step 1 is unavailable
 * (constraint 1 — the panel greys the removal action and shows a banner).
 */
export const trackingUnavailable = signal(false);

/** Outcome of the last removal pass (deleted / skipped / not-found counts). */
export const lastRemovalReport = signal<RemovalReport | null>(null);

/**
 * Duplicate/gap anomalies among the cleared tokens (constraint 4). Surfaced with
 * BOTH readings (typo or deliberately reserved) — never auto-resequenced.
 */
export const sequenceAnomalies = signal<SequenceAnomaly[]>([]);

/** The single global whitespace rule for removal (constraint 3). */
export const whitespaceOverride = signal<WhitespaceOverride>('auto');

/**
 * Candidate paragraphs carrying hand-typed numbers — read-only from feature 02's
 * 'static-numbering/manually-typed' findings. A navigation/preview aid; the actual
 * removal operates on the live Word selection.
 */
export const removalQueue = signal<Finding[]>([]);

// ---------------------------------------------------------------------------
// Queue (read-only consumer of feature 02 findings)
// ---------------------------------------------------------------------------

function gatherManuallyTyped(): Finding[] {
  return Array.from(findingsByCategory('proof').values())
    .flat()
    .filter(f => f.category === MANUALLY_TYPED);
}

// ---------------------------------------------------------------------------
// Initialise (called when the upgrade section mounts)
// ---------------------------------------------------------------------------

export async function activateUpgrade(): Promise<void> {
  upgradeError.value = null;
  trackingUnavailable.value = false;
  lastRemovalReport.value = null;
  sequenceAnomalies.value = [];
  removalQueue.value = gatherManuallyTyped();
  try {
    whitespaceOverride.value = await getWhitespaceOverride();
  } catch {
    whitespaceOverride.value = 'auto';
  }
}

/** Persist and apply a new whitespace override (constraint 3 — one global rule). */
export async function setOverride(value: WhitespaceOverride): Promise<void> {
  await setWhitespaceOverride(value);
  whitespaceOverride.value = value;
}

// ---------------------------------------------------------------------------
// Step 1 — assisted removal (tracked changes)
// ---------------------------------------------------------------------------

export async function runRemoval(): Promise<void> {
  if (upgradeRunning.value) return;
  upgradeRunning.value = true;
  upgradeError.value = null;

  try {
    await wordRun(async ctx => {
      const report = await removeLeadingNumbers(whitespaceOverride.value, ctx);
      lastRemovalReport.value = report;
      trackingUnavailable.value = report.disabled;
      // Surface duplicate/gap readings among the cleared tokens (constraint 4).
      sequenceAnomalies.value = detectSequenceAnomalies(report.tokens);
    });
    // Refresh the candidate queue — feature 02 re-scans separately, but reflect
    // any findings cleared in this session.
    removalQueue.value = gatherManuallyTyped();
  } catch (err) {
    upgradeError.value = err instanceof Error ? err.message : 'Removal failed.';
  } finally {
    upgradeRunning.value = false;
  }
}

// ---------------------------------------------------------------------------
// Step 2 — human-driven re-apply
// ---------------------------------------------------------------------------

export async function runReapply(): Promise<void> {
  if (upgradeRunning.value) return;
  const scheme = selectedScheme.value;
  if (!scheme) { upgradeError.value = 'No scheme selected.'; return; }

  const profile = getActiveProfile();
  upgradeRunning.value = true;
  upgradeError.value = null;

  try {
    await wordRun(async ctx => {
      const result = await startListAndApplyOnSelection(scheme, profile, ctx);
      if (!result) {
        upgradeError.value = 'No cleared paragraphs in the selection to number.';
        return;
      }
      // Surface through the shipped review/report panels (reuse, not duplicate).
      lastApplyReport.value = result.report;
      reviewState.value = result.review;
    });
  } catch (err) {
    upgradeError.value = err instanceof Error ? err.message : 'Re-apply failed.';
  } finally {
    upgradeRunning.value = false;
  }
}
