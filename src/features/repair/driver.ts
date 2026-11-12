// src/features/repair/driver.ts
//
// Preact Signals driver for Feature 05 — Document Repair.
// Holds all UI state for the RepairPanel and the two async orchestrator calls.

import { signal, computed } from '@preact/signals';
import type { RepairReport } from './types.js';
import type { StyleProfile } from './profiles.js';
import { getProfile, listProfileIds, SG_COURT_FILING } from './profiles.js';
import { scan } from './scan.js';
import { apply } from './apply.js';
import type { ApplyOptions } from './apply.js';
import { getNormalisationWarnings, getCommercialOnCourtWarning } from './guards.js';
import type { DocumentHeuristics } from './guards.js';
import type { DocxBundle } from '@shared/ooxml/index.js';

// ---------------------------------------------------------------------------
// Signals
// ---------------------------------------------------------------------------

export const repairProfile = signal<StyleProfile>(SG_COURT_FILING);
export const repairScanReport = signal<RepairReport | null>(null);
export const repairApplyReport = signal<RepairReport | null>(null);
export const repairBusy = signal<boolean>(false);
export const repairError = signal<string | null>(null);

/** Set of finding ids the user has checked (selected for apply). */
export const repairChecked = signal<Set<string>>(new Set());

/** Set of non-protected section-break ids the user explicitly confirmed. */
export const repairConfirmedSectionBreaks = signal<Set<string>>(new Set());

/**
 * Set of orientation-changing (protected:true) section-break ids the user
 * confirmed via the per-break orientation modal. Kept as a separate channel
 * so the apply layer can enforce the hard rule (spec §5.7).
 */
export const repairConfirmedProtectedBreaks = signal<Set<string>>(new Set());

export const repairNormalisationWarning = signal<string | null>(null);
export const repairCommercialOnCourtWarning = signal<string | null>(null);

// ---------------------------------------------------------------------------
// Derived signals
// ---------------------------------------------------------------------------

/** Findings from the active report grouped by severity, or null when no report. */
export const repairGroupedFindings = computed(() => {
  const report = repairScanReport.value ?? repairApplyReport.value;
  if (!report) return null;
  const confidentiality = report.findings.filter(f => f.severity === 'confidentiality');
  const stabilisation  = report.findings.filter(f => f.severity === 'stabilisation');
  const info           = report.findings.filter(f => f.severity === 'info');
  return { confidentiality, stabilisation, info };
});

// ---------------------------------------------------------------------------
// Module-level bundle cache
// ---------------------------------------------------------------------------

/**
 * Holds the DocxBundle produced by scan() so that apply() can re-use the
 * already-read ZIP without a second Office round-trip.
 * Cleared when a new scan starts.
 */
let _scanBundle: DocxBundle | null = null;

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

/** Select a profile by id; silently ignores unknown ids. */
export function setProfile(profileId: string): void {
  const p = getProfile(profileId);
  if (p) repairProfile.value = p;
}

/**
 * Run the full document scan and populate the scan report signal.
 * Auto-populates repairChecked with selectedByDefault findings.
 *
 * @param documentName  Display name shown in the report header.
 */
export async function runScan(documentName: string): Promise<void> {
  repairBusy.value = true;
  repairError.value = null;
  _scanBundle = null;

  try {
    const state = await scan(documentName, repairProfile.value, new Date().toISOString());

    _scanBundle = state.bundle;
    repairScanReport.value  = state.report;
    repairApplyReport.value = null;

    // Court-profile guards (spec §7.2): may uncheck normalisation when the
    // document already matches court formatting, and surface warnings.
    const { uncheckIds } = _updateGuardWarnings(repairProfile.value, state.heuristics);

    // Default selection: all findings that are selectedByDefault and NOT
    // destructive, minus any the court guard says to uncheck.
    const defaults = new Set(
      state.report.findings
        .filter(f => f.selectedByDefault && !f.destructive && !f.requiresConfirm)
        .filter(f => !uncheckIds.includes(f.id))
        .map(f => f.id),
    );
    repairChecked.value                  = defaults;
    repairConfirmedSectionBreaks.value   = new Set();
    repairConfirmedProtectedBreaks.value = new Set();
  } catch (err) {
    repairError.value = err instanceof Error ? err.message : String(err);
  } finally {
    repairBusy.value = false;
  }
}

/** Toggle a finding's checked state. Silently ignored for destructive findings. */
export function toggleFinding(id: string): void {
  const next = new Set(repairChecked.value);
  if (next.has(id)) {
    next.delete(id);
  } else {
    next.add(id);
  }
  repairChecked.value = next;
}

/**
 * Confirm a NON-protected destructive / requires-confirm finding for removal.
 * For orientation-changing (protected:true) breaks use confirmProtectedBreak,
 * which routes through a separate, second confirmation channel (spec §5.7).
 */
export function confirmSectionBreak(id: string): void {
  const next = new Set(repairConfirmedSectionBreaks.value);
  next.add(id);
  repairConfirmedSectionBreaks.value = next;
  // Also add to checked so the apply orchestrator picks it up.
  const checked = new Set(repairChecked.value);
  checked.add(id);
  repairChecked.value = checked;
}

/**
 * Confirm an ORIENTATION-CHANGING (protected:true) section break for removal.
 * Hard rule (spec §5.7): such a break is only ever removed when its id is
 * present in BOTH the confirmed set and the protected-confirmed set, so this
 * adds the id to both channels (plus checked). The UI must gate this behind a
 * dedicated modal warning that the Schedule reverts to portrait.
 */
export function confirmProtectedBreak(id: string): void {
  const protectedSet = new Set(repairConfirmedProtectedBreaks.value);
  protectedSet.add(id);
  repairConfirmedProtectedBreaks.value = protectedSet;

  const confirmedSet = new Set(repairConfirmedSectionBreaks.value);
  confirmedSet.add(id);
  repairConfirmedSectionBreaks.value = confirmedSet;

  const checked = new Set(repairChecked.value);
  checked.add(id);
  repairChecked.value = checked;
}

/**
 * Apply the checked findings to the document using the cached bundle from scan.
 * Replaces the scan report with an apply-phase report.
 */
export async function runApply(): Promise<void> {
  const scanReport = repairScanReport.value;
  if (!scanReport || !_scanBundle) return;

  repairBusy.value = true;
  repairError.value = null;

  try {
    const options: ApplyOptions = {
      selectedFindingIds:                 [...repairChecked.value],
      confirmedSectionBreakIds:           [...repairConfirmedSectionBreaks.value],
      confirmedProtectedSectionBreakIds:  [...repairConfirmedProtectedBreaks.value],
      blankParaMode:                      'collapseMultiple',
      applyMargins:                       false,   // conservative default
      stripInlineFormatting:              false,   // conservative default
    };

    const applyReport = await apply(
      scanReport,
      _scanBundle,
      repairProfile.value,
      options,
      new Date().toISOString(),
    );

    repairApplyReport.value = applyReport;
    repairScanReport.value  = null;
    _scanBundle = null;
  } catch (err) {
    repairError.value = err instanceof Error ? err.message : String(err);
  } finally {
    repairBusy.value = false;
  }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Update the guard-warning signals from the detected document heuristics and
 * the active profile (spec §7.2). Returns the finding ids the court guard wants
 * unchecked so the caller can exclude them from the default selection.
 */
function _updateGuardWarnings(
  profile: StyleProfile,
  heuristics: DocumentHeuristics,
): { uncheckIds: string[] } {
  const { uncheckIds, warningMessage } = getNormalisationWarnings(profile, heuristics);
  repairNormalisationWarning.value     = warningMessage;
  repairCommercialOnCourtWarning.value = getCommercialOnCourtWarning(profile, heuristics);
  return { uncheckIds };
}

// Re-export profile helpers for convenience
export { listProfileIds, getProfile };
