// src/features/repair/guards.ts
import type { StyleProfile } from './profiles.js';

export interface DocumentHeuristics {
  /** Detected line spacing multiple (e.g. 2.0 for double-spaced). */
  detectedLineSpacing: number | null;
  /** Detected body margin in points (smallest of left/right). */
  detectedMarginPt: number | null;
  /** Detected body font size. */
  detectedFontSize: number | null;
  /** Detected body font name. */
  detectedFontName: string | null;
}

/**
 * Heuristic helper — returns true when the detected document properties
 * look like standard court formatting:
 *   - double-spaced (≥ 1.9 multiple)
 *   - margins ≥ 70pt (~25mm, the UK minimum; 35mm for SG is 99pt)
 *   - Times New Roman 12
 */
function looksCourtFormatted(h: DocumentHeuristics): boolean {
  const isDoubleSpaced =
    h.detectedLineSpacing !== null && h.detectedLineSpacing >= 1.9;
  const isCourtMargins =
    h.detectedMarginPt !== null && h.detectedMarginPt >= 70.0;
  const isCourtFont =
    (h.detectedFontName?.includes('Times') ?? false) &&
    h.detectedFontSize === 12;
  return isDoubleSpaced && isCourtMargins && isCourtFont;
}

/**
 * Returns the ids of normalisation-related findings that should be unchecked
 * when the document already matches court profile expectations, plus an
 * optional warning message to surface in the UI.
 *
 * spec §7.2 court-profile guard
 *
 * Only fires when `profile.kind === 'court'`. If the document already looks
 * court-formatted (double-spaced + ≥25mm margins + TNR12), the
 * `spacing.anomaly` finding is unchecked so the user does not inadvertently
 * re-normalise a document that is already correct.
 */
export function getNormalisationWarnings(
  profile: StyleProfile,
  heuristics: DocumentHeuristics,
): { uncheckIds: string[]; warningMessage: string | null } {
  if (profile.kind !== 'court') {
    return { uncheckIds: [], warningMessage: null };
  }

  if (looksCourtFormatted(heuristics)) {
    return {
      uncheckIds: ['spacing.anomaly'],
      warningMessage:
        'Document appears to already match court formatting. Normalisation has been unchecked.',
    };
  }

  return { uncheckIds: [], warningMessage: null };
}

/**
 * Warn when a commercial profile is selected on a court-looking document.
 *
 * Returns a human-readable warning string when:
 *   - `profile.kind === 'commercial'`, AND
 *   - the document appears court-formatted (double-spaced + ≥25mm + TNR12).
 *
 * Returns null otherwise.
 */
export function getCommercialOnCourtWarning(
  profile: StyleProfile,
  heuristics: DocumentHeuristics,
): string | null {
  if (profile.kind !== 'commercial') return null;

  if (looksCourtFormatted(heuristics)) {
    return (
      'This document appears to be formatted for court filing ' +
      '(double-spaced, Times New Roman 12, wide margins). ' +
      'Applying a commercial profile may remove required court formatting. ' +
      'Confirm you want to proceed.'
    );
  }

  return null;
}
