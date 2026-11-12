/**
 * Precedent / Clause Library — save-time validation.
 *
 * Spec: spec/features/07-toc-precedent-templates.md §B.4.1
 *
 * Rules enforced here (pure functions — no Office JS, no DOM):
 *   1. jurisdiction ∈ {SG, EW, OTHER}; if OTHER → jurisdictionOther required.
 *   2. lastReviewed: valid ISO date, not in the future;
 *      warn (not block) when older than DEFAULT_STALENESS_MONTHS.
 *   3. approvedBy non-empty.
 *   4. title non-empty.
 *   5. visibility === 'shared' requires a configured backend.
 *   6. practiceArea non-empty.
 */

import type { ClauseMetadata } from './types.js';
import { DEFAULT_STALENESS_MONTHS } from './types.js';

// ---------------------------------------------------------------------------
// Error / warning types
// ---------------------------------------------------------------------------

export interface ValidationError {
  field: string;
  message: string;
}

export interface ValidationWarning {
  field: string;
  message: string;
}

export interface ValidationResult {
  valid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/** Parse an ISO 8601 date string (YYYY-MM-DD) to a Date object at midnight UTC. */
function parseIsoDate(iso: string): Date | null {
  // Accept YYYY-MM-DD only (spec §B.4.1 stores ISO date)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
  const d = new Date(`${iso}T00:00:00Z`);
  return isNaN(d.getTime()) ? null : d;
}

/** Add N months to a Date (calendar-month arithmetic, UTC). */
function addMonths(date: Date, months: number): Date {
  const result = new Date(date.getTime());
  result.setUTCMonth(result.getUTCMonth() + months);
  return result;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Validate clause metadata at save time.
 *
 * @param meta          The metadata from the save form.
 * @param hasBackend    Whether a configured shared backend is available
 *                      (`getStorageProvider().remote?.isConfigured() === true`).
 * @param now           Optional: override "now" for testing. Defaults to new Date().
 * @param stalenessMonths Optional: override staleness window. Defaults to DEFAULT_STALENESS_MONTHS.
 */
export function validateClauseMetadata(
  meta: Partial<ClauseMetadata>,
  hasBackend: boolean,
  now: Date = new Date(),
  stalenessMonths: number = DEFAULT_STALENESS_MONTHS,
): ValidationResult {
  const errors: ValidationError[] = [];
  const warnings: ValidationWarning[] = [];

  // --- 1. jurisdiction ---
  const jur = meta.jurisdiction;
  if (!jur || !['SG', 'EW', 'OTHER'].includes(jur)) {
    errors.push({ field: 'jurisdiction', message: 'Jurisdiction is required (SG, EW, or Other).' });
  } else if (jur === 'OTHER') {
    const other = meta.jurisdictionOther?.trim() ?? '';
    if (other.length === 0) {
      errors.push({
        field: 'jurisdictionOther',
        message: 'Please specify the jurisdiction (e.g. "New York", "Delaware") when selecting Other.',
      });
    }
  }

  // --- 2. lastReviewed ---
  const lr = meta.lastReviewed?.trim() ?? '';
  if (lr.length === 0) {
    errors.push({ field: 'lastReviewed', message: 'Last reviewed date is required.' });
  } else {
    const reviewed = parseIsoDate(lr);
    if (!reviewed) {
      errors.push({ field: 'lastReviewed', message: 'Last reviewed date must be a valid date (YYYY-MM-DD).' });
    } else {
      // Normalise now to midnight UTC for a fair day-level comparison
      const todayUtc = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
      if (reviewed > todayUtc) {
        errors.push({ field: 'lastReviewed', message: 'Last reviewed date cannot be in the future.' });
      } else {
        // Staleness warning (non-blocking)
        const staleAfter = addMonths(reviewed, stalenessMonths);
        if (staleAfter < todayUtc) {
          warnings.push({
            field: 'lastReviewed',
            message: `This clause was last reviewed more than ${stalenessMonths} months ago. `
              + 'Consider requesting an updated review before sharing.',
          });
        }
      }
    }
  }

  // --- 3. approvedBy ---
  const approvedBy = meta.approvedBy?.trim() ?? '';
  if (approvedBy.length === 0) {
    errors.push({ field: 'approvedBy', message: 'Approved by is required (risk governance).' });
  }

  // --- 4. practiceArea ---
  const practiceArea = meta.practiceArea?.trim() ?? '';
  if (practiceArea.length === 0) {
    errors.push({ field: 'practiceArea', message: 'Practice area is required.' });
  }

  // --- 5. visibility + backend ---
  const visibility = meta.visibility;
  if (!visibility || !['personal', 'shared'].includes(visibility)) {
    errors.push({ field: 'visibility', message: 'Visibility is required (personal or shared).' });
  } else if (visibility === 'shared' && !hasBackend) {
    // This should be blocked in the UI (field disabled); enforce defensively here too.
    errors.push({
      field: 'visibility',
      message: 'Shared visibility requires a configured shared backend (SharePoint or GitHub). '
        + 'No backend is currently configured.',
    });
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

/**
 * Validate the clause title (required, non-empty).
 * Kept separate so the title field can validate independently in the save form.
 */
export function validateTitle(title: string | undefined | null): ValidationError | null {
  const t = title?.trim() ?? '';
  if (t.length === 0) {
    return { field: 'title', message: 'Clause title is required.' };
  }
  return null;
}

/**
 * Validate that a selection is non-empty before the save flow opens.
 * The driver calls this before opening the save form.
 */
export function validateSelectionNonEmpty(html: string, plainText: string): boolean {
  return html.trim().length > 0 || plainText.trim().length > 0;
}
