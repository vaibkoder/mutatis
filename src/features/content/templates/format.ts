/**
 * Template formatters.
 *
 * All date formatting is day-first (DD Month YYYY) per spec CONVENTIONS.md §Dates.
 * No month-first formats (US convention is a bug per invariant checks).
 * ISO 8601 is accepted as input and rendered day-first.
 */

// ---------------------------------------------------------------------------
// Month name table (en-GB)
// ---------------------------------------------------------------------------

const MONTH_NAMES_EN_GB: readonly string[] = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

// ---------------------------------------------------------------------------
// Date formatter
// ---------------------------------------------------------------------------

/**
 * Format a date value (ISO 8601 string, Date object, or "d/m/yyyy" day-first)
 * as "DD Month YYYY" — the mandatory SG/UK display format.
 *
 * Returns null if the input cannot be parsed.
 *
 * Examples:
 *   formatDateDayFirst("2026-06-06") → "06 June 2026"
 *   formatDateDayFirst("06/06/2026") → "06 June 2026"
 *   formatDateDayFirst(new Date(2026, 5, 6)) → "06 June 2026"
 */
export function formatDateDayFirst(input: string | Date): string | null {
  let date: Date;

  if (input instanceof Date) {
    if (isNaN(input.getTime())) return null;
    date = input;
  } else {
    const trimmed = input.trim();

    // ISO 8601: YYYY-MM-DD (or YYYY-MM-DDThh:mm:ss…)
    const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(trimmed);
    if (isoMatch && isoMatch[1] && isoMatch[2] && isoMatch[3]) {
      const year  = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10) - 1; // 0-based
      const day   = parseInt(isoMatch[3], 10);
      date = new Date(year, month, day);
    } else {
      // Little-endian day-first: D/M/YYYY or DD/MM/YYYY (slash or dash)
      const leMatch = /^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})$/.exec(trimmed);
      if (leMatch && leMatch[1] && leMatch[2] && leMatch[3]) {
        const day   = parseInt(leMatch[1], 10);
        const month = parseInt(leMatch[2], 10) - 1;
        const year  = parseInt(leMatch[3], 10);
        date = new Date(year, month, day);
      } else {
        return null;
      }
    }
  }

  if (isNaN(date.getTime())) return null;

  const day   = String(date.getDate()).padStart(2, '0');
  const month = MONTH_NAMES_EN_GB[date.getMonth()];
  const year  = date.getFullYear();

  return `${day} ${month} ${year}`;
}

/**
 * Validate that a date string parses to a real date (used in form validation).
 * Returns null if valid, or an error string if not.
 */
export function validateDate(value: string, label: string): string | null {
  if (!value.trim()) return null; // Required check is handled separately.
  const formatted = formatDateDayFirst(value);
  if (formatted === null) {
    return `${label} must be a valid date (e.g. 06 June 2026 or 2026-06-06).`;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Value formatter (dispatches by variable type)
// ---------------------------------------------------------------------------

/**
 * Format a raw form value for writing back into a content control via insertText.
 * - date → DD Month YYYY
 * - boolean → "Yes" / "No"
 * - everything else → trimmed string
 */
export function formatValueForInsert(
  type: string,
  rawValue: string | boolean | number,
): string {
  if (type === 'date') {
    const str = String(rawValue);
    const formatted = formatDateDayFirst(str);
    return formatted ?? str; // Fallback to the raw string if parsing fails.
  }

  if (type === 'boolean') {
    return rawValue === true || rawValue === 'true' || rawValue === '1' ? 'Yes' : 'No';
  }

  return String(rawValue).trim();
}

// ---------------------------------------------------------------------------
// Registration-number mask helpers
// ---------------------------------------------------------------------------

/**
 * Return a UI hint placeholder for a reg-number input.
 * Reads from config.  Never exposes US-style EIN etc.
 */
export function getRegNumberPlaceholder(scheme: 'UK_CH' | 'SG_UEN'): string {
  if (scheme === 'UK_CH') {
    // 8 alphanumeric chars, two-letter prefix allowed (SC, NI, OC, SO)
    return '00012345 or SC012345';
  }
  // SG_UEN
  return '201912345A or 19990012345K';
}

/**
 * Return the user-facing label for a reg-number scheme.
 */
export function getRegNumberLabel(scheme: 'UK_CH' | 'SG_UEN'): string {
  if (scheme === 'UK_CH') return 'Companies House number (England & Wales)';
  return 'Singapore UEN';
}
