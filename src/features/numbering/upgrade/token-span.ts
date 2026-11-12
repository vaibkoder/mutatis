/**
 * Feature 06 — Numbering upgrade (ADR-0023): leading-token span detector.
 *
 * Pure, offset-preserving. Given a paragraph's plain text, computes the exact
 * character range to delete for the hand-typed leading number prefix, honouring
 * the whitespace auto-rule + global override (ADR-0023 constraint 3).
 *
 * Constraints implemented here:
 *  - #2 Batch guard: returns null when no leading static-number token is detected
 *    (a swept-in hanging paragraph is therefore skipped, never blanked).
 *  - #3 Whitespace: span covers leading whitespace + number + delimiter + trailing
 *    horizontal whitespace; the global override trims how much trailing whitespace
 *    survives (delete-only — Step 1 never inserts).
 *  - #5 Leading-occurrence-only: anchored at the paragraph start; a later in-body
 *    occurrence ("1.5 … see 1.5 above") is never matched.
 *
 * Reuses the canonical leading-number regex vocabulary exported from feature 02's
 * detector. The feature-02 detector trims first and
 * so loses offsets; this module keeps them, which is what a precise delete needs.
 */

import {
  DECIMAL_DOTTED_RE,
  DECIMAL_BARE_RE,
  LETTER_RE,
  ROMAN_RE,
  UPPER_RE,
} from '../../cross-reference/static-numbering.js';
import type { NumberingWhitespaceOverride } from '@shared/config/index.js';

/** Global whitespace override (ADR-0023 constraint 3). `auto` == strip-to-flush. */
export type WhitespaceOverride = NumberingWhitespaceOverride;

export const WHITESPACE_OVERRIDES: readonly WhitespaceOverride[] = [
  'auto',
  'collapse-to-one-space',
  'strip-to-flush',
  'keep-one-tab',
];

/** A resolved deletion range for one paragraph. */
export interface RemovalSpan {
  /** Inclusive start offset to delete (0 — leading whitespace is included). */
  start: number;
  /** Exclusive end offset to delete (after the override is applied). */
  end: number;
  /** The typed number as it appears, e.g. "1.1", "(a)", "(iv)". */
  token: string;
  /** The delimiter character that followed the number (".", ")", tab or space). */
  delimiter: string;
}

// Horizontal whitespace we treat as trailing run: space, tab, NBSP, figure space,
// (Vertical/newline never appears mid-paragraph.)
const TRAILING_WS_RE = /[ \t\u00A0\u2007\u2009\u202F]/;

/** Patterns in precedence order — dotted decimal first, then bare, then limbs. */
const PATTERNS: readonly RegExp[] = [
  DECIMAL_DOTTED_RE,
  DECIMAL_BARE_RE,
  LETTER_RE,
  ROMAN_RE,
  UPPER_RE,
];

interface SpanParts {
  /** End offset of the number+delimiter match (one delimiter char included). */
  delimEnd: number;
  /** End offset after consuming the trailing horizontal whitespace run. */
  trailEnd: number;
  token: string;
  delimiter: string;
}

/** Locate the leading number, its delimiter, and trailing whitespace run. */
function locateParts(text: string): SpanParts | null {
  for (const re of PATTERNS) {
    const m = re.exec(text);
    if (!m || m.index !== 0) continue;

    const full = m[0];
    const leadingWsLen = (/^\s*/.exec(full)?.[0].length) ?? 0;
    const delimEnd = full.length;
    const delimStart = delimEnd - 1;
    const token = text.slice(leadingWsLen, delimStart);
    const delimiter = text.slice(delimStart, delimEnd);

    // Extend across any further trailing horizontal whitespace.
    let trailEnd = delimEnd;
    while (trailEnd < text.length && TRAILING_WS_RE.test(text[trailEnd] ?? '')) {
      trailEnd++;
    }

    return { delimEnd, trailEnd, token, delimiter };
  }
  return null;
}

/**
 * Compute the deletion range for the leading static-number prefix, or null when
 * no leading token is present (batch guard — constraint 2).
 */
export function computeRemovalSpan(
  text: string,
  override: WhitespaceOverride = 'auto',
): RemovalSpan | null {
  const parts = locateParts(text);
  if (!parts) return null;

  const { delimEnd, trailEnd, token, delimiter } = parts;
  const hasTrailingWs = trailEnd > delimEnd;

  let end: number;
  switch (override) {
    case 'collapse-to-one-space':
      // Keep exactly one trailing whitespace char if any existed; else flush.
      end = hasTrailingWs ? trailEnd - 1 : trailEnd;
      break;
    case 'keep-one-tab': {
      // Keep from the first tab in the trailing run (common "1.\tBody" case);
      // otherwise fall back to keeping one whitespace char / flush.
      let kept = -1;
      for (let i = delimEnd; i < trailEnd; i++) {
        if (text[i] === '\t') { kept = i; break; }
      }
      end = kept >= 0 ? kept : (hasTrailingWs ? trailEnd - 1 : trailEnd);
      break;
    }
    case 'strip-to-flush':
    case 'auto':
    default:
      end = trailEnd;
      break;
  }

  // start is always 0 — leading whitespace before the number is removed so body
  // text ends flush (the real list supplies its own indent). constraint 3.
  return { start: 0, end, token, delimiter };
}
