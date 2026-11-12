/**
 * Feature 06 — Numbering: scheme validation and composition helpers.
 * Pure functions — no Office JS dependency, no document access.
 */

import type { NumberingScheme, ListNumberingStyle } from './types.js';

/** Current schema version for export/import compatibility gate. */
export const SCHEMA_VERSION = 1;

/** Maximum list levels supported by Word (0-based, so 9 levels). */
export const MAX_LEVEL_INDEX = 8;

/** Built-in preset ids — reserved; a user scheme importing with one of these gets a fresh id. */
export const BUILTIN_IDS = new Set([
  'preset-uk-sg-commercial',
  'preset-schedule',
  'preset-schedule-part',
  'preset-recitals',
  'preset-plain-outline',
  'preset-us-litigation',
]);

// ---------------------------------------------------------------------------
// Format composition
// ---------------------------------------------------------------------------

/**
 * Compose the final format string passed to setLevelNumbering.
 * When prefix/suffix are empty, returns formatString verbatim (preserving %n tokens).
 * Token integrity (%1..%9) is not altered by this function.
 */
export function composeFormat(
  prefix: string,
  formatString: string,
  suffix: string,
): string {
  if (!prefix && !suffix) return formatString;
  return `${prefix}${formatString}${suffix}`;
}

/**
 * Convert a %N-token format string to the `(string | number)[]` array that
 * Word.List.setLevelNumbering actually accepts.
 *
 * CRITICAL: Word's format array uses **0-based level indices** — the same
 * scale as setLevelNumbering's `level` argument. Our preset notation is
 * 1-based and human-readable (%1 = the first level = level index 0), so each
 * %N token is emitted as the number N-1. Referencing a level *deeper* than the
 * one being set is invalid and throws "The argument you provided is not valid"
 * at ctx.sync() — so %1 MUST become 0, not 1.
 * (Confirmed against the official Word.List.setLevelNumbering reference.)
 *
 * E.g. "%1.%2.%3." → [0, '.', 1, '.', 2, '.']
 *      "(%4)"       → ['(', 3, ')']
 *      "Part %1"    → ['Part ', 0]
 */
export function formatStringToArray(s: string): (string | number)[] {
  const parts: (string | number)[] = [];
  const re = /%(\d)/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = re.exec(s)) !== null) {
    if (match.index > lastIndex) {
      parts.push(s.slice(lastIndex, match.index));
    }
    const digit = match[1] ?? '1';
    // %N (1-based, human notation) → N-1 (0-based level index for Word).
    parts.push(parseInt(digit, 10) - 1);
    lastIndex = re.lastIndex;
  }
  if (lastIndex < s.length) {
    parts.push(s.slice(lastIndex));
  }
  return parts;
}


// ---------------------------------------------------------------------------
// ListNumbering mapping
// ---------------------------------------------------------------------------

/**
 * Map a ListNumberingStyle to the capitalized string literal accepted by
 * Word.List.setLevelNumbering (overload 2 of 2 in the Office JS typings).
 * Word's enum member names are TitleCase: "Arabic", "UpperRoman", etc.
 */
const LIST_NUMBERING_MAP: Readonly<Record<ListNumberingStyle, string>> = {
  arabic:      'Arabic',
  upperRoman:  'UpperRoman',
  lowerRoman:  'LowerRoman',
  upperLetter: 'UpperLetter',
  lowerLetter: 'LowerLetter',
  none:        'None',
};

/**
 * Map a ListNumberingStyle string to the Word.ListNumbering string literal.
 * Throws on an unrecognised value.
 */
export function mapToWordListNumbering(style: string): string {
  const mapped = LIST_NUMBERING_MAP[style as ListNumberingStyle];
  if (!mapped) throw new Error(`Unknown listNumbering style: "${style}"`);
  return mapped;
}

// ---------------------------------------------------------------------------
// Live preview
// ---------------------------------------------------------------------------

/**
 * Render what a given level looks like in context, e.g. "1.1.1" or "(a)".
 * Used by the level editor without touching the document.
 *
 * Strategy: replace each %N token with the ordinal value for that level (defaulting
 * to 1 for every level — giving a plausible single-occurrence preview).
 * The previewCounter map allows simulating "2nd item at level 1" etc. for richer previews,
 * but defaults (all 1) are enough for the edit-time preview column.
 */
export function previewLevel(
  scheme: NumberingScheme,
  levelIndex: number,
  counters?: Readonly<Record<number, number>>,
): string {
  const level = scheme.levels.find(l => l.levelIndex === levelIndex);
  if (!level) return '';

  const format = composeFormat(level.prefix, level.formatString, level.suffix);
  return format.replace(/%(\d)/g, (_match, nStr) => {
    const n = parseInt(nStr, 10);
    const targetLevelIndex = n - 1;
    const count = counters?.[targetLevelIndex] ?? 1;
    const targetLevel = scheme.levels.find(l => l.levelIndex === targetLevelIndex);
    if (!targetLevel) return '';
    return renderCounter(count, targetLevel.listNumbering);
  });
}

function renderCounter(n: number, style: ListNumberingStyle): string {
  switch (style) {
    case 'arabic':      return String(n);
    case 'upperRoman':  return toRoman(n).toUpperCase();
    case 'lowerRoman':  return toRoman(n).toLowerCase();
    case 'upperLetter': return toLetter(n).toUpperCase();
    case 'lowerLetter': return toLetter(n).toLowerCase();
    case 'none':        return '';
  }
}

function toRoman(n: number): string {
  const pairs: [number, string][] = [
    [1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],
    [50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I'],
  ];
  let result = '';
  for (const [val, sym] of pairs) {
    while (n >= val) { result += sym; n -= val; }
  }
  return result;
}

function toLetter(n: number): string {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

export interface ValidationResult {
  valid: boolean;
  errors: string[];
}

/**
 * Validate a NumberingScheme before saving or applying.
 * Rules per spec §7.2:
 *  1. Non-empty name.
 *  2. Levels contiguous from 0.
 *  3. Each %N token in formatString references a defined level (levelIndex = N-1 must exist).
 *  4. Brackets in formatString must balance.
 *  5. schemaVersion must be present and <= SCHEMA_VERSION.
 */
export function validateScheme(scheme: NumberingScheme): ValidationResult {
  const errors: string[] = [];

  if (!scheme.name || scheme.name.trim() === '') {
    errors.push('Scheme name must not be empty.');
  }

  if (!scheme.levels || scheme.levels.length === 0) {
    errors.push('A scheme must have at least one level.');
    return { valid: false, errors };
  }

  // Contiguous from 0
  const indices = [...scheme.levels].map(l => l.levelIndex).sort((a, b) => a - b);
  for (let i = 0; i < indices.length; i++) {
    if (indices[i] !== i) {
      errors.push(`Levels must be contiguous from 0. Found indices: ${indices.join(', ')}.`);
      break;
    }
  }
  const maxIndex = indices[indices.length - 1] ?? 0;
  if (maxIndex > MAX_LEVEL_INDEX) {
    errors.push(`Level index ${maxIndex} exceeds the maximum (${MAX_LEVEL_INDEX}).`);
  }

  const definedLevelIndices = new Set(indices);

  for (const level of scheme.levels) {
    const format = composeFormat(level.prefix, level.formatString, level.suffix);

    // Token references
    const tokenMatches = format.match(/%(\d)/g) ?? [];
    for (const token of tokenMatches) {
      const digit = token[1] ?? '0';
      const n = parseInt(digit, 10);
      if (n === 0 || n > 9) {
        errors.push(`Level ${level.levelIndex}: token ${token} is out of range (%1–%9).`);
      } else if (!definedLevelIndices.has(n - 1)) {
        errors.push(`Level ${level.levelIndex}: token ${token} references level ${n - 1} which is not defined.`);
      }
    }

    // Bracket balance
    let depth = 0;
    for (const ch of format) {
      if (ch === '(') depth++;
      else if (ch === ')') depth--;
      if (depth < 0) { errors.push(`Level ${level.levelIndex}: unmatched closing parenthesis in format.`); break; }
    }
    if (depth > 0) errors.push(`Level ${level.levelIndex}: unmatched opening parenthesis in format.`);
  }

  if (typeof scheme.schemaVersion !== 'number') {
    errors.push('schemaVersion must be a number.');
  } else if (scheme.schemaVersion > SCHEMA_VERSION) {
    errors.push(`schemaVersion ${scheme.schemaVersion} is newer than this version (${SCHEMA_VERSION}); update the add-in.`);
  }

  return { valid: errors.length === 0, errors };
}

// ---------------------------------------------------------------------------
// Duplicate helper
// ---------------------------------------------------------------------------

/**
 * Create an editable copy of a scheme: fresh id, builtIn:false, name suffixed " (copy)".
 * The duplicate is always assignable as a user scheme.
 */
export function duplicateScheme(scheme: NumberingScheme, newId: string): NumberingScheme {
  return {
    ...scheme,
    id: newId,
    name: `${scheme.name} (copy)`,
    builtIn: false,
  };
}
