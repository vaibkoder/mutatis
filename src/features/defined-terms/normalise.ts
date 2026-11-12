/**
 * Term normalisation, lookup, and the longest-match alternation builder.
 * Pure functions — no Office JS, no document access. (spec §4.1)
 */

// Curly / guillemet / prime variants → ASCII quotes.
const DOUBLE_QUOTES = /[“”„«»″]/g;
const SINGLE_QUOTES = /[‘’‚‹›′]/g;

/** Fold smart/guillemet quotes to ASCII `"` and `'`. */
export function foldQuotes(s: string): string {
  return s.replace(DOUBLE_QUOTES, '"').replace(SINGLE_QUOTES, "'");
}

/**
 * Clean a raw candidate into its canonical surface form (case preserved):
 * fold quotes → collapse whitespace → strip surrounding quotes → strip possessive.
 */
export function cleanTerm(raw: string): string {
  let s = foldQuotes(raw);
  s = s.replace(/\s+/g, ' ').trim();
  s = s.replace(/^["']+/, '').replace(/["']+$/, '').trim();
  s = s.replace(/'s$/i, '').trim();
  return s;
}

/**
 * Canonical normaliser → the lowercase map KEY (spec §4.1).
 * Does NOT fold en-GB↔en-US spelling: `Authorised` ≠ `Authorized`.
 */
export function normaliseTerm(raw: string): string {
  return cleanTerm(raw).toLowerCase();
}

/** Naive singular candidates for a plural key, longest-stripped first. */
export function singularForms(key: string): string[] {
  const out: string[] = [];
  if (key.endsWith('ies') && key.length > 3) out.push(key.slice(0, -3) + 'y');
  if (key.endsWith('es') && key.length > 2) out.push(key.slice(0, -2));
  if (key.endsWith('s') && key.length > 1) out.push(key.slice(0, -1));
  return out;
}

/**
 * Resolve a raw candidate to an EXISTING key in `map` (spec §4.1 lookup rules):
 * exact → possessive-stripped (already in normaliseTerm) → naive singular.
 * Never invents a key; returns null when nothing matches. Pluralisation is
 * resolved here at lookup time, never destructively at storage time.
 */
export function resolveLookup(map: ReadonlyMap<string, unknown>, raw: string): string | null {
  const key = normaliseTerm(raw);
  if (map.has(key)) return key;
  for (const cand of singularForms(key)) {
    if (map.has(cand)) return cand;
  }
  return null;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Build a longest-match-first alternation over the given surface forms (spec §4.1).
 * Forms are sorted by descending length so `"Completion Date"` is consumed before
 * `"Completion"`. Returned regex is case-insensitive + global with word boundaries.
 * Returns null when there are no forms (an empty alternation would match everything).
 */
export function buildAlternation(surfaceForms: readonly string[]): RegExp | null {
  const cleaned = surfaceForms
    .map(f => cleanTerm(f))
    .filter(f => f.length > 0);
  if (cleaned.length === 0) return null;
  const unique = Array.from(new Set(cleaned));
  unique.sort((a, b) => b.length - a.length);
  const alt = unique.map(escapeRegExp).join('|');
  // \b is unreliable next to non-word glyphs (e.g. "S$"); guard with lookarounds
  // that simply forbid an adjacent ASCII letter/digit, so "Closing" inside
  // "Closings" still needs the plural lookup rather than a partial hit.
  return new RegExp(`(?<![A-Za-z0-9])(?:${alt})(?![A-Za-z0-9])`, 'gi');
}
