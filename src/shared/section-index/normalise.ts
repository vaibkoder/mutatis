/**
 * Shared number-normalisation for the Section Index.
 *
 * Used symmetrically by the builder (4.3) and the reference scanner (5.1.3)
 * so that "Clause 4.2.1", "clause 4.2.1", and a REF field result all collapse
 * to the same key fragment `4.2.1`, and `(B)` → `(b)` for letters.
 *
 * §4.4 normalisation rules:
 * - Decimal:          `4`, `4.2`, `4.2.1` → unchanged
 * - Lettered sub:     `(a)`, `a)`, `a.`   → `(a)` (parens kept, lowercase)
 * - Roman sub:        `(i)`, `(iv)`, etc.  → `(i)` (parens kept, lowercase)
 * - Recital:          `(A)`, `A.`          → `A`   (uppercase, no parens — keys like `recital:B`)
 * - Part roman:       `I`, `II`, `III`     → uppercase roman  (keys like `part:II`)
 * - Schedule/Annex:   `1`, `A`, `B`        → as-rendered, letters uppercase
 */

/** Purely alphabetic roman numerals (no `i`/`v`/`x` ambiguity guard here — see disambiguateRoman). */
const ROMAN_RE = /^(i{1,3}|iv|vi{0,3}|ix|xi{0,2}|xii|xiii|xiv|xv|xvi|xvii|xviii|xix|xx|xxx|xl|li{0,3}|lv|lx{0,3}|xc|ci{0,3}|cd|cm|[mdclxvi]+)$/i;

/** Single lowercase letter a–z. */
const LOWER_LETTER_RE = /^[a-z]$/;

/** Single uppercase letter A–Z. */
const UPPER_LETTER_RE = /^[A-Z]$/;

/** Decimal number with optional dots. */
const DECIMAL_RE = /^\d+(\.\d+)*$/;

/**
 * Strip surrounding parens and trailing punctuation from a raw listString token.
 * `(b)` → `b`; `4.2.` → `4.2`; `B.` → `B`.
 */
export function stripParens(raw: string): string {
  return raw.replace(/^[([]+/, '').replace(/[)\].]+$/, '').trim();
}

/**
 * Normalise a sub-clause token (letter or roman) into the canonical parens form.
 * Used for lettered and roman sub-clause fragments that form part of a clause key.
 *
 * Canonical form: `(a)`, `(i)`, `(iv)` — lowercase, parens kept.
 */
export function normaliseSubToken(raw: string): string {
  const s = raw.trim();
  // Already in paren form like "(b)" or "(iv)"
  const inner = s.startsWith('(') && s.endsWith(')')
    ? s.slice(1, -1).toLowerCase()
    : stripParens(s).toLowerCase();
  return `(${inner})`;
}

/**
 * Full normalisation for a rendered number string coming from prose or a list token.
 *
 * `context` indicates how to interpret an ambiguous token:
 *  - `'clause-sub'` — expects letter/roman sub-clause: normalise to `(x)` form
 *  - `'clause-body'` — expects decimal clause number: leave as-is
 *  - `'recital'` — uppercase letter (no parens in key, just the letter itself)
 *  - `'part'` — uppercase roman
 *  - `'schedule-or-annex'` — as-rendered, letters uppercased
 *  - `'auto'` — infer from token shape (default, used for prose scanning)
 */
export type NormaliseContext =
  | 'clause-sub'
  | 'clause-body'
  | 'recital'
  | 'part'
  | 'schedule-or-annex'
  | 'auto';

export function normaliseNumber(raw: string, context: NormaliseContext = 'auto'): string {
  const s = raw.trim();
  const inner = stripParens(s);

  if (context === 'recital') {
    // Recital letter: uppercase, no parens.
    return inner.toUpperCase();
  }

  if (context === 'part') {
    // Part roman: uppercase.
    return inner.toUpperCase();
  }

  if (context === 'schedule-or-annex') {
    // Letters uppercase, numbers as-is.
    return UPPER_LETTER_RE.test(inner) || LOWER_LETTER_RE.test(inner)
      ? inner.toUpperCase()
      : inner;
  }

  if (context === 'clause-sub') {
    // Sub-clause: normalise to (x) form, lowercase.
    return normaliseSubToken(s);
  }

  if (context === 'clause-body') {
    return DECIMAL_RE.test(inner) ? inner : inner;
  }

  // 'auto': infer from shape.

  // Compound clause token like "4.2(b)(i)" — return as-is (lowercase), do not strip.
  const COMPOUND_CLAUSE_RE = /^\d+(?:\.\d+)*(?:\([a-z]+\)|\([ivxlcdm]+\))+$/i;
  if (COMPOUND_CLAUSE_RE.test(s)) {
    return s.toLowerCase();
  }

  if (DECIMAL_RE.test(inner)) {
    // Decimal clause number: unchanged.
    return inner;
  }

  // Single uppercase letter: could be a recital or schedule — return uppercase.
  if (UPPER_LETTER_RE.test(inner)) {
    return inner;
  }

  // Single lowercase letter or roman: sub-clause form.
  if (LOWER_LETTER_RE.test(inner) || ROMAN_RE.test(inner.toLowerCase())) {
    return `(${inner.toLowerCase()})`;
  }

  // Already in paren form like "(b)": normalise to lowercase paren.
  if (s.startsWith('(') && s.endsWith(')')) {
    return `(${s.slice(1, -1).toLowerCase()})`;
  }

  return inner;
}

/**
 * Determine whether a token (already stripped of parens) is roman-numeral-shaped
 * and ambiguous with a single letter (`i`, `v`, `x`, `l`, `c`, `d`, `m`).
 * The caller uses sibling context to resolve the ambiguity.
 */
export function isAmbiguousRomanLetter(token: string): boolean {
  const t = token.toLowerCase();
  return t.length === 1 && 'ivxlcdm'.includes(t);
}

/**
 * Determine the interpretation of a sub-clause token given its siblings.
 *
 * `prevSiblingToken` is the normalised previous sibling at the same list level
 * (or null if this is the first item). Returns `'roman'` or `'letter'`.
 *
 * Heuristic: if the previous sibling is clearly a multi-character roman (`ii`, `iii`, `iv`)
 * or the previous sibling ended a roman sequence (only romans in the sequence so far),
 * this is roman. If the previous sibling is a plain alphabetic letter *not* a valid roman
 * (like `g`, `h`), this is a letter.
 *
 * When uncertain (e.g. first item is `(i)`), returns `'roman'` (the more common reading
 * in SG/UK drafting) and marks the caller to set `ambiguousRomanLetter = true`.
 */
export function disambiguateRoman(
  token: string,
  prevSiblingToken: string | null,
): { interpretation: 'roman' | 'letter'; certain: boolean } {
  if (!prevSiblingToken) {
    return { interpretation: 'roman', certain: false };
  }
  const prev = prevSiblingToken.replace(/[()]/g, '').toLowerCase();
  // Previous sibling is multi-char roman like "ii", "iii", "iv" → current is roman.
  if (/^(ii|iii|iv|vi{1,3}|ix|xi{1,2}|xii|xiii|xiv|xv)$/.test(prev)) {
    return { interpretation: 'roman', certain: true };
  }
  // Previous sibling is a non-roman letter like g, h, j, k, n... → current is letter.
  if (/^[b-hj-np-uw-z]$/.test(prev)) {
    return { interpretation: 'letter', certain: true };
  }
  // Previous sibling is single ambiguous char too — assume sequence is consistent.
  return { interpretation: 'roman', certain: false };
}
