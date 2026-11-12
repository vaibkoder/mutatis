/**
 * host-web — STATELESS list-number formatters (spec/features/12 §4 hard-part #1).
 *
 * Two pure, counter-free functions the (later) `renderListNumbers` state machine
 * composes once it knows each level's current counter value:
 *   - `formatInt(n, numFmt)`  — one integer → its rendered token in a given OOXML
 *     `<w:numFmt>` style.
 *   - `applyLvlText(lvlText, perLevelTokens)` — substitute `%1..%9` in a
 *     `<w:lvlText>` template with already-formatted per-level tokens.
 *
 * Host-NEUTRAL: no `Office.*`/`Word.*` token.
 */

/**
 * Render a single counter value `n` in the given OOXML `<w:numFmt w:val>` style.
 *
 * Supported (the legal-drafting set):
 *   - `decimal`            → "1", "2", … "10"
 *   - `decimalZero`        → 0-padded to a minimum width of 2: "01", "09", "10", "100"
 *   - `lowerLetter`        → bijective base-26: 1→"a", 26→"z", 27→"aa", 28→"ab", …
 *   - `upperLetter`        → as lowerLetter, upper-cased
 *   - `lowerRoman`         → "i", "ii", "iv", "ix", …
 *   - `upperRoman`         → as lowerRoman, upper-cased
 *
 * Everything else — `bullet`, `none`, or any unrecognised value — returns `''`
 * (a bullet has no rendered *number*; the bullet glyph lives in `lvlText` literal
 * text and survives `applyLvlText` untouched). Documented, intentional.
 *
 * `n <= 0` is out of range for letter/roman sequences and returns `''` for those
 * formats; `decimal`/`decimalZero` still render the raw integer.
 */
export function formatInt(n: number, numFmt: string): string {
  switch (numFmt) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      return String(n).padStart(2, '0');
    case 'lowerLetter':
      return toAlpha(n);
    case 'upperLetter':
      return toAlpha(n).toUpperCase();
    case 'lowerRoman':
      return toRoman(n);
    case 'upperRoman':
      return toRoman(n).toUpperCase();
    default:
      return '';
  }
}

/**
 * Can this `numFmt` ever put a number on the page?
 *
 * `false` for `bullet`, `none` and any value this build does not implement — the
 * three cases `formatInt` answers `''` to for EVERY counter value. Derived from
 * `formatInt` itself rather than from a second list of format names, so the two
 * cannot drift apart: a format added to the switch is supported here the same
 * day.
 *
 * A level that substitutes such a format occupies a `%N` slot in the template
 * and contributes nothing visible, which matters wherever code asks *"which
 * levels does this label actually show?"* — `ParaNumbering.renders`, and through
 * it §6.1's cascade header (ADR-0042).
 *
 * The residue this does not cover is value-dependent and stated rather than
 * hidden: `lowerLetter`/`upperLetter`/`lowerRoman`/`upperRoman` also render `''`
 * for a counter at or below zero, which needs `<w:start w:val="0"/>` or lower —
 * and `0` empties only the FIRST item of that level, so a header that
 * contradicts its rows needs `start` at `-1` or below (constructed, and it does).
 * A predicate over the DEFINITION cannot see a counter, and moving it to render
 * time would make the answer differ between the two sides of a comparison, which
 * is the one property the caller relies on.
 */
export function rendersToken(numFmt: string): boolean {
  return formatInt(1, numFmt) !== '';
}

/**
 * Substitute `%1..%9` placeholders in an OOXML `<w:lvlText>` template with the
 * already-formatted token for each 1-based level: `%k` → `perLevelTokens[k-1]`.
 * Literal text (including bullet glyphs and any `%` not followed by 1-9) is kept
 * verbatim. A placeholder with no supplied token renders as `''`.
 *
 * Example: `applyLvlText('%1.%2', ['4', '2'])` → `'4.2'`.
 */
export function applyLvlText(lvlText: string, perLevelTokens: readonly string[]): string {
  return lvlText.replace(new RegExp(LVL_TOKEN.source, 'g'), (_match, digit: string) => {
    const idx = Number(digit) - 1;
    return perLevelTokens[idx] ?? '';
  });
}

/**
 * Which 0-based levels a `<w:lvlText>` template actually substitutes.
 *
 * `"%1.%2"` → `{0, 1}`; `"(%2)"` → `{1}`; `"Article"` → `{}` (a literal label
 * draws on no counter at all).
 *
 * Exists so a caller can tell which level definitions a rendered label depends
 * on WITHOUT re-parsing the template itself. `applyLvlText` above and this
 * function share `LVL_TOKEN`, so the two cannot disagree about what a
 * placeholder is — and a caller that keyed on its own regex would eventually
 * disagree with the renderer, which is a change either invented or missed.
 */
export function levelsReferencedBy(lvlText: string): ReadonlySet<number> {
  const out = new Set<number>();
  for (const match of lvlText.matchAll(new RegExp(LVL_TOKEN.source, 'g'))) {
    const digit = match[1];
    if (digit !== undefined) out.add(Number(digit) - 1);
  }
  return out;
}

/**
 * A `%N` level placeholder in a `<w:lvlText>` template. ECMA-376 allows `%1`
 * through `%9`; anything else in the template is literal text.
 */
const LVL_TOKEN = /%([1-9])/;

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

/** Bijective base-26: 1→a, 26→z, 27→aa, 28→ab, … ; n<=0 → ''. */
function toAlpha(n: number): string {
  if (n <= 0) return '';
  let x = n;
  let s = '';
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(97 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

const ROMAN_TABLE: ReadonlyArray<readonly [number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

/** Lower-case Roman numerals; n<=0 → '' (Roman has no zero/negative). */
function toRoman(n: number): string {
  if (n <= 0) return '';
  let x = n;
  let s = '';
  for (const [value, sym] of ROMAN_TABLE) {
    while (x >= value) {
      s += sym;
      x -= value;
    }
  }
  return s;
}
