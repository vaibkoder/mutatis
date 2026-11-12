/**
 * Text normalisation for the Compare pipeline.
 * Spec: spec/features/08-compare.md §5.0 (numbering), §5.2 (text normalisation).
 *
 * All functions are pure.
 */

/**
 * §5.2 normalisation used to produce `ParaRecord.normText` (the hashed text):
 *  - NBSP → space.
 *  - Curly quotes (“ ” ‘ ’) → straight " and '.
 *  - Em-dash (—) → " -- "; en-dash (–) → "-".
 *  - Whitespace runs collapsed to a single space; trimmed.
 *  - Case PRESERVED (defined terms / spelling variants are case-sensitive).
 */
export function normaliseText(text: string): string {
  let out = text;

  // Non-breaking space (and the narrow / figure variants commonly seen) → space.
  out = out.replace(/[   ]/g, ' ');

  // Curly double quotes → straight double quote.
  out = out.replace(/[“”]/g, '"');
  // Curly single quotes / apostrophes → straight apostrophe.
  out = out.replace(/[‘’]/g, "'");

  // Em-dash → spaced double hyphen; en-dash → single hyphen.
  // Em-dash first so the spacing it introduces is folded by the collapse below.
  out = out.replace(/—/g, ' -- ');
  out = out.replace(/–/g, '-');

  // Collapse all whitespace runs (incl. tabs / newlines) to a single space, then trim.
  out = out.replace(/\s+/g, ' ').trim();

  return out;
}

/**
 * {@link normaliseText}, plus where each output character came from.
 *
 * `offsets[i]` is the index in `text`'s INPUT of the character that produced
 * output character `i`: the three characters an em-dash expands to all point at
 * the dash, and a collapsed whitespace run is the first character of that run.
 * The replace author uses it to tell which run, field, link or equation a
 * word-diff character belongs to (ADR-0045).
 *
 * It is a second implementation of the same transform, which the reuse-before-build rule
 * calls a defect unless it is held to the first: `text` is asserted equal to
 * `normaliseText(input)` by a property test, and the one caller checks it again
 * and falls back to the text-only rebuild if they ever disagree.
 */
export function normaliseTextWithOffsets(input: string): { text: string; offsets: number[] } {
  // Pass 1 — the one-to-one and one-to-three substitutions, in the same order
  // `normaliseText` applies them. None of them can produce or consume another's
  // input, so a single per-character pass is equivalent to the chained replaces.
  const chars: string[] = [];
  const from: number[] = [];
  for (let i = 0; i < input.length; i++) {
    const c = input[i]!;
    if (c === ' ' || c === ' ' || c === ' ') {
      chars.push(' ');
      from.push(i);
    } else if (c === '“' || c === '”') {
      chars.push('"');
      from.push(i);
    } else if (c === '‘' || c === '’') {
      chars.push("'");
      from.push(i);
    } else if (c === '—') {
      chars.push(' ', '-', '-', ' ');
      from.push(i, i, i, i);
    } else if (c === '–') {
      chars.push('-');
      from.push(i);
    } else {
      chars.push(c);
      from.push(i);
    }
  }

  // Pass 2 — collapse every whitespace run to one space, then trim.
  const outChars: string[] = [];
  const offsets: number[] = [];
  let inSpace = false;
  for (let k = 0; k < chars.length; k++) {
    const c = chars[k]!;
    if (/\s/.test(c)) {
      if (!inSpace) {
        outChars.push(' ');
        offsets.push(from[k]!);
      }
      inSpace = true;
    } else {
      outChars.push(c);
      offsets.push(from[k]!);
      inSpace = false;
    }
  }
  let start = 0;
  let end = outChars.length;
  if (start < end && outChars[start] === ' ') start++;
  if (end > start && outChars[end - 1] === ' ') end--;
  return { text: outChars.slice(start, end).join(''), offsets: offsets.slice(start, end) };
}

/**
 * A more aggressive normalisation: applies {@link normaliseText}, then lowercases
 * and strips all punctuation. Used ONLY to detect the `whitespacePunct` class
 * (spec §5.2 / §6 row 7) — never for hashing the redline.
 */
export function normaliseTextAggressive(text: string): string {
  const base = normaliseText(text).toLowerCase();
  // Strip everything that is not a letter, digit or whitespace, then re-collapse.
  // \p{L}/\p{N} cover non-ASCII letters/digits (Unicode-aware) for SG/UK text.
  const stripped = base.replace(/[^\p{L}\p{N}\s]/gu, '');
  return stripped.replace(/\s+/g, ' ').trim();
}

/**
 * Detect and strip a leading numbering token from paragraph text (spec §5.0).
 * The prefix is removed from the body and returned separately so it can be
 * restored in output and excluded from the body hash (preventing a single
 * inserted clause from making every later paragraph look "changed").
 *
 * Recognised forms (in priority order):
 *  - Compound decimal, optionally with bracketed sub-parts: `4.2 `, `4.2.1 `,
 *    `4.2(a) `, `4.2(a)(i) `.
 *  - Bracketed token: `(a) `, `(iii) `, `(I) `.
 *  - Plain number with a trailing dot: `1. `, `12. `.
 *
 * The returned `prefix` INCLUDES the trailing whitespace that separated it from
 * the body. If no numbering prefix is found, `prefix` is "" and `body` is the
 * input unchanged.
 *
 * Input is expected to already be {@link normaliseText}d (single-spaced).
 */
export function stripNumberingPrefix(text: string): { prefix: string; body: string } {
  // 1. Compound decimal with optional bracketed sub-parts: 4 / 4.2 / 4.2.1 / 4.2(a) / 4.2(a)(i)
  //    Require at least one '.' OR one bracketed group so a bare "4 word" isn't
  //    swallowed as numbering (a bare leading integer is ambiguous prose).
  const compound = /^(\d+(?:\.\d+)+(?:\([0-9A-Za-z]+\))*|\d+(?:\([0-9A-Za-z]+\))+)(\s+)/.exec(text);
  if (compound) {
    return { prefix: compound[0], body: text.slice(compound[0].length) };
  }

  // 2. Bracketed token: (a) / (iii) / (I) / (12)
  const bracketed = /^(\([0-9A-Za-z]+\)(?:\([0-9A-Za-z]+\))*)(\s+)/.exec(text);
  if (bracketed) {
    return { prefix: bracketed[0], body: text.slice(bracketed[0].length) };
  }

  // 3. Plain number with a trailing dot: 1. / 12.
  const plain = /^(\d+\.)(\s+)/.exec(text);
  if (plain) {
    return { prefix: plain[0], body: text.slice(plain[0].length) };
  }

  return { prefix: '', body: text };
}
