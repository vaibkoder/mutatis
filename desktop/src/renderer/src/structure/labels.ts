/**
 * Clause-number arithmetic: tidying a painted or typed number, telling what kind
 * of number it is, and composing a citation from a chain of them.
 *
 * Citation rules (the ones a lawyer writes): "1." under nothing is "1"; "(a)"
 * under "4.2" is "4.2(a)"; a dotted number such as "14.10" is anchored and starts
 * the citation afresh; a bare token under a citation is bracketed ("6(1)", never
 * "61"); a paragraph of a schedule is "Schedule 2 para 3(a)".
 */

/** "2.8." -> "2.8"; "a)" -> "(a)"; "1)" -> "(1)"; "(a)" unchanged. */
export function tidyLabel(label: string | null | undefined): string {
  if (!label) return '';
  const t = label.trim();
  if (/^[A-Za-z0-9]{1,6}\)$/.test(t)) return `(${t.slice(0, -1)})`;
  return t.replace(/[.:]+$/, '').trim();
}

const ANCHORED = /^(?:[A-Za-z]+\s+)?(?=[A-Za-z0-9.]*\d)[A-Za-z0-9]+(?:\.[A-Za-z0-9]+)+$/;
const ANCHORED_NUM_LETTER = /^[0-9]+[A-Za-z]$/;
const ANCHORED_WORD_INT = /^(?:Section|Clause|Article|Rule|Regulation)\s+[0-9IVXivx]+$/;

export function isAnchored(tidy: string): boolean {
  return ANCHORED.test(tidy) || ANCHORED_WORD_INT.test(tidy) || ANCHORED_NUM_LETTER.test(tidy);
}

/** ['1.', '1.1', '(a)', '(i)'] -> '1.1(a)(i)'. */
export function compose(chain: Iterable<string | null | undefined>): string | null {
  let out = '';
  for (const raw of chain) {
    let t = tidyLabel(raw);
    if (!t) continue;
    if (!out || isAnchored(t)) {
      out = t;
    } else {
      if (/^[A-Za-z0-9]{1,6}$/.test(t)) t = `(${t})`;
      out += t;
    }
  }
  return out || null;
}

/** '(A)' -> 'A'; '14.' -> '14'. */
export function stripLabelPunct(label: string): string {
  return label.trim().replace(/^[([]+|[)\]]+$/g, '').replace(/[.:)]+$/, '').trim();
}

export type NumberClass =
  | 'decimal'
  | 'dotted'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'other';

export interface NumberShape {
  readonly cls: NumberClass;
  /** The token stripped of brackets and dots ("b", "iv", "4.2"). */
  readonly core: string;
  /** Punctuation shape: "(#)", "#.", "#)", "#". Siblings share it. */
  readonly shape: string;
  /** Components of a dotted number. */
  readonly parts: readonly string[];
  /** Word's numFmt settled letter-or-numeral, so "(i)" after "(h)" is not guessed. */
  readonly certain?: boolean;
}

const ROMAN = /^(?=[ivxlcdm]+$)m{0,3}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/;

export function romanValue(s: string): number {
  const map: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  let total = 0;
  const lower = s.toLowerCase();
  for (let i = 0; i < lower.length; i += 1) {
    const v = map[lower[i]!] ?? 0;
    const next = map[lower[i + 1] ?? ''] ?? 0;
    total += v < next ? -v : v;
  }
  return total;
}

export function isRoman(s: string): boolean {
  return s !== '' && ROMAN.test(s.toLowerCase());
}

/** Classify a number from its text, using Word's numFmt when it is known. */
export function numberShape(label: string, numFmt: string | null = null): NumberShape {
  const t = label.trim();
  const shape = t.replace(/[0-9]+|[A-Za-z]+/g, '#').replace(/#(?:\.#)+/g, '#.#');
  const core = t.replace(/^[([]+/, '').replace(/[)\].:]+$/, '');
  const parts = core.split('.');
  if (/^\d+(?:\.\d+)+$/.test(core) || /^\d+(?:\.[0-9A-Za-z]+)+$/.test(core)) {
    return { cls: 'dotted', core, shape, parts };
  }
  if (/^\d+[A-Za-z]?$/.test(core)) return { cls: 'decimal', core, shape, parts };
  const fmt = numFmt ?? '';
  if (/^[a-z]+$/.test(core)) {
    if (fmt === 'lowerRoman') return { cls: 'lowerRoman', core, shape, parts, certain: true };
    if (fmt === 'lowerLetter') return { cls: 'lowerLetter', core, shape, parts, certain: true };
    if (core.length > 1 && isRoman(core)) return { cls: 'lowerRoman', core, shape, parts };
    return { cls: 'lowerLetter', core, shape, parts };
  }
  if (/^[A-Z]+$/.test(core)) {
    if (fmt === 'upperRoman') return { cls: 'upperRoman', core, shape, parts, certain: true };
    if (fmt === 'upperLetter') return { cls: 'upperLetter', core, shape, parts, certain: true };
    if (core.length > 1 && isRoman(core)) return { cls: 'upperRoman', core, shape, parts };
    return { cls: 'upperLetter', core, shape, parts };
  }
  return { cls: 'other', core, shape, parts };
}

/** The letter after `s` ("a" -> "b", "z" -> "aa"), for telling "(i)" the letter from "(i)" the numeral. */
export function nextLetter(s: string): string {
  if (/^([a-z])\1*$/i.test(s) && s.length > 1) return s[0]!.repeat(s.length);
  const c = s.charCodeAt(s.length - 1);
  return s.slice(0, -1) + String.fromCharCode(c + 1);
}

/**
 * A number typed at the start of the paragraph text ("4.2\tThe Buyer", "(a) the",
 * "3. General"). Deliberately strict: a bare integer must be followed by a dot,
 * bracket or tab, so "5 days after" and "1976 Act" are prose.
 */
export function typedNumber(text: string): { token: string; rest: string } | null {
  const m =
    /^\s*(\d{1,3}(?:\.\d{1,3})+\.?|\d{1,3}\.[A-Z]{1,2}(?:\.\d{1,3})*\.?|\d{1,3}\.\d{1,3}[A-Za-z]|\d{1,3}[A-Z]?\.|\(\d{1,3}[A-Za-z]?\)|\d{1,3}\)|\([a-zA-Z]{1,5}\)|[a-zA-Z]\)|[ivxIVX]{1,5}\.|[A-Z]\.)(?=[\t\u00a0 ]|$)[\t\u00a0 ]*/.exec(
      text,
    );
  if (m) return { token: m[1]!, rest: text.slice(m[0].length) };
  // an optional clause in square brackets: "[10.15\tExport Control ...]"
  const bracketed = /^\s*\[(\d{1,3}(?:\.\d{1,3})+\.?|\d{1,3}\.)[\t\u00a0 ]+/.exec(text);
  if (bracketed) return { token: bracketed[1]!, rest: text.slice(bracketed[0].length) };
  const tab = /^\s*(\d{1,3}[A-Z]?)\t+/.exec(text);
  if (tab) return { token: tab[1]!, rest: text.slice(tab[0].length) };
  return null;
}
