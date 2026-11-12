/**
 * Atomic-token tokenizer for the Level-2 word diff.
 * Spec: spec/features/08-compare.md §5.3.
 *
 * Tokenizes text into an array of tokens. Certain spans are ATOMIC — they must
 * never be split mid-token because splitting them destroys legal meaning
 * (`£40,000` → `£50,000` must read as one amount change, not `4`→`5`):
 *
 *  - Quoted defined terms: "Permitted Encumbrance" / 'Permitted Encumbrance'.
 *  - Clause/Schedule refs (config.referenceTokens): `Clause 4.2(a)`,
 *    `Schedule 3`, `paragraph 2 of Schedule 4`.
 *  - Amounts: `S$1,250,000`, `£40,000`, `US$2.5m`, `$1,000`.
 *  - Dates: `1 January 2026`, `1 Jan 2026`, `31 December 2025`.
 *  - Percentages: `7.5%`, `100%`.
 *  - Durations: `30 days`, `3 months`, `three (3) Business Days`, `5 Business Days`.
 *
 * A cross-reference is matched as a whole span but emitted as SEVERAL tokens —
 * the reference noun and the `of` connective as ordinary words, each reference
 * VALUE as the atomic one. P5 makes the number indivisible; it does not make the
 * noun in front of it borrowable, and marking both sides of `Clause 9.13 →
 * Clause 9.14` borrows a word the reader then re-reads to no purpose (P1).
 *
 * Non-atomic spans are split into individual word / whitespace / punctuation
 * tokens. Pure.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { AtomicKind } from '../model.js';

const MONTHS =
  '(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)';

const NUMBER_WORDS =
  '(?:zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand|million|billion)';

const DURATION_UNIT =
  '(?:days?|weeks?|months?|years?|Business\\s+Days?|business\\s+days?|hours?|minutes?)';

/**
 * Does this text carry anything a reader could see — a letter or a digit?
 *
 * The one definition of "readable content" in the feature. `refine.ts` needs the
 * negation to spot a noise-only token; `relocate.ts` needs it to refuse a
 * punctuation-only span; the readability fixture needs it because a struck span
 * with no readable content scores the MAXIMUM on nothing (a measurement artefact).
 * Three copies of the same regex was the kind of duplication the reuse-before-build rule forbids.
 */
export function hasReadableContent(text: string): boolean {
  return /[\p{L}\p{N}]/u.test(text);
}

/** One token of the word diff, carrying the atomic kind it was matched as. */
export interface Token {
  text: string;
  atomic?: AtomicKind;
}

interface Matcher {
  kind: AtomicKind;
  re: RegExp;
  /** Turn a match into tokens. Default: the whole span as one atomic token. */
  emit?: (m: RegExpExecArray) => Token[];
}

/** Build the ordered list of atomic matchers. Order matters (longest/most-specific first). */
function buildAtomicMatchers(config: CompareConfig): Matcher[] {
  const tokens = config.referenceTokens.map(escapeRegExp).join('|');
  // A reference value such as 4.2(a)(i), 3, A.1 — digits, letters, dots, brackets.
  const refValue = '[\\dA-Za-z]+(?:\\.[\\dA-Za-z]+)*(?:\\([0-9A-Za-z]+\\))*';

  return [
    // Quoted defined terms (straight or curly quotes).
    //
    // The SINGLE-quote alternatives are guarded on both sides, because in
    // English legal drafting an apostrophe is far more often a possessive than
    // a quotation mark, and `'[^']+'` unguarded spans everything between any
    // two possessives — `Vendor's … Purchaser's` becomes one very long
    // ATOMIC span, P5 then forbids splitting it, and changing one word inside
    // strikes and re-inserts the lot. §5.2 folds `’ → '` before this runs, so
    // EVERY Word document hits it (ADR-0029, correction of 2026-09-11).
    //
    //   (?<![\p{L}\p{N}])  an opening quote never follows a word character —
    //                      rules out `Vendor'`s and the plural `Vendors'`
    //   (?![\p{L}\p{N}])   a closing quote is never followed by one — rules out
    //                      `'the Vendor'`s, where the "closing" quote is really
    //                      the possessive in `Vendor's`
    //
    // The curly pair carries the same guards: `’` is the typographic
    // possessive, so the defect is identical on any text §5.2 has not folded.
    {
      kind: 'definedTerm',
      re: new RegExp(
        '"[^"]+"|“[^”]+”' +
          '|(?<![\\p{L}\\p{N}])\'[^\']+\'(?![\\p{L}\\p{N}])' +
          '|(?<![\\p{L}\\p{N}])‘[^’]+’(?![\\p{L}\\p{N}])',
        'yu',
      ),
    },

    // Clause/Schedule references, optionally "<tok> N of <tok> M". Captured in
    // parts so the noun and the connective stay ordinary words (module doc).
    {
      kind: 'clauseRef',
      re: new RegExp(
        `(${tokens})(\\s+)(${refValue})(?:(\\s+)(of)(\\s+)(${tokens})(\\s+)(${refValue}))?`,
        'yi',
      ),
      emit: emitClauseRef,
    },

    // Amounts: optional currency prefix (S$, US$, A$, £, €, $) + grouped digits + optional .dd + optional m/bn/k suffix.
    {
      kind: 'amount',
      re: /(?:[A-Z]{0,2}\$|£|€)\s?\d{1,3}(?:,\d{3})*(?:\.\d+)?(?:\s?(?:m|bn|k|million|billion))?/y,
    },

    // Dates: "1 January 2026", "1 Jan 2026", "31 December 2025".
    { kind: 'date', re: new RegExp(`\\d{1,2}\\s+${MONTHS}\\.?\\s+\\d{4}`, 'y') },

    // Percentages: 7.5% / 100%.
    { kind: 'percent', re: /\d+(?:\.\d+)?\s?%/y },

    // Durations: "three (3) Business Days", "30 days", "3 months", "5 Business Days".
    {
      kind: 'duration',
      re: new RegExp(
        `(?:${NUMBER_WORDS}\\s+\\(\\d+\\)\\s+|\\d+\\s+)${DURATION_UNIT}`,
        'y',
      ),
    },
  ];
}

/**
 * Emit a matched cross-reference as noun / connective / VALUE tokens. Only the
 * values carry `clauseRef`, so a renumbering marks the number and leaves the
 * reference noun — identical on both sides — outside the mark.
 */
function emitClauseRef(m: RegExpExecArray): Token[] {
  const out: Token[] = [{ text: m[1]! }, { text: m[2]! }, { text: m[3]!, atomic: 'clauseRef' }];
  if (m[4] !== undefined) {
    out.push(
      { text: m[4]! },
      { text: m[5]! },
      { text: m[6]! },
      { text: m[7]! },
      { text: m[8]! },
      { text: m[9]!, atomic: 'clauseRef' },
    );
  }
  return out;
}

/**
 * Detect the {@link AtomicKind} of a complete token, or undefined if it is an
 * ordinary word/punctuation token. The token is matched in full (anchored).
 *
 * NB a bare reference VALUE (`9.13`) is not recognisable out of context — the
 * value pattern alone matches almost any word — so a caller needing the kind of
 * a tokenizer-produced token reads {@link Token.atomic} from
 * {@link tokenizeTagged} rather than re-deriving it here.
 */
export function atomicKindOf(token: string, config: CompareConfig): AtomicKind | undefined {
  const matchers = buildAtomicMatchers(config);
  for (const { kind, re } of matchers) {
    re.lastIndex = 0;
    const m = re.exec(token);
    if (m && m.index === 0 && m[0].length === token.length) {
      return kind;
    }
  }
  return undefined;
}

/**
 * Tokenize `text`, tagging each token with the atomic kind it was matched as.
 * Atomic spans (see module doc) are emitted whole; a cross-reference is emitted
 * as noun / connective / value tokens with only the values tagged. Whitespace
 * tokens are preserved so the diff and any re-join is faithful.
 */
export function tokenizeTagged(text: string, config: CompareConfig): Token[] {
  const matchers = buildAtomicMatchers(config);
  const out: Token[] = [];
  let i = 0;
  const n = text.length;

  while (i < n) {
    // Try each atomic matcher anchored at the current position (sticky regexes).
    let matched = false;
    for (const { kind, re, emit } of matchers) {
      re.lastIndex = i;
      const m = re.exec(text);
      if (m && m.index === i && m[0].length > 0) {
        if (emit) out.push(...emit(m));
        else out.push({ text: m[0], atomic: kind });
        i += m[0].length;
        matched = true;
        break;
      }
    }
    if (matched) continue;

    // Whitespace run.
    if (/\s/.test(text.charAt(i))) {
      let j = i + 1;
      while (j < n && /\s/.test(text.charAt(j))) j++;
      out.push({ text: text.slice(i, j) });
      i = j;
      continue;
    }

    // Word run (letters/digits/'&/-/. kept together for ordinary words).
    if (/[A-Za-z0-9]/.test(text.charAt(i))) {
      let j = i + 1;
      while (j < n && /[A-Za-z0-9'&.\-]/.test(text.charAt(j))) {
        // Stop a word before a position where an atomic token would start, so we
        // don't fuse e.g. "pay" into "pay£40,000" (rare, but be safe). Only ever
        // break where a reader would see a word boundary: the reference-token
        // match is case-insensitive, so without this "subparagraph 3" splits
        // into "sub" + "paragraph 3" and "Subclause 4.2" loses its "Sub".
        if (/[A-Za-z0-9]/.test(text.charAt(j - 1)) && /[A-Za-z0-9]/.test(text.charAt(j))) {
          j++;
          continue;
        }
        if (startsAtomic(text, j, matchers)) break;
        j++;
      }
      out.push({ text: text.slice(i, j) });
      i = j;
      continue;
    }

    // Single punctuation / other character.
    out.push({ text: text.charAt(i) });
    i++;
  }

  return out;
}

/** Tokenize `text` into token texts. See {@link tokenizeTagged} for the kinds. */
export function tokenize(text: string, config: CompareConfig): string[] {
  return tokenizeTagged(text, config).map((t) => t.text);
}

/**
 * True if an atomic matcher matches `text` starting exactly at `at`.
 *
 * Takes the WHOLE text and an offset rather than a slice, because the
 * defined-term matcher's opening guard is a lookbehind: against a slice it
 * would see nothing before position 0 and pass vacuously, so a word run would
 * break at a possessive apostrophe here and then fail to match in the main
 * loop, splitting `Vendors'` into two tokens. Sticky matching at an offset is
 * otherwise identical for every pattern in the set.
 */
function startsAtomic(text: string, at: number, matchers: readonly Matcher[]): boolean {
  for (const { re } of matchers) {
    re.lastIndex = at;
    const m = re.exec(text);
    if (m && m.index === at && m[0].length > 0) return true;
  }
  return false;
}

/** Escape a string for literal use inside a RegExp. */
function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
