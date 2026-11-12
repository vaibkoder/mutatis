import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding, SuggestedAction } from '@shared/findings/index.js';
import type { AmountPairRecord, TextAnchor } from './types.js';

// ---------------------------------------------------------------------------
// Currency family lookup
// ---------------------------------------------------------------------------

const DOLLAR_FAMILY = new Set(['SGD', 'USD', 'HKD', 'DOLLAR_FAMILY', 'AMBIGUOUS']);
const POUND_FAMILY = new Set(['GBP']);
const EURO_FAMILY = new Set(['EUR']);
const RINGGIT_FAMILY = new Set(['MYR']);

function currencyFamily(code: string | null): string | null {
  if (!code) return null;
  if (DOLLAR_FAMILY.has(code)) return 'dollar';
  if (POUND_FAMILY.has(code)) return 'pound';
  if (EURO_FAMILY.has(code)) return 'euro';
  if (RINGGIT_FAMILY.has(code)) return 'ringgit';
  return code;
}


// ---------------------------------------------------------------------------
// Regex helpers
// ---------------------------------------------------------------------------

// The words group must start with a cardinal number word (never a verb like "Pay").
// Subsequent words may be any letter sequence ("and", "Hundred", etc.).
const NUMBER_WORD_STARTERS =
  'One|Two|Three|Four|Five|Six|Seven|Eight|Nine|Ten' +
  '|Eleven|Twelve|Thirteen|Fourteen|Fifteen|Sixteen|Seventeen|Eighteen|Nineteen' +
  '|Twenty|Thirty|Forty|Fifty|Sixty|Seventy|Eighty|Ninety';

// Anchored number-words fragment: starts with a number word, then allows further words.
const NUM_WORD_FRAG = `(?:${NUMBER_WORD_STARTERS})(?:[\\s\\-]+[A-Za-z]+){0,12}?[\\s\\-]+`;

// Build a pair-extraction regex from the profile's currency table.
// Pattern: <words><currency-word><separator><figure-currency><amount>
function buildPairRegex(profile: JurisdictionProfile): RegExp {
  const table = profile.currency.table;

  // Currency words (sorted longest first to avoid partial matches).
  const currencyWords = Object.keys(table)
    .filter(k => /^[A-Za-z]/.test(k))
    .sort((a, b) => b.length - a.length)
    .map(escapeRe)
    .join('|');

  // Figure-side currency markers: everything that is NOT a plain English currency word.
  // Currency words start with uppercase + lowercase (Dollars, Pounds, Ringgit, etc.).
  // Symbols & ISO codes (S$, SGD, £, $, HK$, RM, GBP, …) do NOT match that pattern.
  const figureSideKeys = Object.keys(table)
    .filter(k => !/^[A-Z][a-z]/.test(k))
    .sort((a, b) => b.length - a.length)
    .map(escapeRe)
    .join('|');

  const figCurPart = `(?:${figureSideKeys})\\s?`;
  const figAmtPart = `\\d{1,3}(?:,\\d{3})*(?:\\.\\d{1,2})?`;

  // Full pair: words + currency-word + optional opener + figure
  const src =
    `(${NUM_WORD_FRAG}(?:${currencyWords}))` +          // group 1: words including currency word
    `\\s*` +
    `(?:[\\(\\-–—,:]\\.?\\s*)?` +                       // optional separator
    `((?:${figCurPart})(${figAmtPart}))` +               // group 2: full figure text; group 3: amount digits
    `\\)?`;

  return new RegExp(src, 'g');
}

// ---------------------------------------------------------------------------
// Words-side parsing
// ---------------------------------------------------------------------------

const SUFFIX_RE = /\band\s+(?:no|zero|\d{1,2})\s*\/\s*100(?:\s+cents?)?\b/gi;
const ONLY_FILLER_RE = /\b(?:only|exactly|nett)\b/gi;

// Sub-unit pattern: "… Dollars and Fifty Cents" — splits on "and <N> Cents/Pence"
const SUB_UNIT_RE = /\band\s+(.+?)\s+(?:Cents?|Pence|Pennies)\b/i;

// ---------------------------------------------------------------------------
// Inline word-to-number parser (replaces the 'words-to-numbers' npm package,
// which crashes at import time in Safari/WebKit).
// Handles sub-thousand amounts: handles up to 999 correctly.
// Scale words (billion/million/thousand) are handled by parseSegmented below.
// ---------------------------------------------------------------------------

const _W2N_ONES = [
  'zero','one','two','three','four','five','six','seven','eight','nine',
  'ten','eleven','twelve','thirteen','fourteen','fifteen','sixteen','seventeen',
  'eighteen','nineteen',
];
const _W2N_TENS = ['','','twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];

function wordsToNumbers(text: string): number | null {
  if (!text || !text.trim()) return null;
  const words = text.toLowerCase().replace(/-/g, ' ').split(/\s+/).filter(
    w => w && w !== 'and',
  );
  let current = 0;
  for (const word of words) {
    const onesIdx = _W2N_ONES.indexOf(word);
    if (onesIdx !== -1) { current += onesIdx; continue; }
    const tensIdx = _W2N_TENS.indexOf(word);
    if (tensIdx !== -1) { current += tensIdx * 10; continue; }
    if (word === 'hundred') { current = (current === 0 ? 1 : current) * 100; continue; }
    return null; // unrecognised word
  }
  return current;
}

// Scale words and their numeric values (descending order so we split on the largest first).
const SCALES: Array<[RegExp, number]> = [
  [/\bbillions?\b/i, 1e9],
  [/\bmillions?\b/i, 1e6],
  [/\bthousands?\b/i, 1e3],
];

/**
 * Segment-based amount parser that correctly handles the British legal form
 * "One Hundred and Fifty Thousand" = 150,000.
 *
 * The standard `words-to-numbers` library misreads this as 50,100 because it
 * processes "hundred" and "fifty thousand" as additive top-level terms.
 * We fix this by splitting on scale words (billion/million/thousand), parsing
 * each segment independently (sub-thousand amounts work correctly in the lib),
 * and multiplying by the scale.
 */
function parseSegmented(text: string): number | null {
  const lower = text.toLowerCase().trim();
  if (!lower) return null;

  // Try each scale word from largest to smallest.
  for (const [scaleRe, scaleVal] of SCALES) {
    const idx = lower.search(scaleRe);
    if (idx === -1) continue;

    const highPart = lower.slice(0, idx).trim();
    const scaleMatch = lower.slice(idx).match(scaleRe);
    const afterScale = lower.slice(idx + (scaleMatch?.[0]?.length ?? 0)).trim();

    const highVal = highPart
      ? (wordsToNumbers(highPart) as number | null | undefined) ?? null
      : null;
    const lowVal = afterScale ? parseSegmented(afterScale) : 0;

    if (typeof highVal !== 'number') return null;
    return highVal * scaleVal + (lowVal ?? 0);
  }

  // No scale word — parse directly (sub-thousand amounts)
  const result = wordsToNumbers(lower) as number | string | null | undefined;
  return typeof result === 'number' ? result : null;
}

function parseWordsAmount(
  rawWords: string,
  profile: JurisdictionProfile,
): { value: number | null; currency: string | null } {
  const table = profile.currency.table;

  let text = rawWords;

  // Step 1: Handle sub-unit "… Dollars and Fifty Cents"
  let centsValue = 0;
  const subUnitMatch = SUB_UNIT_RE.exec(text);
  if (subUnitMatch) {
    const centsWords = subUnitMatch[1] ?? '';
    const parsed = wordsToNumbers(centsWords) as number | null | undefined;
    centsValue = typeof parsed === 'number' ? parsed / 100 : 0;
    text = text.slice(0, subUnitMatch.index ?? text.length).trim();
  }

  // Step 2: Strip "and No/100 Cents" and similar suffixes
  text = text.replace(SUFFIX_RE, '').trim();

  // Step 3: Strip "only", "exactly", "nett" at end
  text = text.replace(ONLY_FILLER_RE, '').trim();

  // Step 4: Identify and strip the currency word
  let currency: string | null = null;
  const currencyWordKeys = Object.keys(table)
    .filter(k => /^[A-Za-z]/.test(k))
    .sort((a, b) => b.length - a.length);

  for (const key of currencyWordKeys) {
    const re = new RegExp(`\\b${escapeRe(key)}\\b`, 'i');
    if (re.test(text)) {
      currency = table[key]?.code ?? null;
      text = text.replace(re, '').trim();
      break;
    }
  }

  // Step 5: Use segment-based parser (handles British "and" in "N hundred and M thousand")
  const numericValue = parseSegmented(text);

  return { value: numericValue !== null ? numericValue + centsValue : null, currency };
}

// ---------------------------------------------------------------------------
// Figure-side parsing
// ---------------------------------------------------------------------------

function parseFigureAmount(
  figureText: string,
  profile: JurisdictionProfile,
): { value: number | null; currency: string | null } {
  const table = profile.currency.table;

  // Try to extract a leading currency symbol/code (longest match first)
  const symbolKeys = Object.keys(table).sort((a, b) => b.length - a.length);
  let currency: string | null = null;
  let digits = figureText.trim();

  for (const key of symbolKeys) {
    if (digits.startsWith(key)) {
      currency = table[key]?.code ?? null;
      digits = digits.slice(key.length).trim();
      break;
    }
  }

  // Remove thousands commas and parse
  const cleaned = digits.replace(/,/g, '');
  const value = cleaned.length > 0 ? parseFloat(cleaned) : null;

  return { value: Number.isFinite(value ?? NaN) ? value : null, currency };
}

// ---------------------------------------------------------------------------
// Main scanner
// ---------------------------------------------------------------------------

export function scanAmounts(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Finding[] {
  return analyseAmounts(paras, profile).findings;
}

/**
 * Same single pass as {@link scanAmounts}, but also returns every
 * {@link AmountPairRecord} (anchor filled, `status` classified) for the
 * AI-export deal-facts serialiser (feature 11). Findings are byte-identical to
 * {@link scanAmounts} — no new scanning.
 */
export function analyseAmounts(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): { findings: Finding[]; records: AmountPairRecord[] } {
  const findings: Finding[] = [];
  const collected: AmountPairRecord[] = [];
  const pairRe = buildPairRegex(profile);

  for (const para of paras) {
    pairRe.lastIndex = 0;
    let m: RegExpExecArray | null;

    while ((m = pairRe.exec(para.text)) !== null) {
      const rawText = m[0];
      const wordsText = m[1] ?? '';
      const figureText = m[2] ?? '';

      // Exclude structural identifiers: figure side must have a currency indicator.
      // The regex already requires one, but guard against empty figureCurrency too.
      const words = parseWordsAmount(wordsText, profile);
      const figure = parseFigureAmount(figureText, profile);

      if (figure.currency === null) continue; // no currency → structural (Clause 3, 30 days, etc.)

      const anchor: TextAnchor = {
        paragraphId: para.id,
        startOffset: m.index,
        endOffset: m.index + rawText.length,
      };

      const record = classify(rawText, wordsText, figureText, words, figure, profile);
      record.anchor = anchor; // classify() leaves a placeholder anchor for the caller to fill
      collected.push(record);
      const finding = recordToFinding(record, anchor, para);
      if (finding) findings.push(finding);
    }
  }

  return { findings, records: collected };
}

// ---------------------------------------------------------------------------
// Classification
// ---------------------------------------------------------------------------

function classify(
  rawText: string,
  wordsText: string,
  figureText: string,
  words: { value: number | null; currency: string | null },
  figure: { value: number | null; currency: string | null },
  profile: JurisdictionProfile,
): AmountPairRecord {
  const { centTolerance } = profile.currency;

  // Determine status
  const wc = words.currency;
  const fc = figure.currency;

  // Bare $ on the figure side → soft Informational only (never Critical from bare $)
  const figureIsAmbiguous = fc === 'AMBIGUOUS';

  // Value comparison
  let valueMismatch = false;
  if (words.value !== null && figure.value !== null) {
    valueMismatch = Math.abs(words.value - figure.value) > centTolerance;
  }

  // Currency family comparison.
  // DOLLAR_FAMILY (word "Dollars") is consistent with any dollar-family figure (SGD/USD/HKD).
  // Assert mismatch only across incompatible families; never assert mismatch when figure is AMBIGUOUS.
  const wcFamily = wc === 'DOLLAR_FAMILY' ? 'dollar' : currencyFamily(wc);
  const fcFamily = figureIsAmbiguous ? null : currencyFamily(fc);
  const currencyMismatch =
    wcFamily !== null && fcFamily !== null &&
    wcFamily !== fcFamily;

  let status: AmountPairRecord['status'];
  if (words.value === null && figure.value === null) {
    status = 'unparseable';
  } else if (figureIsAmbiguous) {
    // Bare $ → soft regardless of value agreement
    status = valueMismatch ? 'value-mismatch' : 'ambiguous-currency';
  } else if (valueMismatch && currencyMismatch) {
    status = 'both-mismatch';
  } else if (valueMismatch) {
    status = 'value-mismatch';
  } else if (currencyMismatch) {
    status = 'currency-mismatch';
  } else {
    status = 'match';
  }

  return {
    anchor: { paragraphId: '', startOffset: 0, endOffset: 0 }, // filled by caller
    rawText,
    wordsText,
    wordsCurrency: wc,
    wordsValue: words.value,
    figureText,
    figureCurrency: fc,
    figureValue: figure.value,
    isRange: false,
    qualifier: null,
    status,
  };
}

// ---------------------------------------------------------------------------
// Finding construction
// ---------------------------------------------------------------------------

function recordToFinding(
  record: AmountPairRecord,
  anchor: TextAnchor,
  para: ParaNode,
): Finding | null {
  const { status, wordsText, wordsValue, figureText, figureValue,
          wordsCurrency, figureCurrency } = record;

  if (status === 'match' || status === 'unparseable') return null;

  const idBase = `factual-consistency/amount/${anchor.paragraphId}-${anchor.startOffset}`;

  switch (status) {
    case 'value-mismatch':
    case 'both-mismatch': {
      const suggestion = buildAmountSuggestion(wordsText, figureText, wordsValue, figureValue);
      const finding: Finding = {
        id: idBase,
        category: 'factual-consistency/amount-mismatch',
        featureId: 'factual-consistency',
        severity: 'error',
        title: `Amount in words (${wordsValue ?? '?'}) does not match figure ${figureText} (${figureValue ?? '?'})`,
        detail: `Words: "${wordsText}" → ${wordsValue ?? 'unparseable'}; Figure: "${figureText}" → ${figureValue ?? 'unparseable'}`,
        location: { paraId: para.id, charStart: anchor.startOffset, charEnd: anchor.endOffset, label: para.id },
      };
      if (suggestion) finding.suggestedAction = suggestion;
      return finding;
    }
    case 'currency-mismatch':
      return {
        id: idBase,
        category: 'factual-consistency/currency-mismatch',
        featureId: 'factual-consistency',
        severity: 'error',
        title: `Currency words (${wordsCurrency ?? '?'}) disagree with figure currency (${figureCurrency ?? '?'})`,
        detail: `"${wordsText}" vs "${figureText}"`,
        location: { paraId: para.id, charStart: anchor.startOffset, charEnd: anchor.endOffset, label: para.id },
      };
    case 'ambiguous-currency':
      return {
        id: idBase,
        category: 'factual-consistency/ambiguous-currency',
        featureId: 'factual-consistency',
        severity: 'info',
        title: `Bare '$' — currency not specified (S$/US$/HK$?). Consider qualifying.`,
        detail: `"${figureText}"`,
        location: { paraId: para.id, charStart: anchor.startOffset, charEnd: anchor.endOffset, label: para.id },
      };
  }
}

function buildAmountSuggestion(
  wordsText: string,
  figureText: string,
  wordsValue: number | null,
  figureValue: number | null,
): SuggestedAction | null {
  if (wordsValue === null || figureValue === null) return null;

  return {
    label: 'View suggested corrections',
    apply: async (_ctx) => {
      // Read-only — never auto-edits. User sees both options in the UI.
    },
  };
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
