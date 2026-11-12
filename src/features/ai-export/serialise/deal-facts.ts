/**
 * Deal-facts serialiser (feature 11 — spec §4.5, rule F4: normalise + preserve raw).
 *
 * Projects the factual-consistency records into a clean deal sheet: dates with
 * both raw and ISO forms, amounts with both word and figure values plus any
 * mismatch flag, and unfilled placeholders. The AI gets machine values without
 * losing how the text appears in the document.
 */
import type { DateRecord, AmountPairRecord, PlaceholderHit } from '@features/factual-consistency/index.js';
import type { DealFactsExport, DateEntry, AmountEntry, PlaceholderEntry } from '../types.js';
import type { Citer } from '../cite.js';

export interface DealFactsInput {
  dates: readonly DateRecord[];
  amounts: readonly AmountPairRecord[];
  placeholders: readonly PlaceholderHit[];
}

export interface DealFactsDeps {
  cite: Citer;
}

const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const MISMATCH: Partial<Record<AmountPairRecord['status'], AmountEntry['mismatch']>> = {
  'value-mismatch': 'value',
  'currency-mismatch': 'currency',
  'both-mismatch': 'both',
};

export function serialiseDealFacts(input: DealFactsInput, deps: DealFactsDeps): DealFactsExport {
  return {
    dates: input.dates.map(d => toDateEntry(d, deps)),
    amounts: input.amounts.map(a => toAmountEntry(a, deps)),
    placeholders: input.placeholders.map(p => toPlaceholderEntry(p, deps)),
  };
}

function toDateEntry(rec: DateRecord, deps: DealFactsDeps): DateEntry {
  return {
    raw: rec.rawText,
    iso: rec.parsed ? isoDate(rec.parsed) : '',
    day_of_week: rec.dayOfWeek !== null ? (DAY_NAMES[rec.dayOfWeek] ?? '') : '',
    role: rec.roleHint,
    flagged_weekend: rec.flaggedWeekend === true,
    clause: deps.cite(rec.anchor.paragraphId),
  };
}

function toAmountEntry(rec: AmountPairRecord, deps: DealFactsDeps): AmountEntry {
  const entry: AmountEntry = {
    raw: rec.rawText,
    value: rec.figureValue,
    currency: rec.figureCurrency && rec.figureCurrency !== 'AMBIGUOUS' ? rec.figureCurrency : '',
    clause: deps.cite(rec.anchor.paragraphId),
  };
  if (rec.wordsText) entry.words_raw = rec.wordsText;
  if (rec.wordsValue !== null) entry.words_value = rec.wordsValue;
  const mismatch = MISMATCH[rec.status];
  if (mismatch) entry.mismatch = mismatch;
  return entry;
}

function toPlaceholderEntry(hit: PlaceholderHit, deps: DealFactsDeps): PlaceholderEntry {
  return {
    raw: hit.rawText,
    category: hit.category,
    clause: deps.cite(hit.anchor.paragraphId),
  };
}

/** Local-date yyyy-mm-dd (avoids the UTC shift `toISOString()` would introduce). */
function isoDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
