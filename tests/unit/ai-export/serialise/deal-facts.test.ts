import { describe, it, expect } from 'vitest';
import { serialiseDealFacts } from '@features/ai-export/serialise/deal-facts.js';
import type { DateRecord, AmountPairRecord, PlaceholderHit } from '@features/factual-consistency/index.js';

const cite = (id: string) => (id === 'pd' ? 'Clause 2' : id === 'pa' ? 'Clause 3' : 'Clause 4');

const anchor = (paragraphId: string) => ({ paragraphId, startOffset: 0, endOffset: 1 });

describe('serialiseDealFacts', () => {
  const dates: DateRecord[] = [{
    anchor: anchor('pd'),
    rawText: '6 June 2026',
    parsed: new Date(2026, 5, 6), // Saturday
    format: 'dd-month-yyyy',
    centuryAmbiguous: false,
    isTemplate: false,
    roleHint: 'execution',
    dayOfWeek: 6,
    flaggedWeekend: true,
  }];
  const amounts: AmountPairRecord[] = [{
    anchor: anchor('pa'),
    rawText: 'One Hundred Pounds (£200)',
    wordsText: 'One Hundred Pounds',
    wordsCurrency: 'GBP', wordsValue: 100,
    figureText: '£200', figureCurrency: 'GBP', figureValue: 200,
    isRange: false, qualifier: null, status: 'value-mismatch',
  }];
  const placeholders: PlaceholderHit[] = [{
    anchor: anchor('pp'), rawText: '[TBC]', category: 'keyword', suppressed: false, confidence: 'high',
  }];

  it('flags the weekend date and normalises iso/day (F4)', () => {
    const out = serialiseDealFacts({ dates, amounts: [], placeholders: [] }, { cite });
    const d = out.dates[0]!;
    expect(d.flagged_weekend).toBe(true);
    expect(d.iso).toBe('2026-06-06');
    expect(d.day_of_week).toBe('Saturday');
    expect(d.role).toBe('execution');
    expect(d.clause).toBe('Clause 2');
  });

  it('carries the amount mismatch and both raw + value forms', () => {
    const out = serialiseDealFacts({ dates: [], amounts, placeholders: [] }, { cite });
    const a = out.amounts[0]!;
    expect(a.mismatch).toBe('value');
    expect(a.value).toBe(200);
    expect(a.currency).toBe('GBP');
    expect(a.words_value).toBe(100);
    expect(a.words_raw).toBe('One Hundred Pounds');
  });

  it('omits mismatch when the amount matches', () => {
    const matched: AmountPairRecord[] = [{ ...amounts[0]!, status: 'match', wordsValue: 200 }];
    const out = serialiseDealFacts({ dates: [], amounts: matched, placeholders: [] }, { cite });
    expect(out.amounts[0]!.mismatch).toBeUndefined();
  });

  it('forwards placeholder category and clause', () => {
    const out = serialiseDealFacts({ dates: [], amounts: [], placeholders }, { cite });
    expect(out.placeholders[0]!.category).toBe('keyword');
    expect(out.placeholders[0]!.raw).toBe('[TBC]');
    expect(out.placeholders[0]!.clause).toBe('Clause 4');
  });

  it('empties iso/day for an unparsed date', () => {
    const unparsed: DateRecord[] = [{ ...dates[0]!, parsed: null, dayOfWeek: null, flaggedWeekend: false }];
    const out = serialiseDealFacts({ dates: unparsed, amounts: [], placeholders: [] }, { cite });
    expect(out.dates[0]!.iso).toBe('');
    expect(out.dates[0]!.day_of_week).toBe('');
    expect(out.dates[0]!.flagged_weekend).toBe(false);
  });
});
