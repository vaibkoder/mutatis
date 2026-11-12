import type { TextAnchor } from '@shared/doc-model/anchor.js';

export type { TextAnchor };

/** A date found in a paragraph, with parsed details and a role hint. */
export interface DateRecord {
  anchor: TextAnchor;
  rawText: string;
  parsed: Date | null;
  format: 'dd-month-yyyy' | 'numeric-slash' | 'numeric-dash' | 'ordinal-long' | 'template' | 'unknown';
  centuryAmbiguous: boolean;
  isTemplate: boolean;
  roleHint: 'execution' | 'completion' | 'effective' | 'asof' | 'other';
  dayOfWeek: 0 | 1 | 2 | 3 | 4 | 5 | 6 | null;
  /**
   * True when this date was flagged as a weekend execution/completion date
   * (the deal-facts AI-export reads it directly — F4). Set during scanning;
   * absent on records the scanner never weekend-tested. Optional so existing
   * record literals remain valid.
   */
  flaggedWeekend?: boolean;
}

/** A currency-annotated amount pair (words + parenthetical figure). */
export interface AmountPairRecord {
  anchor: TextAnchor;
  rawText: string;
  wordsText: string;
  wordsCurrency: string | null;
  wordsValue: number | null;
  figureText: string;
  figureCurrency: string | null;
  figureValue: number | null;
  isRange: boolean;
  qualifier: 'up-to' | 'not-less-than' | 'at-least' | 'approx' | null;
  status: 'match' | 'value-mismatch' | 'currency-mismatch' | 'both-mismatch' | 'ambiguous-currency' | 'unparseable';
}

/** A placeholder hit with suppression metadata. */
export interface PlaceholderHit {
  anchor: TextAnchor;
  rawText: string;
  category: 'filled-bullet' | 'keyword' | 'note-to-draft' | 'insert-directive' | 'bare-brackets' | 'percent-bullet';
  suppressed: boolean;
  suppressReason?: 'whitelist' | 'case-citation' | 'signature-line';
  confidence: 'high' | 'medium';
}
