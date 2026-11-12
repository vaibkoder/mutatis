import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import { addFindings, removeByPrefix } from '@shared/findings/index.js';
import type { DateRecord, AmountPairRecord, PlaceholderHit } from './types.js';
import { analysePlaceholders } from './placeholders.js';
import { analyseAmounts } from './amounts.js';
import { analyseDates } from './dates.js';

/** Findings + the raw records that produced them (retained for AI-export). */
export interface FactualConsistencyAnalysis {
  findings: Finding[];
  /** Non-template dates with role/day/weekend-flag resolved. */
  dates: DateRecord[];
  /** Every currency word↔figure pair, classified. */
  amounts: AmountPairRecord[];
  /** De-duplicated non-suppressed placeholder hits. */
  placeholders: PlaceholderHit[];
}

/**
 * Run all three Factual Consistency sub-checks over the Document-Model
 * paragraphs in a single pass and return findings + the raw records. Pure
 * (no findings-store mutation). All processing is in-memory over
 * ParaNode.text (AC-6 — no re-walk).
 */
export function analyseFactualConsistency(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): FactualConsistencyAnalysis {
  const placeholders = analysePlaceholders(paras, profile);
  const amounts = analyseAmounts(paras, profile);
  const dates = analyseDates(paras, profile);
  return {
    findings: [...placeholders.findings, ...amounts.findings, ...dates.findings],
    dates: dates.records,
    amounts: amounts.records,
    placeholders: placeholders.records,
  };
}

/**
 * Findings-only convenience wrapper (parity with the other checkers).
 *
 * Called by the "Run Proof checks" action on the Proof tab.
 */
export function runFactualConsistency(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Finding[] {
  return analyseFactualConsistency(paras, profile).findings;
}

/**
 * Clear previous factual-consistency findings and emit fresh results.
 * Returns the analysis so the Proof card can retain the raw records for the
 * AI-export deal-facts serialiser (feature 11).
 */
export function clearThenEmit(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): FactualConsistencyAnalysis {
  removeByPrefix('factual-consistency');
  const analysis = analyseFactualConsistency(paras, profile);
  addFindings(analysis.findings);
  return analysis;
}
