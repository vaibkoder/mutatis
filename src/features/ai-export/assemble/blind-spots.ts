/**
 * Blind-spot disclosure assembly (feature 11 — spec §4.1, rule F6, ADR-0008).
 *
 * Every export states what it did NOT cover, so an LLM consumer cannot assert
 * "no issues" over a region Mutatis never analysed. The array is the machine
 * analogue of the human "what is NOT compared" banner.
 */
import type { CheckId } from '../types.js';

/** A checker that did not run → the disclosure for the data it would have produced. */
const ABSENT_CHECK_DISCLOSURE: Record<CheckId, string> = {
  'defined-terms':
    'Defined-terms checker was not run; glossary and usage findings absent.',
  'cross-reference':
    'Cross-reference / section-index checker was not run; clause-tree absent.',
  'factual-consistency':
    'Factual-consistency checker was not run; deal dates, amounts, and placeholders absent.',
  compare: 'Compare was not run; change set absent.',
  'party-role': 'Party-role checker was not run; cast list absent.',
};

/** Always appended regardless of what ran (spec §4.1). */
const ALWAYS_APPEND: readonly string[] = [
  'Headers, footers, and text boxes are not scanned.',
  'Image-only content (scanned pages, embedded pictures) is not text-compared.',
  'Cross-document and external references are detected but not resolved (Office JS cannot read other files).',
];

const ALL_CHECKS: readonly CheckId[] = [
  'defined-terms',
  'cross-reference',
  'factual-consistency',
  'compare',
  'party-role',
];

/**
 * Build the `blind_spots` array: one entry per checker NOT in `checksRun`,
 * followed by the three fixed always-append entries.
 */
export function buildBlindSpots(checksRun: readonly CheckId[]): string[] {
  const ran = new Set(checksRun);
  const out: string[] = [];
  for (const check of ALL_CHECKS) {
    if (!ran.has(check)) out.push(ABSENT_CHECK_DISCLOSURE[check]);
  }
  out.push(...ALWAYS_APPEND);
  return out;
}
