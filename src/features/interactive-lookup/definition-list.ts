/**
 * Accumulating definition-list state for the Interactive Lookup panel (feature 04b).
 * Pure — no Office JS, no signals. The driver holds the list in a signal and
 * uses `promoteDefinition` to update it on each defined-term lookup.
 */

import type { LookupDefinition, LookupDefByXref } from './types.js';

/** One entry in the side-panel definition list. */
export interface DefinitionListEntry {
  /** Normalised lookup key — the dedupe identity. */
  termKey: string;
  /** Surface form, e.g. "Completion Date". */
  display: string;
  /** The resolved definition result (single or definition-by-cross-ref). */
  result: LookupDefinition | LookupDefByXref;
}

/** Cap on how many definitions are retained in the list (bounded memory). */
export const MAX_DEFINITION_LIST = 25;

/**
 * Insert `entry` at the front of the list, de-duplicating by `termKey`
 * (re-looking-up a term promotes its existing slot rather than duplicating it,
 * and refreshes its stored result). The list is bounded to `MAX_DEFINITION_LIST`.
 *
 * Returns a new array; never mutates the input.
 */
export function promoteDefinition(
  list: ReadonlyArray<DefinitionListEntry>,
  entry: DefinitionListEntry,
): DefinitionListEntry[] {
  const withoutDup = list.filter(e => e.termKey !== entry.termKey);
  return [entry, ...withoutDup].slice(0, MAX_DEFINITION_LIST);
}
