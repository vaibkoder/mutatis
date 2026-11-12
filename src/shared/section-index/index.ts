export type {
  SectionIndex,
  SectionNode,
  SectionKey,
  SectionKind,
  SectionIndexEntry,
} from './types.js';
export { buildSectionIndex } from './builder.js';
export {
  normaliseNumber,
  normaliseSubToken,
  stripParens,
  isAmbiguousRomanLetter,
  disambiguateRoman,
} from './normalise.js';

import type { SectionIndex, SectionKey, SectionNode } from './types.js';

/** Resolve a key against the index. [] = not found, [x] = unique, [x,y…] = ambiguous. */
export function resolveKey(index: SectionIndex, key: SectionKey): SectionNode[] {
  return index.resolve(key);
}

/** Classification of a cross-reference resolution. */
export type RefClassification = 'ok' | 'not-found' | 'ambiguous';

export function classifyKey(index: SectionIndex, key: SectionKey): RefClassification {
  const entries = resolveKey(index, key);
  if (entries.length === 0) return 'not-found';
  if (entries.length > 1) return 'ambiguous';
  return 'ok';
}
