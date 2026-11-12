/**
 * Diff layer barrel for Compare (Wave 1-A, tasks A2–A5).
 * Spec: spec/features/08-compare.md §5.1–5.4, §9.3/9.6.
 */
export { computeLCS } from './lcs.js';
export { structuralDiff } from './structural.js';
export { gapPairingFor, alignGap, MAX_ALIGN_WORK, MAX_ALIGN_WORK_TOTAL, MAX_ALIGN_CELLS, ALIGN_CELL_OVERHEAD } from './gap-pairing.js';
export type { GapPairing } from './gap-pairing.js';
export { tokenize, tokenizeTagged, atomicKindOf } from './atomic-tokens.js';
export type { Token } from './atomic-tokens.js';
export { wordDiff } from './word.js';
export { sideText } from './side-text.js';
export { insertBeforeDelete } from './mark-order.js';
export { slideEditBoundaries } from './anchor.js';
export { refineReplaceBlocks } from './refine.js';
export { cleanupPreservingRelocations } from './relocate.js';
export type { DmpDiff, TokenAt } from './relocate.js';
export { detectTransposition } from './transpose.js';
export { detectMoves } from './move.js';
export { detectPropertyChanges, comparePropertyPair } from './property.js';
export {
  demoteDissimilarReplaces,
  survivesDemotion,
  bagOf,
  bagDice,
  diceFromOverlap,
  paragraphSimilarity,
  contentTokens,
  contiguousShare,
  longestContiguousRun,
} from './replace-threshold.js';
export { detectMergesAndSplits, rangeDiffText, PARA_SEAM } from './merge-split.js';
export { diffTables } from './tables.js';
