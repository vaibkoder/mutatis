/**
 * Compare at scale — public surface (ADR-0037; spec/features/17-compare-batch.md).
 *
 * One-to-many today (SCOPE row 16); many-to-many (row 17) attaches here.
 *
 * Everything below is host-free and lives in the standalone/web target only.
 * The Word add-in has no batch surface, and that is structural, not an
 * oversight: ADR-0003's two-documents-in-memory `getFileAsync` ceiling makes an
 * N-way queue inside the task pane wrong by construction.
 */
export { BatchComparePanel } from './components/BatchComparePanel.js';
export type { BatchComparePanelProps } from './components/BatchComparePanel.js';

export { createBatchDriver } from './driver.js';
export { makeStubBatchDriverState } from './ui-types.js';
export type { BatchDriverState, BatchSlot, PairDetail } from './ui-types.js';

export { runOneToManyBatch, CROSS_FORMAT_REASON } from './runner.js';
export type {
  BatchDocumentInput,
  BatchProgress,
  BatchRun,
  RunOneToManyInput,
} from './runner.js';

export { buildBatchIndex, tallyBatch } from './index-model.js';
export { emptyClassificationTally } from './model.js';
export type {
  ArtefactHandle,
  BatchIndex,
  BatchMode,
  BatchRow,
  BatchRowStatus,
  BatchTotals,
} from './model.js';

export {
  downloadArtefact,
  downloadRowArchive,
  isDownloadable,
  redlineOf,
} from './artefacts.js';

export {
  BatchInputTooLargeError,
  MAX_BATCH_VARIANTS,
  MAX_BUNDLE_DOCUMENTS,
  MAX_DOCX_BYTES,
} from './limits.js';
export type { BatchCapKind } from './limits.js';
