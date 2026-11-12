export type {
  CompareSourceKind,
  CompareSource,
  ActiveDocumentSource,
  DocxArrayBufferSource,
  TextSource,
  CompareDirection,
  LoadedSides,
} from './load.js';

export { loadBothSides } from './load.js';

export { synthesiseDocx, splitLines } from './text/synthesise-docx.js';
export type { SynthesiseOptions, BlankLinePolicy } from './text/synthesise-docx.js';

export type {
  CompareParts, PartRef, HeaderFooterKind, HeaderFooterType, HeaderFooterRole,
} from './parts.js';
export { extractParts } from './parts.js';
