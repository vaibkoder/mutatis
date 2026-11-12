/**
 * Feature 08 — Compare: output layer barrel (C1 + C2).
 * Pure OOXML authoring (revision markup, body author, summary annotations) plus
 * the materialise write-out that delegates to the OOXML Service.
 */
export {
  wrapIns,
  wrapDel,
  toDelText,
  wrapMoveFrom,
  wrapMoveTo,
  wrapParaRuns,
} from './revision-markup.js';

export {
  authorBody,
  makeAuthorContext,
  splitParagraph,
  parseParagraph,
} from './author.js';
export type { AuthorContext } from './author.js';

export { buildSummaryAnnotations } from './annotations.js';

export {
  materialise, replaceBody, assembleRedlinePackage, trailingSectPr, relationshipIdAllocator,
} from './writeout.js';
export type { WriteoutOptions, AuxiliaryParts, NewPart, PartRels } from './writeout.js';
