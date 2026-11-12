export type {
  DocxBundle,
  OoxmlService,
  ReadOptions,
  WriteOptions,
} from './types.js';
// PizZipBundle is host-free and lives in its own module (bundle.ts) so importing
// it never pulls in the Office-coupled service (12f confidentiality split).
export { PizZipBundle } from './bundle.js';
export {
  OoxmlServiceImpl,
  ooxmlService,
  detectArtifactParagraph,
  toReadDocxError,
} from './service.js';
