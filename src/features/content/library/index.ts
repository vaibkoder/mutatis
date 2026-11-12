// Barrel for sub-feature B — Precedent / Clause Library.
export { LibraryPanel } from './LibraryPanel.js';
export type { ClauseEntry, ClauseIndexRecord, ClauseJurisdiction, DriftResult } from './types.js';
export { validateClauseMetadata, validateTitle, validateSelectionNonEmpty } from './validation.js';
export { putClause, getClause, deleteClause, loadIndex, exportLibrary, importLibrary, hasSharedBackend } from './library-store.js';
export { LibrarySearch, getLibrarySearch, _resetLibrarySearch } from './search.js';
export type { FacetFilters, SearchResult } from './search.js';
export { parseClauseTag, scanDrift, formatDateDDMonthYYYY } from './drift.js';
export {
  requiresJurisdictionWarning,
  jurisdictionLabel,
  docJurisdictionLabel,
  buildClauseTag,
  buildClauseTitle,
  xmlEscape,
  stripControlChars,
  normaliseQuotesAndEncoding,
  insertClause,
  getDocJurisdiction,
} from './insert.js';
