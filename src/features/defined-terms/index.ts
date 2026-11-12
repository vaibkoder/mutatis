export { analyseDefinedTerms, runDefinedTerms } from './analyse.js';
export type { DefinedTermsOptions, DefinedTermsAnalysis } from './analyse.js';
export { clearThenEmit } from './driver.js';
export type {
  DefinitionsMap, DefinitionEntry, DefinitionRecord,
  UsageIndex, UsageRecord, DefinitionSource, DefinitionShape, TextAnchor,
} from './types.js';

// Harvester reuse surface (consumed by feature 09 Party-Role Consistency —
// ADR-0016; reuse before build). These build the shared
// DefinitionsMap and provide term normalisation; party-role re-reads them
// rather than re-scanning the document.
export { buildDefinitionsMap } from './definitions.js';
export { mergeTableDefinitions } from './tables.js';
export { normaliseTerm, cleanTerm, resolveLookup, buildAlternation } from './normalise.js';
export { makeLocator, findingId, type Locator } from './locate.js';
