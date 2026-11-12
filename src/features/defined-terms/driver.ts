import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import { addFindings, removeByPrefix } from '@shared/findings/index.js';
import type { DefinedTermsOptions, DefinedTermsAnalysis } from './analyse.js';
import { analyseDefinedTerms } from './analyse.js';

/**
 * Clear previous defined-terms findings and emit fresh results.
 * Returns the analysis so the Proof card can render the cross-ref/cross-doc
 * banner and the "No definitions detected" state.
 *
 * The only store-touching function in this feature — kept apart from the pure
 * `analyseDefinedTerms`/`runDefinedTerms` in `analyse.ts` so an analysis-only
 * consumer (e.g. Compare-at-scale's pairing sketch, ADR-0037) never pulls in
 * `@shared/findings`'s store module, whose module-level `signal()` call drags
 * in `@preact/signals` and cannot be tree-shaken. Import `analyse.ts` directly
 * for that use; this barrel-adjacent split keeps every existing consumer of the
 * `index.ts` barrel unchanged.
 */
export function clearThenEmit(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  opts?: DefinedTermsOptions,
): DefinedTermsAnalysis {
  removeByPrefix('defined-terms');
  const analysis = analyseDefinedTerms(paras, profile, opts);
  addFindings(analysis.findings);
  return analysis;
}
