import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import { addFindings, removeByPrefix } from '@shared/findings/index.js';
import type { PartyRoleAnalysis } from './analyse.js';
import { analysePartyRole } from './analyse.js';

/**
 * Clear previous party-role findings and emit fresh results.
 *
 * The only store-touching function in this feature — kept apart from the pure
 * `analysePartyRole`/`runPartyRole` in `analyse.ts` so an analysis-only
 * consumer (e.g. Compare-at-scale's pairing sketch, ADR-0037) never pulls in
 * `@shared/findings`'s store module and the `@preact/signals` it carries.
 * Parity with `defined-terms/driver.ts`.
 */
export function clearThenEmit(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): PartyRoleAnalysis {
  removeByPrefix('party-role');
  const analysis = analysePartyRole(paras, profile);
  addFindings(analysis.findings);
  return analysis;
}
