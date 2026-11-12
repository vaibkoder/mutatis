/**
 * Party-Role Consistency analysis (feature 09 — ADR-0016).
 *
 * Orchestrates PR1 + PR2 over the in-memory Document Model snapshot, REUSING the
 * defined-terms harvester (DefinitionsMap) and the shared Section Index. Issues
 * NO document walk of its own — one Document Model load feeds this checker and
 * Defined Terms in the same Proof run. Parity with `defined-terms/analyse.ts`.
 *
 * Pure (no findings-store mutation) — kept apart from `driver.ts` so a
 * consumer that only wants the analysis (e.g. the Compare-at-scale pairing
 * sketch, ADR-0037) never pulls in `@shared/findings`'s store module and the
 * `@preact/signals` it carries. The defined-terms harvester is imported from
 * its pure submodules directly, not the defined-terms barrel (`index.js`),
 * for the same reason: the barrel also re-exports that feature's own
 * store-touching `clearThenEmit`.
 */
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import { buildSectionIndex } from '@shared/section-index/index.js';
import { buildDefinitionsMap } from '@features/defined-terms/definitions.js';
import { mergeTableDefinitions } from '@features/defined-terms/tables.js';
import { makeLocator } from '@features/defined-terms/locate.js';
import type { RoleRecord, PartyAnchor } from './types.js';
import { buildRoleRecords } from './roles.js';
import { partiesRegionParaIds } from './region.js';
import { checkUndefinedRoles } from './pr1.js';
import { checkInconsistentLabels, scanPartyAnchors } from './pr2.js';

export interface PartyRoleAnalysis {
  findings: Finding[];
  /** PR1 — roles used but never introduced. */
  undefinedRoleCount: number;
  /** PR2 — same-party-two-labels findings. */
  inconsistentLabelCount: number;
  /** The per-role model, retained for the AI-export cast-list serialiser (feature 11). */
  roles: RoleRecord[];
  /** `<Entity> (the "Role")` introductions — the cast-list `entities` source. */
  anchors: PartyAnchor[];
}

/**
 * Run PR1 + PR2 and return findings + counts. Pure (no findings-store mutation).
 */
export function analysePartyRole(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): PartyRoleAnalysis {
  const bodyParas = paras.filter(p => p.container === 'body');
  const tableParas = paras.filter(p => p.container === 'tableCell');
  const footnoteParas = paras.filter(p => p.container === 'footnote');

  // REUSE: harvest definitions exactly as Defined Terms does (no re-scan logic).
  const map = buildDefinitionsMap(bodyParas, footnoteParas, profile);
  mergeTableDefinitions(map, tableParas, profile);

  // REUSE: shared Section Index to locate the parties/definitions region.
  const index = buildSectionIndex(paras, profile);
  const preamble = partiesRegionParaIds(index, paras, profile);

  const records = buildRoleRecords(map, bodyParas, footnoteParas, profile, preamble);
  const locate = makeLocator();

  const pr1 = checkUndefinedRoles(records, profile, locate);
  const pr2 = checkInconsistentLabels(paras, profile, locate);

  return {
    findings: [...pr2, ...pr1],
    undefinedRoleCount: pr1.length,
    inconsistentLabelCount: pr2.length,
    roles: [...records.values()],
    anchors: scanPartyAnchors(paras, profile),
  };
}

/** Findings-only convenience wrapper (parity with the other checkers). */
export function runPartyRole(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): Finding[] {
  return analysePartyRole(paras, profile).findings;
}
