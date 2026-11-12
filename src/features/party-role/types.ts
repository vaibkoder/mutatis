/**
 * Party-Role Consistency — data model (feature 09, ADR-0016).
 * Built per analysis run and discarded; no persistence. Pure types + the
 * code-side config default so a profile that omits `partyRoleConsistency`
 * still works.
 */
import type { TextAnchor } from '@shared/doc-model/anchor.js';
import type { JurisdictionProfile, PartyRoleConsistencyConfig } from '@shared/config/index.js';

export type { TextAnchor };

/** One party-role label, with where it was introduced and where it is used. */
export interface RoleRecord {
  /** normaliseTerm(label), e.g. "purchaser". */
  key: string;
  /** Canonical surface form, e.g. "Purchaser". */
  label: string;
  /** Introduced = defined (any DefinitionsMap record) OR present in the parties-preamble region. */
  introduced: boolean;
  /** Introduction sites (definition anchors / preamble occurrence) for navigation. */
  introSites: TextAnchor[];
  /** Body+footnote uses, excluding occurrences inside an introduction. */
  uses: TextAnchor[];
}

/** One `<Entity> (the "Role")` introduction — the raw material for PR2 Form 1. */
export interface PartyAnchor {
  /** Normalised preceding entity name (suffix-folded), e.g. "acme". */
  entityKey: string;
  /** Surface entity name, e.g. "ACME Pte Ltd". */
  entityText: string;
  /** normaliseTerm(role). */
  roleKey: string;
  /** Canonical role label. */
  roleLabel: string;
  /** The parenthetical's anchor. */
  anchor: TextAnchor;
}

/** Conservative defaults (spec §6) when a profile omits the block. */
export const DEFAULT_PARTY_ROLE_CONFIG: PartyRoleConsistencyConfig = {
  minUsesForUndefined: 3,
  pr1Severity: 'warning',
  pr2Severity: 'warning',
  excludeCollectiveRoles: ['Party', 'Parties'],
};

/** Resolve the effective config for a profile (its own block, or the defaults). */
export function resolvePartyRoleConfig(profile: JurisdictionProfile): PartyRoleConsistencyConfig {
  return profile.partyRoleConsistency ?? DEFAULT_PARTY_ROLE_CONFIG;
}
