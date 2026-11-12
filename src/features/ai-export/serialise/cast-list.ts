/**
 * Cast-list serialiser (feature 11 — spec §4.6).
 *
 * Projects the party-role model into a cast list: each role, the entity names
 * observed against it, where it was introduced, how often it is used, and any
 * consistency finding the checker raised.
 */
import type { RoleRecord, PartyAnchor } from '@features/party-role/index.js';
import type { CastListExport, PartyEntry } from '../types.js';
import type { Citer } from '../cite.js';

export interface CastListInput {
  roles: readonly RoleRecord[];
  anchors: readonly PartyAnchor[];
}

export interface CastListDeps {
  cite: Citer;
}

export function serialiseCastList(input: CastListInput, deps: CastListDeps): CastListExport {
  const entitiesByRole = groupEntities(input.anchors);
  const inconsistent = inconsistentRoleKeys(input.anchors);

  const parties = input.roles.map<PartyEntry>(rec => {
    const entry: PartyEntry = {
      role: rec.label,
      entities: entitiesByRole.get(rec.key) ?? [],
      body_uses: rec.uses.length,
    };

    if (rec.introduced && rec.introSites.length > 0) {
      entry.introduced_at = deps.cite(rec.introSites[0]!.paragraphId);
    }

    if (!rec.introduced && rec.uses.length > 0) {
      entry.finding = 'used-not-introduced';
    } else if (inconsistent.has(rec.key)) {
      entry.finding = 'inconsistent-label';
    }

    return entry;
  });

  return { parties };
}

/** Distinct entity names observed for each role key (from "(the X)" introductions). */
function groupEntities(anchors: readonly PartyAnchor[]): Map<string, string[]> {
  const out = new Map<string, Set<string>>();
  for (const a of anchors) {
    const set = out.get(a.roleKey) ?? new Set<string>();
    set.add(a.entityText);
    out.set(a.roleKey, set);
  }
  return new Map([...out].map(([k, v]) => [k, [...v]]));
}

/** Role keys that share an introducing entity with another role (PR2 Form-1 co-reference). */
function inconsistentRoleKeys(anchors: readonly PartyAnchor[]): Set<string> {
  const rolesByEntity = new Map<string, Set<string>>();
  for (const a of anchors) {
    const set = rolesByEntity.get(a.entityKey) ?? new Set<string>();
    set.add(a.roleKey);
    rolesByEntity.set(a.entityKey, set);
  }
  const flagged = new Set<string>();
  for (const roles of rolesByEntity.values()) {
    if (roles.size > 1) for (const r of roles) flagged.add(r);
  }
  return flagged;
}
