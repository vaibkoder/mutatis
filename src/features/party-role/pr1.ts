/**
 * PR1 — Role used-but-never-introduced (feature 09 — ADR-0016, spec §5.5).
 *
 * A party-role label used at least `minUsesForUndefined` times in the operative
 * text but never introduced (no definition record, not in the parties preamble)
 * is flagged. The recurrence gate is the conservative control (cf. ADR-0015):
 * a role used once or twice is presumed incidental and suppressed.
 */
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { Locator } from '@features/defined-terms/index.js';
import type { RoleRecord, TextAnchor } from './types.js';
import { resolvePartyRoleConfig } from './types.js';
import { PR_CATEGORIES, makePartyFinding, anchorOrder } from './locate.js';

export function checkUndefinedRoles(
  records: ReadonlyMap<string, RoleRecord>,
  profile: JurisdictionProfile,
  locate: Locator,
): Finding[] {
  const cfg = resolvePartyRoleConfig(profile);
  const built: Array<{ finding: Finding; first: TextAnchor }> = [];

  for (const rec of records.values()) {
    if (rec.introduced) continue;
    if (rec.uses.length < cfg.minUsesForUndefined) continue;

    const sorted = [...rec.uses].sort(anchorOrder);
    const occurrences = sorted.map(a => locate(a, rec.label));
    const first = sorted[0]!;
    built.push({
      first,
      finding: makePartyFinding(
        PR_CATEGORIES.undefinedRole,
        cfg.pr1Severity,
        `'${rec.label}' is used ${rec.uses.length} times but never introduced as a party`,
        occurrences[0]!,
        first,
        `Define '${rec.label}' (e.g. in the parties clause or definitions), or confirm it is an ordinary word.`,
        occurrences,
      ),
    });
  }

  // Stable document order by first use (reuses the shared anchor comparator).
  return built.sort((a, b) => anchorOrder(a.first, b.first)).map(b => b.finding);
}
