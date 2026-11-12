/**
 * Build the per-role model (feature 09 — ADR-0016).
 *
 * A `RoleRecord` answers two questions for each party-role label in the config
 * vocabulary: was it *introduced* (defined anywhere, or named in the parties
 * preamble), and where is it *used*. Everything is read from the already-built
 * `DefinitionsMap` (the defined-terms harvester) and the cached paragraph text —
 * no second document walk. (spec §5.)
 */
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
// Pure `normalise.js` submodule directly, not the defined-terms barrel — see
// `locate.ts` for why.
import { buildAlternation, normaliseTerm, resolveLookup } from '@features/defined-terms/normalise.js';
import type { DefinitionsMap } from '@features/defined-terms/index.js';
import type { RoleRecord, TextAnchor } from './types.js';
import { resolvePartyRoleConfig } from './types.js';

interface Span { start: number; end: number; }

/** Canonical role surface forms from config, minus the generic collective roles. */
export function roleVocabulary(profile: JurisdictionProfile): string[] {
  const cfg = resolvePartyRoleConfig(profile);
  const collective = new Set(cfg.excludeCollectiveRoles.map(normaliseTerm));
  const seen = new Set<string>();
  const out: string[] = [];
  for (const label of profile.definedTerms.exclusions.partyRoles) {
    const key = normaliseTerm(label);
    if (!key || collective.has(key) || seen.has(key)) continue;
    seen.add(key);
    out.push(label.trim());
  }
  return out;
}

function anchorOf(paragraphId: string, start: number, len: number): TextAnchor {
  return { paragraphId, startOffset: start, endOffset: start + len };
}

/** Map a surface form to a role key in `records` (exact, then plural/possessive). */
function resolveRole(records: ReadonlyMap<string, RoleRecord>, surface: string): string | null {
  const key = normaliseTerm(surface);
  if (records.has(key)) return key;
  return resolveLookup(records, surface);
}

/**
 * Build the role model. `preambleParaIds` (from the Section Index, Clause 5.2)
 * marks the parties-preamble region: roles named there count as introduced, and
 * their occurrences are not counted as operative uses.
 */
export function buildRoleRecords(
  map: DefinitionsMap,
  bodyParas: ReadonlyArray<ParaNode>,
  footnoteParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  preambleParaIds: ReadonlySet<string> = new Set(),
): Map<string, RoleRecord> {
  const labels = roleVocabulary(profile);
  const records = new Map<string, RoleRecord>();
  for (const label of labels) {
    const key = normaliseTerm(label);
    records.set(key, { key, label, introduced: false, introSites: [], uses: [] });
  }
  if (records.size === 0) return records;

  // 1. Introductions recorded in the DefinitionsMap (inline parenthetical,
  //    verb-anchored, table, footnote, cross-reference — all already harvested).
  const defSpans = new Map<string, Span[]>();
  for (const entry of map.values()) {
    const roleKey = resolveRole(records, entry.canonical);
    if (roleKey) {
      const rec = records.get(roleKey)!;
      rec.introduced = true;
      for (const d of entry.definitions) rec.introSites.push(d.anchor);
    }
    // Track every definition span (for ALL terms) so a role use that sits inside
    // any definition's term span is excluded from operative uses.
    for (const d of entry.definitions) {
      const arr = defSpans.get(d.anchor.paragraphId) ?? [];
      arr.push({ start: d.anchor.startOffset, end: d.anchor.endOffset });
      defSpans.set(d.anchor.paragraphId, arr);
    }
  }

  const re = buildAlternation(labels);
  if (!re) return records;

  // 2. Preamble introductions — roles named in the parties region without a verb.
  for (const p of bodyParas) {
    if (!preambleParaIds.has(p.id)) continue;
    re.lastIndex = 0;
    for (let m = re.exec(p.text); m; m = re.exec(p.text)) {
      if (/^[a-z]/.test(m[0])) continue; // party labels are capitalised
      const roleKey = resolveRole(records, m[0]);
      if (!roleKey) continue;
      const rec = records.get(roleKey)!;
      rec.introduced = true;
      rec.introSites.push(anchorOf(p.id, m.index, m[0].length));
    }
  }

  // 3. Operative uses — body + footnote, EXCLUDING the preamble and any span that
  //    sits inside a definition's term span. Lowercase matches are ordinary words
  //    ("the buyer of goods"), never a party label — skipped.
  const scanUses = (paras: ReadonlyArray<ParaNode>): void => {
    for (const p of paras) {
      if (preambleParaIds.has(p.id)) continue;
      const spans = defSpans.get(p.id) ?? [];
      re.lastIndex = 0;
      for (let m = re.exec(p.text); m; m = re.exec(p.text)) {
        const start = m.index;
        const end = start + m[0].length;
        if (/^[a-z]/.test(m[0])) continue;
        if (spans.some(s => start < s.end && end > s.start)) continue;
        const roleKey = resolveRole(records, m[0]);
        if (!roleKey) continue;
        records.get(roleKey)!.uses.push(anchorOf(p.id, start, m[0].length));
      }
    }
  };
  scanUses(bodyParas);
  scanUses(footnoteParas);

  return records;
}
