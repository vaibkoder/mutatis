/**
 * PR2 — Inconsistent label / same party, two names (feature 09 — ADR-0016, spec §5.3–5.4).
 *
 * Fires ONLY on explicit textual evidence of co-reference — never on the mere
 * co-presence of two role words (a genuine dual-role deal must not be flagged):
 *  - Form 1: two distinct role labels introduced against the SAME entity name
 *            (`ACME Pte Ltd (the "Buyer")` … `ACME Pte Ltd (the "Purchaser")`).
 *  - Form 2: one role defined as the other (`"Purchaser" means the Buyer`).
 *
 * Both scans read the cached paragraph text — no document walk.
 */
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding, FindingLocation } from '@shared/findings/index.js';
import type { Locator } from '@features/defined-terms/index.js';
// Pure `normalise.js` submodule directly, not the defined-terms barrel — see
// `locate.ts` for why.
import { normaliseTerm } from '@features/defined-terms/normalise.js';
import type { PartyAnchor, TextAnchor } from './types.js';
import { resolvePartyRoleConfig } from './types.js';
import { roleVocabulary } from './roles.js';
import { PR_CATEGORIES, makePartyFinding, anchorOrder } from './locate.js';

const QUOTE = `["'“”‘’]?`;

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Embeddable role alternation source, longest-first so "Disclosing Party" wins. */
function roleAltSource(labels: readonly string[]): string {
  return [...labels].sort((a, b) => b.length - a.length).map(escapeRegExp).join('|');
}

/** Fold an entity name to a comparison key: lowercase, drop dots, strip trailing suffix. */
function normaliseEntity(raw: string, suffixes: readonly string[]): string {
  let s = raw.toLowerCase().replace(/\./g, ' ').replace(/[",;]+/g, ' ').replace(/\s+/g, ' ').trim();
  const normSuffixes = suffixes
    .map(x => x.toLowerCase().replace(/\./g, ' ').replace(/\s+/g, ' ').trim())
    .filter(x => x.length > 0)
    .sort((a, b) => b.length - a.length);
  let changed = true;
  while (changed) {
    changed = false;
    for (const suf of normSuffixes) {
      if (s === suf) return '';
      if (s.endsWith(' ' + suf)) {
        s = s.slice(0, s.length - suf.length - 1).trim();
        changed = true;
        break;
      }
    }
  }
  return s;
}

/**
 * The Title-Case entity run immediately preceding an index, if any. Tokens are
 * separated by whitespace only — a comma breaks the run, so a sentence lead-in
 * ("Separately, ACME Pte Ltd") does not get absorbed into the entity key (which
 * would split one party across two keys and miss a real Form-1 slip).
 */
function entityBefore(text: string, before: number): { entityText: string; offset: number } | null {
  const head = text.slice(0, before).replace(/\s+$/, '');
  const m = /([A-Z][A-Za-z0-9&.'’-]*(?:\s+[A-Z][A-Za-z0-9&.'’-]*){0,6})\s*$/.exec(head);
  if (!m || !m[1]) return null;
  return { entityText: m[1].trim(), offset: m.index };
}

function anchorOf(paragraphId: string, start: number, len: number): TextAnchor {
  return { paragraphId, startOffset: start, endOffset: start + len };
}

/** Form 1 raw material — every `<Entity> (the "Role")` introduction. */
export function scanPartyAnchors(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): PartyAnchor[] {
  const labels = roleVocabulary(profile);
  if (labels.length === 0) return [];
  const suffixes = profile.entityFields.suffixes;
  const alt = roleAltSource(labels);
  const re = new RegExp(`\\(\\s*(?:the|each|a|an|this)?\\s*${QUOTE}(${alt})${QUOTE}\\s*\\)`, 'gi');

  const out: PartyAnchor[] = [];
  for (const p of paras) {
    if (p.container !== 'body') continue;
    re.lastIndex = 0;
    for (let m = re.exec(p.text); m; m = re.exec(p.text)) {
      const roleSurface = m[1] ?? '';
      if (/^[a-z]/.test(roleSurface)) continue;
      const ent = entityBefore(p.text, m.index);
      if (!ent) continue;
      const entityKey = normaliseEntity(ent.entityText, suffixes);
      if (!entityKey) continue;
      out.push({
        entityKey,
        entityText: ent.entityText,
        roleKey: normaliseTerm(roleSurface),
        roleLabel: roleSurface,
        anchor: anchorOf(p.id, m.index, m[0].length),
      });
    }
  }
  return out;
}

export function checkInconsistentLabels(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  locate: Locator,
): Finding[] {
  const cfg = resolvePartyRoleConfig(profile);
  const labels = roleVocabulary(profile);
  if (labels.length === 0) return [];
  const labelByKey = new Map(labels.map(l => [normaliseTerm(l), l]));

  const findings: Finding[] = [];
  const emitted = new Set<string>();

  const pairKey = (a: string, b: string): string => [a, b].sort().join('|');

  // -- Form 1: shared introducing entity ------------------------------------
  const anchors = scanPartyAnchors(paras, profile);
  const byEntity = new Map<string, PartyAnchor[]>();
  for (const a of anchors) {
    const arr = byEntity.get(a.entityKey) ?? [];
    arr.push(a);
    byEntity.set(a.entityKey, arr);
  }
  for (const [, group] of byEntity) {
    const roleKeys = Array.from(new Set(group.map(a => a.roleKey)));
    if (roleKeys.length < 2) continue;
    const key = roleKeys.slice().sort().join('|');
    if (emitted.has(key)) continue;
    emitted.add(key);
    // For >2 roles also block the pairwise keys (avoid Form-2 duplicates).
    for (let i = 0; i < roleKeys.length; i++) {
      for (let j = i + 1; j < roleKeys.length; j++) emitted.add(pairKey(roleKeys[i]!, roleKeys[j]!));
    }

    const sites = [...group].sort((x, y) => anchorOrder(x.anchor, y.anchor));
    const occ: FindingLocation[] = sites.map(s => locate(s.anchor, `${s.roleLabel} — ${s.entityText}`));
    const labelList = roleKeys.map(k => `'${labelByKey.get(k) ?? k}'`).join(' and ');
    findings.push(
      makePartyFinding(
        PR_CATEGORIES.inconsistentLabel,
        cfg.pr2Severity,
        `${labelList} appear to be the same party`,
        occ[0]!,
        sites[0]!.anchor,
        `Introduced against the same party "${group[0]!.entityText}" — use one consistent label throughout.`,
        occ,
      ),
    );
  }

  // -- Form 2: definition-by-aliasing ("Purchaser" means the Buyer) ----------
  const alt = roleAltSource(labels);
  const f2 = new RegExp(
    `(?<![A-Za-z0-9])${QUOTE}(${alt})${QUOTE}\\s*(?:means|shall mean)\\s+(?:the\\s+)?(${alt})(?![A-Za-z0-9])\\s*(?=[.;,)]|$)`,
    'gi',
  );
  for (const p of paras) {
    if (p.container === 'tableCell') continue;
    f2.lastIndex = 0;
    for (let m = f2.exec(p.text); m; m = f2.exec(p.text)) {
      const aSurface = m[1] ?? '';
      const bSurface = m[2] ?? '';
      if (/^[a-z]/.test(aSurface) || /^[a-z]/.test(bSurface)) continue;
      const aKey = normaliseTerm(aSurface);
      const bKey = normaliseTerm(bSurface);
      if (aKey === bKey) continue;
      const key = pairKey(aKey, bKey);
      if (emitted.has(key)) continue;
      emitted.add(key);

      const aAnchor = anchorOf(p.id, m.index, aSurface.length + 1);
      const bStart = p.text.indexOf(bSurface, m.index + aSurface.length);
      const bAnchor = anchorOf(p.id, bStart >= 0 ? bStart : m.index, bSurface.length);
      const occ: FindingLocation[] = [locate(aAnchor, aSurface), locate(bAnchor, bSurface)];
      findings.push(
        makePartyFinding(
          PR_CATEGORIES.inconsistentLabel,
          cfg.pr2Severity,
          `'${labelByKey.get(aKey) ?? aSurface}' and '${labelByKey.get(bKey) ?? bSurface}' appear to be the same party`,
          occ[0]!,
          aAnchor,
          `One label is defined as the other ("${aSurface}" means the ${bSurface}) — use one consistent label.`,
          occ,
        ),
      );
    }
  }

  return findings;
}
