import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { DefinitionsMap, UsageIndex, UsageRecord } from './types.js';
import { buildAlternation, resolveLookup } from './normalise.js';

interface Span { start: number; end: number; }

/**
 * Build the UsageIndex: every occurrence of a defined surface form across body,
 * table and footnote text, matched longest-first. The occurrence inside a term's
 * own definition (its quoted/candidate span) is flagged isDefinitionOccurrence
 * and excluded from "used" counts by the sub-checks. (spec §5.3)
 */
export function buildUsageIndex(
  map: DefinitionsMap,
  bodyParas: ReadonlyArray<ParaNode>,
  tableParas: ReadonlyArray<ParaNode>,
  footnoteParas: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
): UsageIndex {
  const idx: UsageIndex = new Map();
  const forms = Array.from(map.values()).map(e => e.canonical);
  const re = buildAlternation(forms);
  if (!re) return idx;

  // Definition term spans per paragraph, to mark definition occurrences.
  const defSpans = new Map<string, Span[]>();
  for (const entry of map.values()) {
    for (const d of entry.definitions) {
      const arr = defSpans.get(d.anchor.paragraphId) ?? [];
      arr.push({ start: d.anchor.startOffset, end: d.anchor.endOffset });
      defSpans.set(d.anchor.paragraphId, arr);
    }
  }

  const determiners = new Set(profile.definedTerms.determiners.map(d => d.toLowerCase()));

  const scan = (paras: ReadonlyArray<ParaNode>, inFootnote: boolean): void => {
    for (const p of paras) {
      re.lastIndex = 0;
      for (let m = re.exec(p.text); m; m = re.exec(p.text)) {
        const surface = m[0];
        const start = m.index;
        const end = start + surface.length;
        const key = resolveLookup(map, surface);
        if (!key) continue;

        const pre = p.text.slice(0, start).replace(/\s+$/, '');
        const preWord = (pre.match(/(\S+)$/)?.[1] ?? '').toLowerCase().replace(/[^a-z'’]/g, '');
        const isSentenceInitial = pre === '' || /[.!?:\n]$/.test(pre);
        const isLowercase = /^[a-z]/.test(surface);
        const isDefinitionOccurrence = (defSpans.get(p.id) ?? [])
          .some(s => start < s.end && end > s.start);

        const rec: UsageRecord = {
          surface,
          isLowercase,
          precededByDeterminer: determiners.has(preWord),
          isSentenceInitial,
          anchor: { paragraphId: p.id, startOffset: start, endOffset: end },
          inFootnote,
          isDefinitionOccurrence,
        };
        const arr = idx.get(key) ?? [];
        arr.push(rec);
        idx.set(key, arr);
      }
    }
  };

  scan(bodyParas, false);
  scan(tableParas, false);
  scan(footnoteParas, true);
  return idx;
}
