/**
 * Parties-preamble region (feature 09 — ADR-0016, spec §5.2).
 *
 * The parties block ("BETWEEN: (1) ACME … and (2) BIDCO …") and the recitals
 * usually precede Clause 1 and may carry no definitional verb, so a role can be
 * introduced there without landing in the DefinitionsMap. We locate that region
 * via the shared Section Index — paragraphs before the first clause, plus the
 * Definitions/Interpretation clause's own span — and treat roles named there as
 * introduced. Pure read of an already-built index; no document walk.
 */
import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { SectionIndex } from '@shared/section-index/index.js';

const DEFINITIONS_HEADING = /\b(?:Definitions?|Interpretation)\b/i;

export function partiesRegionParaIds(
  index: SectionIndex,
  paras: ReadonlyArray<ParaNode>,
  _profile: JurisdictionProfile,
): Set<string> {
  const idxById = new Map<string, number>();
  const byIndex = new Map<number, ParaNode>();
  for (const p of paras) {
    idxById.set(p.id, p.index);
    byIndex.set(p.index, p);
  }

  const ids = new Set<string>();
  const clauseNodes = index.nodes.filter(n => n.kind === 'clause');

  // 1. Preamble: every body paragraph before the first clause (BETWEEN lines +
  //    recitals). If the document has no clause structure at all, there is no
  //    preamble boundary — leave it empty so uses are still counted.
  if (clauseNodes.length > 0) {
    const firstClauseIdx = Math.min(
      ...clauseNodes.map(n => idxById.get(n.paraId) ?? Number.POSITIVE_INFINITY),
    );
    for (const p of paras) {
      if (p.container === 'body' && p.index < firstClauseIdx) ids.add(p.id);
    }

    // 2. Definitions / Interpretation clause span: from a top-level clause whose
    //    heading says "Definitions"/"Interpretation" up to the next top clause.
    const topClauses = clauseNodes
      .filter(n => n.level === 0)
      .map(n => ({ idx: idxById.get(n.paraId) ?? -1 }))
      .filter(x => x.idx >= 0)
      .sort((a, b) => a.idx - b.idx);

    for (let i = 0; i < topClauses.length; i++) {
      const startIdx = topClauses[i]!.idx;
      const heading = byIndex.get(startIdx)?.text ?? '';
      if (!DEFINITIONS_HEADING.test(heading)) continue;
      const endIdx = topClauses[i + 1]?.idx ?? Number.POSITIVE_INFINITY;
      for (const p of paras) {
        if (p.container === 'body' && p.index >= startIdx && p.index < endIdx) ids.add(p.id);
      }
    }
  }

  return ids;
}
