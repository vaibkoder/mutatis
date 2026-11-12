import type { ParaNode } from '@shared/doc-model/index.js';
import type { JurisdictionProfile } from '@shared/config/index.js';
import type { Finding } from '@shared/findings/index.js';
import type { DefinitionsMap, UsageIndex } from './types.js';
import { buildDefinitionsMap } from './definitions.js';
import { mergeTableDefinitions } from './tables.js';
import { buildUsageIndex } from './usage.js';
import {
  checkUsedNotDefined,
  checkDefinedNotUsed,
  checkDefinedMultiple,
  checkCapitalisation,
} from './subchecks.js';
import { makeLocator } from './locate.js';

export interface DefinedTermsOptions {
  /** Override the configured capMode for the smart/strict toggle re-run. */
  capModeOverride?: 'smart' | 'strict';
}

export interface DefinedTermsAnalysis {
  findings: Finding[];
  /** Distinct terms defined by cross-reference within this document. */
  crossRefCount: number;
  /** Distinct terms defined by cross-reference to an external document (never resolved). */
  crossDocCount: number;
  /** Number of distinct defined terms detected. */
  definitionCount: number;
  /** True when no definitions were detected → "No definitions detected" (spec §8 / T15). */
  noDefinitions: boolean;
  /** The capitalisation mode actually applied. */
  capMode: 'smart' | 'strict';
  /**
   * The harvested definitions, retained for the AI-export serialiser (feature 11).
   * Empty map when no definitions were detected. No new scanning — this is the
   * same map the sub-checks consumed.
   */
  definitionsMap: DefinitionsMap;
  /**
   * The usage index, retained for AI-export variant detection (F5). Empty map on
   * the no-definitions early return (usage is not built in that case).
   */
  usageIndex: UsageIndex;
}

function countCrossRefs(map: DefinitionsMap): { crossRefCount: number; crossDocCount: number } {
  let crossRefCount = 0;
  let crossDocCount = 0;
  for (const entry of map.values()) {
    if (entry.definitions.some(d => d.isCrossRef)) crossRefCount += 1;
    if (entry.definitions.some(d => d.isCrossDoc)) crossDocCount += 1;
  }
  return { crossRefCount, crossDocCount };
}

/**
 * Run the four Defined-Terms sub-checks over the Document-Model paragraphs and
 * return findings plus banner counts. Pure (no findings-store mutation) — kept
 * apart from `driver.ts` so a consumer that only wants the analysis (e.g. the
 * Compare-at-scale pairing sketch, ADR-0037) never pulls in `@shared/findings`'s
 * store module and the `@preact/signals` it carries. See `driver.ts` for why.
 *
 * Paragraphs are split by container: body prose, table cells, footnote bodies —
 * all from the single in-memory snapshot (no document re-walk).
 */
export function analyseDefinedTerms(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  opts: DefinedTermsOptions = {},
): DefinedTermsAnalysis {
  const bodyParas = paras.filter(p => p.container === 'body');
  const tableParas = paras.filter(p => p.container === 'tableCell');
  const footnoteParas = paras.filter(p => p.container === 'footnote');

  const map = buildDefinitionsMap(bodyParas, footnoteParas, profile);
  mergeTableDefinitions(map, tableParas, profile);

  const capMode = opts.capModeOverride ?? profile.definedTerms.capMode;
  const { crossRefCount, crossDocCount } = countCrossRefs(map);
  const definitionCount = map.size;

  // Empty-map case: report "No definitions detected" rather than flagging every
  // capitalised word; sub-checks 2–4 suppressed and sub-check 1 kept quiet.
  if (definitionCount === 0) {
    return {
      findings: [],
      crossRefCount,
      crossDocCount,
      definitionCount,
      noDefinitions: true,
      capMode,
      definitionsMap: map,
      usageIndex: new Map(),
    };
  }

  const usage = buildUsageIndex(map, bodyParas, tableParas, footnoteParas, profile);
  const locate = makeLocator();

  const findings: Finding[] = [
    ...checkUsedNotDefined(map, bodyParas, tableParas, profile, locate),
    ...checkDefinedNotUsed(map, usage, locate),
    ...checkDefinedMultiple(map, profile, locate),
    ...checkCapitalisation(map, usage, capMode, profile, locate, paras),
  ];

  return {
    findings,
    crossRefCount,
    crossDocCount,
    definitionCount,
    noDefinitions: false,
    capMode,
    definitionsMap: map,
    usageIndex: usage,
  };
}

/** Findings-only convenience wrapper (parity with factual-consistency). */
export function runDefinedTerms(
  paras: ReadonlyArray<ParaNode>,
  profile: JurisdictionProfile,
  opts?: DefinedTermsOptions,
): Finding[] {
  return analyseDefinedTerms(paras, profile, opts).findings;
}
