/**
 * "Copy for AI" exports block for the Proof tab (feature 11 — spec §6).
 *
 * Builds the document-level ExportContext from the live Doc Model + Config, then
 * renders the composite Briefing Pack button and the per-checker Copy buttons.
 * Each per-checker button is enabled only when that checker's last run returned
 * data (its analysis is non-null). Pure presentation over already-computed data.
 */
import { findingsByCategory } from '@shared/findings/index.js';
import { DocModel } from '@shared/doc-model/index.js';
import { getStoredJurisdiction } from '@shared/config/index.js';
import {
  buildExport, buildBriefingPack, CopyButton, BriefingPackButton,
  type ExportContext, type AiExportEnvelope,
} from '@features/ai-export/index.js';
import { makeCiter } from '@features/ai-export/cite.js';
import type { DefinedTermsAnalysis } from '@features/defined-terms/index.js';
import type { CrossReferenceAnalysis } from '@features/cross-reference/index.js';
import type { FactualConsistencyAnalysis } from '@features/factual-consistency/index.js';
import type { PartyRoleAnalysis } from '@features/party-role/index.js';

export interface ProofExportsProps {
  defined: DefinedTermsAnalysis | null;
  xref: CrossReferenceAnalysis | null;
  factual: FactualConsistencyAnalysis | null;
  party: PartyRoleAnalysis | null;
}

function documentName(): string {
  try {
    if (typeof Office !== 'undefined' && Office.context?.document?.url) {
      const base = Office.context.document.url.split(/[\\/]/).pop();
      if (base) return base;
    }
  } catch { /* Office unavailable — fall through */ }
  return 'Untitled';
}

function jurisdiction(): ExportContext['jurisdiction'] {
  // Config stores 'UK' (England & Wales); the export envelope uses the 'EW' code.
  return getStoredJurisdiction() === 'SG' ? 'SG' : 'EW';
}

function makeContext(xref: CrossReferenceAnalysis | null): ExportContext {
  const paras = DocModel.allParagraphs.value;
  const textById = new Map(paras.map(p => [p.id, p.text]));
  return {
    generated: new Date().toISOString(),
    document: documentName(),
    jurisdiction: jurisdiction(),
    index: xref?.index ?? null,
    orderedParaIds: paras.map(p => p.id),
    paraText: (id) => textById.get(id) ?? '',
  };
}

/** Undefined-but-used terms (briefing-pack open items), read from the findings store. */
function undefinedTerms(ctx: ExportContext): Array<{ raw: string; clause: string }> {
  const findings = findingsByCategory('proof').get('defined-terms/used-not-defined') ?? [];
  const cite = makeCiter(ctx.index, ctx.orderedParaIds);
  return findings.map(f => {
    const m = /^"(.+?)" is used but never defined/.exec(f.title);
    return { raw: m?.[1] ?? f.title, clause: f.location.paraId ? cite(f.location.paraId) : '' };
  });
}

export function ProofExports({ defined, xref, factual, party }: ProofExportsProps) {
  const anyRun = defined || xref || factual || party;
  if (!anyRun) return null;

  const buildGlossary = (): AiExportEnvelope | null =>
    defined ? buildExport('defined-terms', { map: defined.definitionsMap, usage: defined.usageIndex }, makeContext(xref)) : null;

  const buildOutline = (): AiExportEnvelope | null =>
    xref ? buildExport('cross-reference', xref.index, makeContext(xref)) : null;

  const buildDealFacts = (): AiExportEnvelope | null =>
    factual ? buildExport('factual-consistency', {
      dates: factual.dates, amounts: factual.amounts, placeholders: factual.placeholders,
    }, makeContext(xref)) : null;

  const buildCastList = (): AiExportEnvelope | null =>
    party ? buildExport('party-role', { roles: party.roles, anchors: party.anchors }, makeContext(xref)) : null;

  const buildBriefing = (): AiExportEnvelope | null => {
    const ctx = makeContext(xref);
    return buildBriefingPack({
      ...(defined ? { definedTerms: { map: defined.definitionsMap, usage: defined.usageIndex } } : {}),
      ...(xref ? { sectionOutline: xref.index } : {}),
      ...(factual ? { dealFacts: { dates: factual.dates, amounts: factual.amounts, placeholders: factual.placeholders } } : {}),
      ...(party ? { castList: { roles: party.roles, anchors: party.anchors } } : {}),
      undefinedTerms: undefinedTerms(ctx),
    }, ctx);
  };

  return (
    <section class="dp-ai-export" aria-label="Copy for AI">
      <div class="dp-ai-export__header">
        <h2 class="dp-ai-export__heading">Copy for AI</h2>
        <BriefingPackButton build={buildBriefing} />
      </div>
      <ul class="dp-ai-export__list">
        <li class="dp-ai-export__item">
          <span class="dp-ai-export__label">Defined-terms glossary</span>
          <CopyButton build={buildGlossary} fileLabel="defined-terms"
            {...(defined ? {} : { disabledReason: 'Run Proof checks first' })} />
        </li>
        <li class="dp-ai-export__item">
          <span class="dp-ai-export__label">Section outline</span>
          <CopyButton build={buildOutline} fileLabel="section-outline"
            {...(xref ? {} : { disabledReason: 'Run Proof checks first' })} />
        </li>
        <li class="dp-ai-export__item">
          <span class="dp-ai-export__label">Deal facts</span>
          <CopyButton build={buildDealFacts} fileLabel="deal-facts"
            {...(factual ? {} : { disabledReason: 'Run Proof checks first' })} />
        </li>
        <li class="dp-ai-export__item">
          <span class="dp-ai-export__label">Party / role cast list</span>
          <CopyButton build={buildCastList} fileLabel="cast-list"
            {...(party ? {} : { disabledReason: 'Run Proof checks first' })} />
        </li>
      </ul>
    </section>
  );
}
