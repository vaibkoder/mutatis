import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { extractParts } from '@features/compare/io/parts.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { parseNumberingScheme } from '@shared/numbering/numbering-scheme.js';
import { buildRenderModel, type RenderInput } from '@shared/render-core/input/index.js';
import { getCompareConfig, SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { ChangeClass } from '@features/compare/model.js';
import type { RcChangeClass } from '@shared/render-core/types.js';
import type { CompareOutcome } from './protocol.js';
import { buildRedlineContext } from '../ai/context.js';
import { buildOutline, type StructureOutline } from '../structure/outline.js';

/**
 * One Word-vs-Word comparison, end to end. Pure enough to test directly.
 *
 * IMPORT DISCIPLINE — the two rules that matter:
 *   - `@shared/ooxml/bundle.js`, NOT the `@shared/ooxml` barrel: the barrel
 *     re-exports the Office-coupled service.
 *   - `@shared/render-core/input/index.js` + `.../types.js`, NOT the render-core
 *     barrel: the barrel pulls in the PDF output path (@pdfme/pdf-lib, fontkit)
 *     and the DOM preview, neither of which exists in a worker and neither of
 *     which the flowed-prose redline needs.
 *
 * Also never `getActiveProfile()` — it reads localStorage, which a worker has
 * no access to. The profile is passed in.
 */

export interface ComparisonInput {
  readonly originalBytes: ArrayBuffer | Uint8Array;
  readonly revisedBytes: ArrayBuffer | Uint8Array;
  readonly jurisdiction?: 'UK' | 'SG';
  readonly author?: string;
  readonly date?: string;
  /** File names, used for the AI context payload's baseline/revised labels. */
  readonly originalName?: string;
  readonly revisedName?: string;
  readonly onProgress?: (pct: number) => void;
  /** Duck-typed on purpose — the pipeline polls `.aborted`, not an AbortSignal. */
  readonly abortSignal?: { aborted: boolean };
}

/**
 * Compare's `ChangeClass` and render-core's `RcChangeClass` are the same eight
 * member names declared in two modules. Upstream's `toRcClassMap` does this
 * rename, but it sits beside code that constructs a Worker and touches
 * `document`, so it is replicated rather than imported.
 */
function toRcClassMap(
  source: ReadonlyMap<string, readonly ChangeClass[]>,
): ReadonlyMap<string, readonly RcChangeClass[]> {
  const out = new Map<string, readonly RcChangeClass[]>();
  for (const [id, classes] of source) out.set(id, classes as readonly RcChangeClass[]);
  return out;
}

export async function runComparison(input: ComparisonInput): Promise<CompareOutcome> {
  const originalBundle = new PizZipBundle(input.originalBytes);
  const revisedBundle = new PizZipBundle(input.revisedBytes);

  const originalParts = extractParts(originalBundle);
  const revisedParts = extractParts(revisedBundle);

  // The structure of each document, read from its own bytes. Never allowed to
  // stop a comparison: a document the analyser cannot read still gets its redline.
  const structure = {
    original: outlineOrNull(input.originalBytes),
    revised: outlineOrNull(input.revisedBytes),
  };

  const profile = input.jurisdiction === 'SG' ? SG_PROFILE : UK_PROFILE;
  const config = getCompareConfig(profile);

  const result = await runPipeline({
    originalParts,
    revisedParts,
    config,
    author: input.author ?? 'Mutatis Desktop',
    date: input.date ?? new Date().toISOString(),
    ...(input.onProgress ? { onProgress: input.onProgress } : {}),
    ...(input.abortSignal ? { abortSignal: input.abortSignal } : {}),
  });

  // Numbering resolves off the REVISED bundle: the redline's clause numbers
  // should read as the document the reviewer is moving towards.
  // parseNumberingScheme returns an empty scheme when numbering.xml is absent,
  // so this needs no guard.
  const numbering = parseNumberingScheme(revisedBundle);

  const renderInput: RenderInput = {
    authoredBodyXml: result.authoredBodyXml,
    ...(revisedParts.stylesXml !== null ? { stylesXml: revisedParts.stylesXml } : {}),
    ...(revisedParts.footnotesXml !== null ? { footnotesXml: revisedParts.footnotesXml } : {}),
    numbering,
    classByRevId: toRcClassMap(result.classByRevId),
  };

  return {
    model: buildRenderModel(renderInput),
    summary: result.summary,
    ops: result.ops,
    authoredBodyXml: result.authoredBodyXml,
    quarantinedOriginal: result.quarantinedOriginal,
    quarantinedRevised: result.quarantinedRevised,
    structure,
    // Built HERE because originalParas/revisedParas exist only inside this
    // function. They are far too large to ship back to the UI thread just so it
    // could derive the same payload.
    context: buildRedlineContext({
      ops: result.ops,
      summary: result.summary,
      originalParas: result.originalParas,
      revisedParas: result.revisedParas,
      baseline: input.originalName ?? 'Original',
      revised: input.revisedName ?? 'Modified',
      jurisdiction: input.jurisdiction === 'SG' ? 'SG' : 'UK',
      quarantinedOriginal: result.quarantinedOriginal,
      quarantinedRevised: result.quarantinedRevised,
      generated: input.date ?? new Date().toISOString(),
      structure,
    }),
  };
}

function outlineOrNull(bytes: ArrayBuffer | Uint8Array): StructureOutline | null {
  try {
    return buildOutline(bytes);
  } catch {
    return null;
  }
}
