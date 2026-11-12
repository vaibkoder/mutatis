/**
 * Feature 08 — Compare: one auxiliary part — a header, a footer, one footnote or
 * endnote — analysed and authored by the BODY's stages (ADR-0047).
 *
 * There is no second pipeline: a part's paragraphs go through `preprocessSide`,
 * `analyseRecords` (alignment, merge/split, demotion, properties, Word tables,
 * word diff, moves, classification, mark order) and `authorBlocks`, exactly as
 * the body's do, with the caller's `AuthorContext` so every `w:id` in the
 * package stays unique.
 *
 * Self-check (ADR-0008): the authored blocks must reconstruct both sides —
 * accept-all the revised part, reject-all the original — by the 08i yardstick
 * (`reconstruction.ts`), in text and paragraph count. A part that fails is
 * NOT authored (`blocksXml: null`): the caller keeps the original part, as
 * before ADR-0047, and its ops are still listed — reported, never dropped.
 *
 * Spec: spec/features/08-compare.md §5.5, §9.7, §9.8.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { DiffOp, ParaRecord } from '../model.js';
import { analyseRecords } from '../pipeline.js';
import { preprocessSide, flattenTrackedChanges } from '../preprocess/index.js';
import { authorBlocks, type AuthorContext } from './author.js';
import { reconstruct, textLost, type ViewReport } from './reconstruction.js';

export interface PartInput {
  /** The part's id — a role id (`header:s0:default`) or a note id (`footnote:3`). */
  partId: string;
  /** The original side's part XML (whole part or one note element); null ⇒ absent. */
  originalXml: string | null;
  revisedXml: string | null;
  config: CompareConfig;
  author: string;
  date: string;
  ctx: AuthorContext;
}

export interface PartAuthoring {
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  /** Every op (equal included), ids prefixed with the part id. */
  ops: DiffOp[];
  /** True when some op is not `equal`. */
  changed: boolean;
  /**
   * The authored blocks (paragraphs and tables) when `changed` and they pass the
   * self-check; `null` when changed and they failed it; `undefined` when
   * nothing changed (the caller keeps the part as it is).
   */
  blocksXml: string | null | undefined;
}

/** The part's XML with pre-existing revisions resolved, as the pipeline compared it. */
export function flattenedPart(xml: string | null, policy: CompareConfig['flattenOriginal']): string {
  return xml === null ? '<w:body/>' : flattenTrackedChanges(xml, policy);
}

/**
 * A view reproduces its document: every paragraph's text (08i's fold-set
 * tolerance on reject-all included) and the paragraph count. Run formatting is
 * not required here: where the body's own author falls back to the text rebuild
 * a paragraph's run formatting goes (ADR-0045 § Accepted limits 5), and a part
 * gets exactly the body's author — refusing it would leave the change unmarked,
 * which is worse than the formatting the body already accepts losing.
 */
function reproduces(v: ViewReport): boolean {
  return textLost(v).length === 0 && v.expectedParagraphs === v.actualParagraphs;
}

/** Analyse and author one part. */
export async function analyseAndAuthorPart(input: PartInput): Promise<PartAuthoring> {
  const { partId, config, author, date, ctx } = input;
  const originalParas = input.originalXml === null
    ? [] : preprocessSide(input.originalXml, partId, config.flattenOriginal, config).paras;
  const revisedParas = input.revisedXml === null
    ? [] : preprocessSide(input.revisedXml, partId, config.flattenRevised, config).paras;
  const { ops: raw } = await analyseRecords({ originalParas, revisedParas, config, author, date });
  const ops = prefixPartOps(raw, partId);
  const changed = ops.some((op) => op.type !== 'equal');
  if (!changed) return { originalParas, revisedParas, ops, changed, blocksXml: undefined };

  const originalFlat = flattenedPart(input.originalXml, config.flattenOriginal);
  const revisedFlat = flattenedPart(input.revisedXml, config.flattenRevised);
  const passes = (blocks: string): boolean => {
    const report = reconstruct(`<w:body>${blocks}</w:body>`, originalFlat, revisedFlat);
    return reproduces(report.accept) && reproduces(report.reject);
  };

  const blocks = authorBlocks(raw, originalParas, revisedParas, ctx);
  return { originalParas, revisedParas, ops, changed, blocksXml: passes(blocks) ? blocks : null };
}

/**
 * Prefix a part's op ids (and the move pairing between them) with the part id,
 * so they can sit beside the body's ops in one list. A part's renumbering
 * cascade is not collapsed — its annotation belongs to the part's own analysis,
 * which the summary does not carry — so `collapsedInto` is dropped: the op is
 * shown, never hidden behind a header the list cannot draw.
 */
export function prefixPartOps(ops: readonly DiffOp[], partId: string): DiffOp[] {
  return ops.map((op) => {
    const { collapsedInto: _dropped, ...rest } = op;
    void _dropped;
    return {
      ...rest,
      id: `${partId}/${op.id}`,
      ...(op.moveTargetId !== undefined ? { moveTargetId: `${partId}/${op.moveTargetId}` } : {}),
    };
  });
}
