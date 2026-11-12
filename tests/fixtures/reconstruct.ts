/**
 * The reconstruction yardstick.
 *
 * The comparison itself is production code — `output/reconstruction.ts`, which
 * the replace author's self-check also runs (one
 * implementation) — and is re-exported here with a count summary that is safe
 * to print for a real document: counts and
 * positions only.
 */
import type { ViewReport } from '@features/compare/output/reconstruction.js';

export {
  reconstruct, isClean, textLost, paragraphViews, compareViews, canonicalRunProperties,
  type ParaView, type ViewReport, type ReconstructionReport,
} from '@features/compare/output/reconstruction.js';

/** A one-line count summary, safe to print for a real document. */
export function summarise(v: ViewReport): string {
  return (
    `paragraphs ${v.actualParagraphs}/${v.expectedParagraphs}, ` +
    `text differs ${v.textDiffers.length} (whitespace only ${v.whitespaceOnly.length}, ` +
    `unrevised only ${v.unrevisedOnly.length}, characters excused ${v.excusedCharacters}), ` +
    `formatting lost ${v.formattingDiffers.length} (formatting-only source change ${v.formattingOnly.length})`
  );
}
