/**
 * Pre-processing pipeline for Compare (Wave 1-A, task A1).
 * Spec: spec/features/08-compare.md §5.0.
 *
 * Barrel + pipeline helper: flatten existing tracked changes → quarantine
 * residual revision/comment markup → extract ParaRecord[]. Pure throughout.
 */
import type { CompareConfig, FlattenPolicy } from '@shared/config/index.js';

import type { ParaRecord } from '../model.js';

import { extractParaRecords } from './pararecord.js';
import { quarantineRevisions } from './quarantine.js';
import { flattenTrackedChanges } from './flatten.js';
import { type NumberingContext, assignNumberLabels } from './numbering.js';
import { groupTableUnits } from './table-units.js';

export { fnv1a } from './hash.js';
export {
  normaliseText,
  normaliseTextAggressive,
  stripNumberingPrefix,
} from './normalise.js';
export { flattenTrackedChanges } from './flatten.js';
export { quarantineRevisions, countComments } from './quarantine.js';
export { extractParaRecords } from './pararecord.js';
export { buildNumberingContext, assignNumberLabels } from './numbering.js';
export { groupTableUnits, flattenTableUnits } from './table-units.js';
export type { NumberingContext } from './numbering.js';

/** Result of pre-processing one side (one part). */
export interface PreprocessResult {
  paras: ParaRecord[];
  /** Revision/comment elements quarantined after flatten. */
  quarantinedCount: number;
}

/**
 * Pre-process one document part end-to-end:
 *   flatten(policy) → quarantine → extractParaRecords.
 *
 * @param documentXml  raw part XML (may carry pre-existing tracked changes).
 * @param partId       'document' | 'header1' | 'footer1' | 'footnotes' | …
 * @param policy       flatten policy for this side (acceptAll | rejectAll).
 * @param config       compare config.
 * @param numbering    THIS side's styles + numbering scheme (ADR-0042). Optional:
 *                     a caller with no `styles.xml`/`numbering.xml` to give (the
 *                     header/footer part diff, the text and PDF paths) leaves
 *                     every record un-numbered, which is what it was before.
 */
export function preprocessSide(
  documentXml: string,
  partId: string,
  policy: FlattenPolicy,
  config: CompareConfig,
  numbering?: NumberingContext,
): PreprocessResult {
  const flattened = flattenTrackedChanges(documentXml, policy);
  const { cleanXml, quarantinedCount } = quarantineRevisions(flattened);
  const flat = extractParaRecords(cleanXml, partId, config, numbering);
  // The rendered label needs the whole part in document order — the counter
  // state machine is a sequence, not a per-paragraph function.
  if (numbering) assignNumberLabels(flat, numbering.scheme);
  // ADR-0046 — tables become units AFTER numbering, so every counter and label
  // sees the cell paragraphs exactly as before. Every part, not only the body:
  // a header, footer or note is authored by the body's author (ADR-0047), and a
  // table read as loose cell paragraphs would be written back as loose
  // paragraphs — its structure gone.
  const paras = groupTableUnits(flat, cleanXml);
  return { paras, quarantinedCount };
}
