/**
 * Compare pre-process — the numbering a paragraph does not state.
 * Spec: spec/features/08-compare.md §9.4 · ADR-0042.
 *
 * A clause's number is `(numId, ilvl)` resolved through the paragraph style
 * chain and then counted in document order. Neither half was available to
 * Compare: `ParaRecord.numberingPrefix` is scraped from literal run text, so it
 * is `''` for every automatically numbered paragraph, which is all of them in a
 * firm template. Two paragraphs at different list levels therefore compared
 * EQUAL and a re-levelled clause produced no op at all.
 *
 * This module supplies the context that fixes that, and nothing else: it
 * resolves each paragraph's numbering IDENTITY (`@shared/numbering`'s
 * `resolveParaNumPr`) and renders each paragraph's LABEL with the counter state
 * machine that already exists (`renderListNumbers`). Both are shared code —
 * there is no second resolver and no second counter here.
 *
 * `hash` is NOT touched. Identity and label ride alongside the text as
 * comparable properties; folding them into the hash would break the alignment
 * the whole readability plan exists to protect (ADR-0042).
 */
import {
  type NumberingScheme,
  type StyleNumberingSource,
  type ParaNumPr,
  parseNumberingSchemeXml,
  parseStyleNumberingSource,
  renderListNumbers,
} from '@shared/numbering/index.js';

import type { ParaRecord } from '../model.js';

/** Everything needed to resolve and render one side's paragraph numbering. */
export interface NumberingContext {
  styles: StyleNumberingSource;
  scheme: NumberingScheme;
}

/**
 * Build the context for one side from its raw parts. Either part may be absent —
 * and what that yields is NOT "every paragraph un-numbered", which this comment
 * claimed. Measured on a paragraph carrying a direct
 * `<w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr>`:
 *
 *     no context at all   -> numbering: null
 *     context(null, null) -> { numId: '3', ilvl: 1, key: '1\u0001',
 *                              label: null, path: null, renders: null }
 *
 * A direct `<w:numPr>` needs neither part to be READ, so the identity resolves
 * and only the rendered label does not. That is the better behaviour — a
 * re-levelling is still detected in a package whose `numbering.xml` is missing —
 * and it is not what the comment said, which matters because this is the one
 * place that tells a caller what the absent path does.
 *
 * Passing NO context is the different thing, and is what the Text and PDF paths
 * do: `preprocessSide` then records no numbering at all. Every call site builds
 * contexts for the two sides the same way, so neither shape can produce a
 * one-sided difference.
 */
export function buildNumberingContext(
  stylesXml: string | null | undefined,
  numberingXml: string | null | undefined,
): NumberingContext {
  return {
    styles: parseStyleNumberingSource(stylesXml),
    scheme: parseNumberingSchemeXml(numberingXml),
  };
}

/**
 * Assign `numberLabel` to each record that carries a numbering identity, by
 * running the shared counter state machine over the whole part in document
 * order. MUTATES the records — they were built moments ago by
 * `extractParaRecords` and are not yet observable elsewhere; doing it as a
 * second pass is what lets the counter see the sequence, which is the one thing
 * a per-paragraph resolver structurally cannot.
 *
 * Counters are per part (`document`, each header, each footer), which is also
 * how Word treats them: a footer's list does not continue the body's.
 */
export function assignNumberLabels(records: ParaRecord[], scheme: NumberingScheme): void {
  const refs: ParaNumPr[] = [];
  for (const rec of records) {
    if (rec.numbering) {
      refs.push({ paraId: String(rec.index), numId: rec.numbering.numId, ilvl: rec.numbering.ilvl });
    }
  }
  if (refs.length === 0) return;

  const labels = renderListNumbers(scheme, refs);
  for (const rec of records) {
    if (!rec.numbering) continue;
    // Absent from the map = the (numId, ilvl) resolved to no level definition.
    // Leave the label null rather than inventing one; the IDENTITY is still
    // comparable, so a change is still visible.
    const rendered = labels.get(String(rec.index));
    rec.numbering.label = rendered?.listString ?? null;
    rec.numbering.path = rendered?.path ?? null;
  }
}
