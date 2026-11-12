/**
 * Human citation resolver (feature 11 — rule F3: cite in human terms, never paraId).
 *
 * The checkers anchor their records to a paragraph id. The Section Index only
 * indexes *heading* paragraphs, so to cite the clause that *contains* an
 * arbitrary paragraph we resolve the nearest section heading at-or-before the
 * target paragraph in document order, and return its `displayLabel`
 * (e.g. "Clause 4.2", "Schedule 1"). Pure in-memory lookup — no scanning.
 */
import type { SectionIndex } from '@shared/section-index/index.js';

export type Citer = (paragraphId: string) => string;

/**
 * Build a citation function.
 *
 * @param index         the Section Index, or `null` when Cross-Reference has not
 *                      run (citations degrade to "" and blind_spots notes it).
 * @param orderedParaIds all paragraph ids in document order (from the Doc Model).
 */
export function makeCiter(index: SectionIndex | null, orderedParaIds: readonly string[]): Citer {
  if (!index) return () => '';

  const docPos = new Map<string, number>();
  orderedParaIds.forEach((id, i) => {
    if (!docPos.has(id)) docPos.set(id, i);
  });

  // Section headings sorted by their document position.
  const headings = index.nodes
    .map(n => ({ pos: docPos.get(n.paraId) ?? -1, label: n.displayLabel, paraId: n.paraId }))
    .filter(h => h.pos >= 0)
    .sort((a, b) => a.pos - b.pos);

  return (paragraphId: string): string => {
    // Exact heading match first.
    const exact = index.byParaId.get(paragraphId);
    if (exact) return exact.displayLabel;

    const pos = docPos.get(paragraphId);
    if (pos === undefined) return '';

    // Nearest heading at-or-before this paragraph (linear scan from the end).
    let label = '';
    for (const h of headings) {
      if (h.pos <= pos) label = h.label;
      else break;
    }
    return label;
  };
}
