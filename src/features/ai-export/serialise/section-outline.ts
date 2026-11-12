/**
 * Section-outline serialiser (feature 11 — spec §4.3).
 *
 * Rebuilds the clause tree from the flat Section Index (`SectionNode.parentKey`)
 * as the navigational skeleton. Clause *bodies* are deliberately not exported in
 * v1 — the outline lets an LLM direct its own reading of the raw text.
 */
import type { SectionIndex, SectionNode } from '@shared/section-index/index.js';
import type { SectionOutlineExport, ClauseNode } from '../types.js';

export interface SectionOutlineDeps {
  /** Heading text of a section paragraph by id. */
  paraText: (paragraphId: string) => string;
}

export function serialiseSectionOutline(
  index: SectionIndex,
  deps: SectionOutlineDeps,
): SectionOutlineExport {
  // Build a ClauseNode per SectionNode, preserving document order.
  const byKey = new Map<string, ClauseNode>();
  const made: Array<{ node: SectionNode; clause: ClauseNode }> = [];

  for (const node of index.nodes) {
    const clause: ClauseNode = {
      number: node.displayLabel,
      heading: deps.paraText(node.paraId).trim(),
      level: node.level + 1, // SectionNode.level is 0-based; spec wants 1-based
      kind: node.kind,
    };
    made.push({ node, clause });
    if (!byKey.has(node.key)) byKey.set(node.key, clause);
  }

  const roots: ClauseNode[] = [];
  for (const { node, clause } of made) {
    const parent = node.parentKey ? byKey.get(node.parentKey) : undefined;
    if (parent) {
      (parent.children ??= []).push(clause);
    } else {
      roots.push(clause);
    }
  }

  return { clauses: roots };
}
