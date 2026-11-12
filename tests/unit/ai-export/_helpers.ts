/**
 * Shared mock factories for the AI-export unit tests. Not a test file (no
 * `.test.` suffix, so Vitest does not collect it).
 */
import type { SectionIndex, SectionNode, SectionKey } from '@shared/section-index/index.js';

export function makeNode(p: Partial<SectionNode> & { key: string; displayLabel: string; paraId: string }): SectionNode {
  return {
    kind: 'clause',
    rawNumber: p.displayLabel.replace(/^\D+/, ''),
    level: 0,
    isStaticNumbered: false,
    bookmarkNames: [],
    getRange: () => { throw new Error('not used in tests'); },
    ...p,
  } as SectionNode;
}

export function makeIndex(nodes: SectionNode[]): SectionIndex {
  const byKey = new Map<SectionKey, SectionNode[]>();
  const byParaId = new Map<string, SectionNode>();
  for (const n of nodes) {
    (byKey.get(n.key) ?? byKey.set(n.key, []).get(n.key)!).push(n);
    byParaId.set(n.paraId, n);
  }
  return {
    nodes,
    byKey,
    byParaId,
    ordered: nodes,
    has: (k) => byKey.has(k),
    resolve: (k) => byKey.get(k) ?? [],
  };
}

/** A paraText resolver backed by a plain id→text map. */
export function paraTextFrom(map: Record<string, string>): (id: string) => string {
  return (id) => map[id] ?? '';
}
