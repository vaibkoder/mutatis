/**
 * Flatten pre-existing tracked changes in a document part before diffing.
 * Spec: spec/features/08-compare.md §5.0.
 *
 * A document may already contain <w:ins> / <w:del> / <w:moveFrom> / <w:moveTo>
 * revision markup. Diffing raw markup-laden XML produces garbage, so we first
 * resolve those revisions to plain run content:
 *
 *  - acceptAll (default): keep <w:ins> content; drop <w:del> & <w:moveFrom>
 *    content — the "as-if-accepted" text.
 *  - rejectAll: inverse — keep <w:del> content (delText→t) & <w:moveFrom>;
 *    drop <w:ins> & <w:moveTo>.
 *
 * In every case the <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo> WRAPPER elements are
 * unwrapped (their surviving children are spliced into the parent run list), and
 * the standalone move-range markers (<w:moveToRangeStart/End>,
 * <w:moveFromRangeStart/End>) are stripped.
 *
 * A revision on a PARAGRAPH MARK (`<w:pPr><w:rPr><w:del/>`) is a paragraph
 * join: the policy that removes the mark joins the paragraph to the one after
 * it (see {@link markRemovedBy}). The joined paragraph keeps the NEXT
 * paragraph's properties — the surviving mark is the one that carries them. That
 * is the rule this module applies; how Word desktop resolves it is unmeasured
 * until the live-host smoke test (spec 08 §5.0).
 *
 * Implemented as a structural transform over the fast-xml-parser tree so the
 * `w:` namespace prefix is preserved. Pure — no side effects.
 */
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

import type { FlattenPolicy } from '@shared/config/index.js';

const PARSER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  // Keep whitespace-only text nodes — significant inside <w:t xml:space="preserve">.
  trimValues: false,
  // Parse no values as numbers/booleans; OOXML is all strings.
  parseTagValue: false,
  parseAttributeValue: false,
  // Self-closing tags (e.g. <w:moveToRangeStart .../>) round-trip cleanly.
  suppressEmptyNode: true,
} as const;

const BUILDER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  suppressEmptyNode: true,
  format: false,
} as const;

/** preserveOrder node: a single-key object whose key is the tag name. */
type PoNode = Record<string, unknown>;

/** Revision wrappers that resolve to "keep children" or "drop children" per policy. */
const INS_LIKE = new Set(['w:ins', 'w:moveTo']);
const DEL_LIKE = new Set(['w:del', 'w:moveFrom']);

/** Standalone range markers stripped unconditionally during flatten. */
const MOVE_RANGE_MARKERS = new Set([
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
]);

/**
 * Flatten tracked-change markup in a part's XML according to `policy`.
 * Returns the flattened XML string.
 */
export function flattenTrackedChanges(documentXml: string, policy: FlattenPolicy): string {
  const parser = new XMLParser(PARSER_OPTS);
  const tree = parser.parse(documentXml) as PoNode[];

  const flattened = flattenNodes(tree, policy);

  const builder = new XMLBuilder(BUILDER_OPTS);
  return builder.build(flattened) as string;
}

/**
 * Table revision marks (ADR-0046). A row or cell is itself inserted or deleted
 * by a mark in its properties — `<w:trPr><w:del/>`, `<w:tcPr><w:cellIns/>` —
 * not by wrapping it, so the wrapper rules above cannot reach it: they only
 * drop the mark and its runs' text, and accepting a deleted row would leave an
 * empty row behind. The row or cell that does not survive the policy is
 * dropped whole; the cell marks themselves are always stripped.
 */
const CELL_MARKS = new Set(['w:cellIns', 'w:cellDel', 'w:cellMerge']);

/** True when `node` (a `w:tr` or `w:tc`) is removed by `policy`. */
function droppedByPolicy(node: PoNode, tag: string, policy: FlattenPolicy): boolean {
  const propsTag = tag === 'w:tr' ? 'w:trPr' : 'w:tcPr';
  const [insMark, delMark] = tag === 'w:tr' ? ['w:ins', 'w:del'] : ['w:cellIns', 'w:cellDel'];
  const children = node[tag];
  if (!Array.isArray(children)) return false;
  for (const child of children as PoNode[]) {
    if (tagOf(child) !== propsTag) continue;
    const props = child[propsTag];
    if (!Array.isArray(props)) return false;
    const marks = new Set((props as PoNode[]).map(tagOf));
    return policy === 'acceptAll' ? marks.has(delMark) : marks.has(insMark);
  }
  return false;
}

/** True when a `w:tbl` node has at least one `w:tr` child. */
function hasRows(node: PoNode, tag: string): boolean {
  const children = node[tag];
  return Array.isArray(children) && (children as PoNode[]).some((child) => tagOf(child) === 'w:tr');
}

/**
 * Revisions a paragraph MARK carries (`<w:pPr><w:rPr>…</w:rPr></w:pPr>`, the
 * direct `rPr` — never the one inside a `w:pPrChange`), by the policy that
 * removes the mark: accepting removes a deleted (or moved-away) mark,
 * rejecting an inserted (or moved-here) one.
 */
const MARK_REMOVED_BY: Record<FlattenPolicy, ReadonlySet<string>> = {
  acceptAll: new Set(['w:del', 'w:moveFrom']),
  rejectAll: new Set(['w:ins', 'w:moveTo']),
};

/**
 * Does `policy` remove this paragraph's MARK (ADR-0032, ADR-0035)?
 *
 * A paragraph mark carrying `<w:del>` is a paragraph break that was deleted:
 * accepting it JOINS the paragraph to the next one, which is how Word records
 * two clauses merged into one and how ADR-0035 removes a whole deleted
 * paragraph. Read without this, a merge accepted read as two paragraphs and a
 * deleted clause left a blank one behind. `<w:ins>` on the mark is the mirror,
 * resolved by rejecting.
 */
function markRemovedBy(node: PoNode, policy: FlattenPolicy): boolean {
  const pPr = childrenArray(node, 'w:p').find((c) => tagOf(c) === 'w:pPr');
  const rPr = pPr === undefined ? undefined : childrenArray(pPr, 'w:pPr').find((c) => tagOf(c) === 'w:rPr');
  if (rPr === undefined) return false;
  return childrenArray(rPr, 'w:rPr').some((c) => MARK_REMOVED_BY[policy].has(tagOf(c) ?? ''));
}

/**
 * Zero-width markup that may sit between two paragraphs without stopping the
 * first joining the second (it stays where it is, before the joined paragraph).
 */
const BETWEEN_PARAGRAPHS = new Set([
  'w:bookmarkStart', 'w:bookmarkEnd', 'w:commentRangeStart', 'w:commentRangeEnd',
  'w:permStart', 'w:permEnd', 'w:proofErr', ...MOVE_RANGE_MARKERS,
]);

/**
 * Is there a paragraph to join `nodes[i]` to — the next block in the same
 * container, a `w:p`? The last paragraph of a body or a table cell, or one
 * followed by a table, has none: Word keeps it (a body and a cell must end with
 * a paragraph), so it stays, its mark revision resolved.
 */
function joinsNext(nodes: readonly PoNode[], i: number): boolean {
  for (let k = i + 1; k < nodes.length; k++) {
    const tag = tagOf(nodes[k]!);
    if (tag === undefined || BETWEEN_PARAGRAPHS.has(tag)) continue;
    return tag === 'w:p';
  }
  return false;
}

/** The children array of an element node, or [] for a leaf. */
function childrenArray(node: PoNode, tag: string): PoNode[] {
  const value = node[tag];
  return Array.isArray(value) ? (value as PoNode[]) : [];
}

/**
 * `para` with `runs` put first, after its `w:pPr`: the paragraph a removed
 * mark joins into keeps its own properties and attributes (its mark survives;
 * unmeasured in Word, see the module doc).
 */
function prependRuns(para: PoNode, runs: readonly PoNode[]): PoNode {
  const children = childrenArray(para, 'w:p');
  const at = children.findIndex((c) => tagOf(c) === 'w:pPr') + 1;
  return { ...para, 'w:p': [...children.slice(0, at), ...runs, ...children.slice(at)] };
}

/** Recursively flatten an ordered list of preserveOrder nodes. */
function flattenNodes(nodes: PoNode[], policy: FlattenPolicy): PoNode[] {
  const out: PoNode[] = [];
  /** Content of paragraphs whose mark this policy removed, waiting for the next paragraph. */
  let joining: PoNode[] = [];

  for (const [i, node] of nodes.entries()) {
    const tag = tagOf(node);

    if (tag === 'w:p') {
      let kept = recurseInto(node, policy);
      if (joining.length > 0) {
        kept = prependRuns(kept, joining);
        joining = [];
      }
      if (markRemovedBy(node, policy) && joinsNext(nodes, i)) {
        joining = childrenArray(kept, 'w:p').filter((c) => tagOf(c) !== 'w:pPr');
        continue;
      }
      out.push(kept);
      continue;
    }

    // Strip standalone move-range markers entirely.
    if (tag !== undefined && MOVE_RANGE_MARKERS.has(tag)) {
      continue;
    }

    // A row or cell that this policy removes goes whole; cell marks go always.
    if (tag !== undefined && CELL_MARKS.has(tag)) continue;
    if ((tag === 'w:tr' || tag === 'w:tc') && droppedByPolicy(node, tag, policy)) continue;

    // <w:ins>/<w:moveTo>: keep children iff accepting; else drop the whole subtree.
    if (tag !== undefined && INS_LIKE.has(tag)) {
      if (policy === 'acceptAll') {
        out.push(...flattenChildrenOf(node, tag, policy));
      }
      continue;
    }

    // <w:del>/<w:moveFrom>: keep children iff rejecting; else drop. When kept,
    // delText must be converted back to ordinary text (handled in convert step).
    if (tag !== undefined && DEL_LIKE.has(tag)) {
      if (policy === 'rejectAll') {
        const kept = flattenChildrenOf(node, tag, policy);
        out.push(...kept.map(delToText));
      }
      continue;
    }

    // Ordinary node: recurse into its children (if any) and keep it in place —
    // except a table whose every row this policy removed, which goes whole, as
    // Word removes it: a table inserted with tracked changes, rejected, leaves
    // no empty table shell behind.
    const kept = recurseInto(node, policy);
    if (tag === 'w:tbl' && hasRows(node, tag) && !hasRows(kept, tag)) continue;
    out.push(kept);
  }

  return out;
}

/** Flatten the children array of a wrapper node and return the resulting list. */
function flattenChildrenOf(node: PoNode, tag: string, policy: FlattenPolicy): PoNode[] {
  const children = node[tag];
  if (!Array.isArray(children)) return [];
  return flattenNodes(children as PoNode[], policy);
}

/** Recurse into a non-revision node's children, rebuilding it with attributes intact. */
function recurseInto(node: PoNode, policy: FlattenPolicy): PoNode {
  const tag = tagOf(node);
  if (tag === undefined) return node; // text node (#text) or attrs container — leave as-is.

  const value = node[tag];
  if (!Array.isArray(value)) return node; // leaf / self-closing — nothing to recurse.

  const rebuilt: PoNode = { ...node, [tag]: flattenNodes(value as PoNode[], policy) };
  return rebuilt;
}

/**
 * When a deleted run survives (rejectAll), convert <w:delText> → <w:t> and
 * <w:delInstrText> → <w:instrText> so the text re-enters the normal run stream.
 * Applied recursively to a kept node tree.
 */
function delToText(node: PoNode): PoNode {
  const tag = tagOf(node);
  if (tag === undefined) return node;

  let newTag = tag;
  if (tag === 'w:delText') newTag = 'w:t';
  else if (tag === 'w:delInstrText') newTag = 'w:instrText';

  const value = node[tag];
  if (!Array.isArray(value)) {
    if (newTag === tag) return node;
    // Rename leaf tag, preserving its attrs container if present.
    const { [tag]: v, ...rest } = node;
    return { ...rest, [newTag]: v };
  }

  const recurChildren = (value as PoNode[]).map(delToText);
  if (newTag === tag) {
    return { ...node, [tag]: recurChildren };
  }
  const { [tag]: _omit, ...rest } = node;
  void _omit;
  return { ...rest, [newTag]: recurChildren };
}

/** The element tag of a preserveOrder node (the single non-attrs key), or undefined. */
function tagOf(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@') continue; // attributes container in preserveOrder mode
    return key;
  }
  return undefined;
}
