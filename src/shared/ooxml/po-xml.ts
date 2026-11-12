/**
 * host-web — shared `fast-xml-parser` (`preserveOrder`) helpers.
 *
 * The single home for the preserveOrder walk idiom used across the web OOXML
 * readers (`ooxml-loader.ts`, `numbering-scheme.ts`). Extracted so the loader and
 * the numbering parser share ONE parser configuration + node-shape helpers rather
 * than each re-inventing `tagOf`/`childrenOf`/`attrOf` (reuse before build).
 *
 * Host-NEUTRAL by construction: no `Office.*`/`Word.*` token — `host-web` stays off
 * the `no-host-imports` allow-list.
 *
 * preserveOrder node shape (fast-xml-parser): each element is an object with a
 * single tag key whose value is the child-node array, plus an optional `:@`
 * attribute bag. Text leaves are `{ '#text': string }`. Example for
 * `<w:lvl w:ilvl="0"><w:start w:val="1"/></w:lvl>`:
 *   { 'w:lvl': [ { 'w:start': [], ':@': { '@_w:val': '1' } } ], ':@': { '@_w:ilvl': '0' } }
 */
import { XMLParser } from 'fast-xml-parser';

/** A preserveOrder node: one tag key (+ child array) plus an optional `:@` bag. */
export type PoNode = Record<string, unknown>;

/** The shared parser options. `preserveOrder` keeps document order + duplicate tags. */
export const PRESERVE_ORDER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  suppressEmptyNode: true,
} as const;

/** Parse an OOXML part string into a preserveOrder node array. */
export function parsePreserveOrder(xml: string): PoNode[] {
  return new XMLParser(PRESERVE_ORDER_OPTS).parse(xml) as PoNode[];
}

/** The element tag of a preserveOrder node (its single non-attrs key), or undefined. */
export function tagOf(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@') continue;
    return key;
  }
  return undefined;
}

/** The child-node array of a preserveOrder element, or []. */
export function childrenOf(node: PoNode): PoNode[] {
  const tag = tagOf(node);
  if (tag === undefined) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as PoNode[]) : [];
}

/** Read one attribute (`@_w:val` etc.) off a preserveOrder node's `:@` bag. */
export function attrOf(node: PoNode, name: string): string | undefined {
  const attrs = node[':@'] as Record<string, unknown> | undefined;
  const v = attrs?.[name];
  return typeof v === 'string' ? v : undefined;
}
