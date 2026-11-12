/**
 * Quarantine residual revision and comment markup after flattening.
 * Spec: spec/features/08-compare.md §5.0, §9.2.
 *
 * After {@link flattenTrackedChanges} resolves <w:ins>/<w:del>/<w:moveFrom>/
 * <w:moveTo>, any residual revision elements (e.g. nested or malformed markup
 * that escaped flattening) plus comment anchors must be stripped so they cannot
 * leak into output and so comments don't perturb text extraction. Quarantined
 * elements are COUNTED (reported upstream — "Original contained N comments,
 * excluded from compare") but never silently merged.
 *
 * Structural transform over the fast-xml-parser tree; the `w:` prefix is
 * preserved. Pure.
 */
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

const PARSER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  suppressEmptyNode: true,
} as const;

const BUILDER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  suppressEmptyNode: true,
  format: false,
} as const;

type PoNode = Record<string, unknown>;

/**
 * Elements quarantined wholesale (the element and its subtree are removed).
 * Wrapper revision elements (<w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>) should
 * already be gone post-flatten; if any survive we drop their content here
 * (conservative: residual revision content is not trusted as body text).
 */
const QUARANTINE_TAGS = new Set([
  'w:ins',
  'w:del',
  'w:moveFrom',
  'w:moveTo',
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
  'w:commentRangeStart',
  'w:commentRangeEnd',
  'w:commentReference',
]);

/** Tags whose removal counts as "a comment was excluded" for reporting. */
const COMMENT_TAGS = new Set(['w:commentReference']);

/**
 * Strip residual revision + comment markup from a (typically already-flattened)
 * part. Returns the clean XML and the count of quarantined revision/comment
 * elements removed.
 */
export function quarantineRevisions(documentXml: string): {
  cleanXml: string;
  quarantinedCount: number;
} {
  const parser = new XMLParser(PARSER_OPTS);
  const tree = parser.parse(documentXml) as PoNode[];

  let quarantinedCount = 0;
  const strip = (nodes: PoNode[]): PoNode[] => {
    const out: PoNode[] = [];
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag !== undefined && QUARANTINE_TAGS.has(tag)) {
        quarantinedCount++;
        continue;
      }
      const value = tag !== undefined ? node[tag] : undefined;
      if (tag !== undefined && Array.isArray(value)) {
        out.push({ ...node, [tag]: strip(value as PoNode[]) });
      } else {
        out.push(node);
      }
    }
    return out;
  };

  const cleaned = strip(tree);

  const builder = new XMLBuilder(BUILDER_OPTS);
  return { cleanXml: builder.build(cleaned) as string, quarantinedCount };
}

/**
 * Count only the comment anchors in a part (without modifying it) — used to
 * report "N comments excluded" separately from revision quarantine.
 */
export function countComments(documentXml: string): number {
  const parser = new XMLParser(PARSER_OPTS);
  const tree = parser.parse(documentXml) as PoNode[];
  let count = 0;
  const walk = (nodes: PoNode[]): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag !== undefined && COMMENT_TAGS.has(tag)) count++;
      const value = tag !== undefined ? node[tag] : undefined;
      if (tag !== undefined && Array.isArray(value)) walk(value as PoNode[]);
    }
  };
  walk(tree);
  return count;
}

/** The element tag of a preserveOrder node (single non-attrs key), or undefined. */
function tagOf(node: PoNode): string | undefined {
  for (const key of Object.keys(node)) {
    if (key === ':@') continue;
    return key;
  }
  return undefined;
}
