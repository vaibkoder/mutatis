/**
 * Extract OOXML part strings from a DocxBundle for the Compare pipeline.
 * Spec: spec/features/08-compare.md §3, §5.0, §9.7/9.8; ADR-0047.
 *
 * A header or footer is found by its ROLE — the section whose `<w:sectPr>`
 * references it and the reference's `w:type` (default / first / even) — through
 * `word/_rels/document.xml.rels`, never by its file name. Word numbers the files
 * as it saves them, so `header1.xml` on one side can be the first-page header
 * and on the other the default one; matched by name they would compare two
 * different things (ADR-0047).
 */
import type { DocxBundle } from '@shared/ooxml/index.js';

/** A part's package path and its XML. */
export interface PartRef {
  /** Package path, e.g. `word/header2.xml`. */
  path: string;
  xml: string;
}

export type HeaderFooterKind = 'header' | 'footer';
export type HeaderFooterType = 'default' | 'first' | 'even';

/** One `w:headerReference` / `w:footerReference` of one section. */
export interface HeaderFooterRole {
  /** Stable role id, e.g. `header:s0:default` — see {@link roleId}. */
  id: string;
  kind: HeaderFooterKind;
  /** 0-based section index: the order of the body's `<w:sectPr>` elements. */
  section: number;
  type: HeaderFooterType;
  /** The relationship id the reference carries in its own package. */
  rId: string;
  /** The part it resolves to through `document.xml.rels`. */
  path: string;
}

/** The set of OOXML parts extracted from one side of a compare operation. */
export interface CompareParts {
  /** word/document.xml — the main body. Always present. */
  documentXml: string;
  /** Header parts by role id (`header:s0:default` …). */
  headers: Map<string, PartRef>;
  /** Footer parts by role id (`footer:s0:default` …). */
  footers: Map<string, PartRef>;
  /** Every header/footer reference of every section, in section order. */
  roles: HeaderFooterRole[];
  /** How many sections the body has (its `<w:sectPr>` elements). */
  sectionCount: number;
  /** The footnotes part (through its relationship) — null when absent. */
  footnotesXml: string | null;
  /** The endnotes part (through its relationship) — null when absent. */
  endnotesXml: string | null;
  /** Package paths of the two notes parts, when present (ADR-0047 writes them back there). */
  notesPaths?: { footnotes: string | null; endnotes: string | null };
  /** word/comments.xml — null when the part is absent. */
  commentsXml: string | null;
  /** word/styles.xml — null when the part is absent. */
  stylesXml: string | null;
  /** word/numbering.xml — null when the part is absent. */
  numberingXml: string | null;
}

const REL_TYPE_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** The role id of a header/footer reference: `<kind>:s<section>:<type>`. */
export function roleId(kind: HeaderFooterKind, section: number, type: HeaderFooterType): string {
  return `${kind}:s${section}:${type}`;
}

/** One relationship of a `.rels` part. */
export interface Relationship {
  id: string;
  /** The type's last segment, e.g. `header`, `footnotes`, `image`. */
  type: string;
  target: string;
  external: boolean;
}

const attr = (tag: string, name: string): string | undefined =>
  new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];

/** The relationships of a `.rels` part (empty when absent). */
export function parseRelationships(relsXml: string | undefined): Relationship[] {
  if (relsXml === undefined) return [];
  const out: Relationship[] = [];
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id');
    const type = attr(m[0], 'Type');
    const target = attr(m[0], 'Target');
    if (id === undefined || type === undefined || target === undefined) continue;
    out.push({
      id,
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      external: attr(m[0], 'TargetMode') === 'External',
    });
  }
  return out;
}

/** Resolve a relationship target against the directory of its source part. */
export function resolveTarget(sourceDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segments = sourceDir === '' ? [] : sourceDir.split('/');
  for (const s of target.split('/')) {
    if (s === '..') segments.pop();
    else if (s !== '.' && s !== '') segments.push(s);
  }
  return segments.join('/');
}

/** Where a section's `<w:sectPr>` sits in `document.xml`. */
export interface SectionSpan {
  start: number;
  end: number;
  xml: string;
}

/**
 * Every section's `<w:sectPr>` in document order — the ones in a paragraph's
 * properties, then the body's last. The OLD properties a `w:sectPrChange`
 * carries are a `w:sectPr` too and are not a section, so they are skipped.
 */
export function sectionSpans(documentXml: string): SectionSpan[] {
  const spans: SectionSpan[] = [];
  const re = /<w:sectPr\b[^>]*?(\/?)>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(documentXml)) !== null) {
    if (m[1] === '/') {
      spans.push({ start: m.index, end: re.lastIndex, xml: m[0] });
      continue;
    }
    // A sectPr cannot contain another except inside its sectPrChange.
    let close = documentXml.indexOf('</w:sectPr>', re.lastIndex);
    const change = documentXml.indexOf('<w:sectPrChange', re.lastIndex);
    if (change !== -1 && change < close) {
      const changeEnd = documentXml.indexOf('</w:sectPrChange>', change);
      if (changeEnd !== -1) close = documentXml.indexOf('</w:sectPr>', changeEnd);
    }
    if (close === -1) break;
    const end = close + '</w:sectPr>'.length;
    spans.push({ start: m.index, end, xml: documentXml.slice(m.index, end) });
    re.lastIndex = end;
  }
  return spans;
}

/** A header/footer reference inside one `<w:sectPr>`. */
export interface SectionReference {
  kind: HeaderFooterKind;
  type: HeaderFooterType;
  rId: string;
  /** The reference element itself. */
  xml: string;
}

/** The header and footer references of one `<w:sectPr>` (its own, not its change's). */
export function sectionReferences(sectPrXml: string): SectionReference[] {
  const own = sectPrXml.replace(/<w:sectPrChange\b[\s\S]*?<\/w:sectPrChange>/g, '');
  const out: SectionReference[] = [];
  for (const m of own.matchAll(/<w:(header|footer)Reference\b[^>]*>/g)) {
    const rId = attr(m[0], 'r:id');
    if (rId === undefined) continue;
    const t = attr(m[0], 'w:type');
    out.push({
      kind: m[1] as HeaderFooterKind,
      type: t === 'first' || t === 'even' ? t : 'default',
      rId,
      xml: m[0],
    });
  }
  return out;
}

/** Extract the relevant OOXML parts from a bundle. */
export function extractParts(bundle: DocxBundle): CompareParts {
  const docXml = bundle.part('word/document.xml');
  if (docXml === undefined) {
    throw new Error('extractParts: word/document.xml missing from bundle');
  }

  const rels = parseRelationships(bundle.part('word/_rels/document.xml.rels'));
  const relById = new Map(rels.map((r) => [r.id, r]));
  const pathOfType = (type: string, fallback: string): string | null => {
    const rel = rels.find((r) => r.type === type && !r.external);
    const path = rel ? resolveTarget('word', rel.target) : fallback;
    return bundle.hasPart(path) ? path : null;
  };
  const footnotesPath = pathOfType('footnotes', 'word/footnotes.xml');
  const endnotesPath = pathOfType('endnotes', 'word/endnotes.xml');

  const roles: HeaderFooterRole[] = [];
  const headers = new Map<string, PartRef>();
  const footers = new Map<string, PartRef>();
  const spans = sectionSpans(docXml);
  spans.forEach((span, section) => {
    for (const ref of sectionReferences(span.xml)) {
      const rel = relById.get(ref.rId);
      if (rel === undefined || rel.external || rel.type !== ref.kind) continue;
      const path = resolveTarget('word', rel.target);
      const xml = bundle.part(path);
      if (xml === undefined) continue;
      const id = roleId(ref.kind, section, ref.type);
      if (headers.has(id) || footers.has(id)) continue; // a repeated reference: the first wins, as in Word
      roles.push({ id, kind: ref.kind, section, type: ref.type, rId: ref.rId, path });
      (ref.kind === 'header' ? headers : footers).set(id, { path, xml });
    }
  });

  return {
    documentXml: docXml,
    headers,
    footers,
    roles,
    sectionCount: spans.length,
    footnotesXml: footnotesPath === null ? null : bundle.part(footnotesPath) ?? null,
    endnotesXml: endnotesPath === null ? null : bundle.part(endnotesPath) ?? null,
    notesPaths: { footnotes: footnotesPath, endnotes: endnotesPath },
    commentsXml: bundle.part('word/comments.xml') ?? null,
    stylesXml: bundle.part('word/styles.xml') ?? null,
    numberingXml: bundle.part('word/numbering.xml') ?? null,
  };
}

/** The relationship type URI for a part kind (`header`, `footnotes` …). */
export function relationshipType(type: string): string {
  return `${REL_TYPE_BASE}${type}`;
}

/**
 * The part a section SHOWS for a role: its own reference, or — Word's "link to
 * previous" — the nearest earlier section's reference of the same kind and type.
 */
export function effectivePart(
  parts: CompareParts,
  kind: HeaderFooterKind,
  section: number,
  type: HeaderFooterType,
): { roleId: string; part: PartRef } | null {
  const map = kind === 'header' ? parts.headers : parts.footers;
  for (let s = section; s >= 0; s--) {
    const id = roleId(kind, s, type);
    const part = map.get(id);
    if (part !== undefined) return { roleId: id, part };
  }
  return null;
}
