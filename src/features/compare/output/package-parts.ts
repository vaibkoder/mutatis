/**
 * Feature 08 — Compare: where each header and footer goes in the redline
 * package, and the revised document's section references pointed there
 * (ADR-0047).
 *
 * The redline is a clone of the ORIGINAL package, but most of its body — every
 * unchanged paragraph, a section break's properties among them — is written
 * from the REVISED document, whose `w:headerReference`/`w:footerReference`
 * carry relationship ids of the REVISED package. In the original's relationships
 * the same id can be another header, a hyperlink or nothing. So before the
 * pipeline runs, every header/footer reference of the revised `document.xml` is
 * rewritten, per section and type, to the relationship id its role's part has
 * in the output package:
 *
 *  - a role the original references itself is written back at the original's
 *    path, under the original's relationship id (the original's own references,
 *    in struck paragraphs, stay valid);
 *  - a role only the revised document references gets a NEW part and a new
 *    relationship id (`rIdDpN`), so the authored part — the revised paragraphs
 *    inserted, or diffed against the part the original inherits there — has
 *    somewhere to live;
 *  - a role only the original references is added to the revised section's
 *    `<w:sectPr>`, so the redline still shows its part (its paragraphs struck);
 *  - with headers (or footers) excluded from comparison, a reference is pointed
 *    at the original's part of the same role, or dropped when the original has
 *    none — never left pointing into the wrong package.
 *
 * Section properties themselves are not compared (out of scope).
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { CompareParts, HeaderFooterKind, HeaderFooterType } from '../io/index.js';
import { effectivePart, roleId, sectionSpans, sectionReferences } from '../io/parts.js';
import { pairRoles, type RolePair } from '../parts-diff.js';

/** One authored header/footer and where it goes. */
export interface PlannedPart {
  pair: RolePair;
  /** Output package path. */
  path: string;
  /** Its relationship id in the output `document.xml.rels`. */
  rId: string;
  /** The original package has no part at `path` for this role: add it. */
  isNew: boolean;
}

export interface HeaderFooterPlan {
  /** The revised `document.xml`, every header/footer reference pointing into the output package. */
  revisedDocumentXml: string;
  /** Every header/footer to author, in role order. */
  parts: PlannedPart[];
}

const TYPES: readonly HeaderFooterType[] = ['default', 'first', 'even'];

/**
 * Plan the output headers and footers and remap the revised document's
 * references. `originalPaths` are the original package's part paths (a new part
 * takes a name none of them has).
 */
export function planHeaderFooters(
  original: CompareParts,
  revised: CompareParts,
  config: Pick<CompareConfig, 'excludeHeaders' | 'excludeFooters'>,
  originalPaths: readonly string[],
  newRelationshipId: () => string,
): HeaderFooterPlan {
  const enabled = (kind: HeaderFooterKind): boolean =>
    kind === 'header' ? !config.excludeHeaders : !config.excludeFooters;
  const taken = new Set(originalPaths);
  const claimed = new Set<string>();
  const parts: PlannedPart[] = [];
  /** role id → output relationship id, for every role either side references. */
  const output = new Map<string, string>();
  const originalRole = new Map(original.roles.map((r) => [r.id, r]));

  for (const kind of ['header', 'footer'] as const) {
    if (!enabled(kind)) continue;
    const byParts = new Map<string, PlannedPart>();
    for (const pair of pairRoles(original, revised, kind)) {
      const own = originalRole.get(pair.roleId);
      let planned: PlannedPart;
      if (own !== undefined && !claimed.has(own.path)) {
        planned = { pair, path: own.path, rId: own.rId, isNew: false };
      } else {
        let n = 1;
        while (taken.has(`word/${kind}${n}.xml`)) n++;
        planned = { pair, path: `word/${kind}${n}.xml`, rId: newRelationshipId(), isNew: true };
      }
      taken.add(planned.path);
      claimed.add(planned.path);
      parts.push(planned);
      byParts.set(partsKey(pair), planned);
    }
    // Every role, including one `pairRoles` folded into an earlier pair that
    // resolves to the same two parts, points at its pair's output part.
    for (const role of [...original.roles, ...revised.roles]) {
      if (role.kind !== kind || output.has(role.id)) continue;
      const shown = (side: CompareParts): string => effectivePart(side, kind, role.section, role.type)?.part.path ?? '';
      const planned = byParts.get(`${shown(original)}|${shown(revised)}`);
      if (planned !== undefined) output.set(role.id, planned.rId);
    }
  }

  return { revisedDocumentXml: remapReferences(original, revised, enabled, output), parts };
}

const partsKey = (pair: RolePair): string => `${pair.original?.path ?? ''}|${pair.revised?.path ?? ''}`;

/** Rewrite (and add) the revised document's header/footer references, section by section. */
function remapReferences(
  original: CompareParts,
  revised: CompareParts,
  enabled: (kind: HeaderFooterKind) => boolean,
  output: ReadonlyMap<string, string>,
): string {
  const xml = revised.documentXml;
  const originalRole = new Map(original.roles.map((r) => [r.id, r]));
  /** The output relationship id for a role, or null to drop the reference. */
  const target = (kind: HeaderFooterKind, section: number, type: HeaderFooterType): string | null => {
    const id = roleId(kind, section, type);
    return enabled(kind) ? output.get(id) ?? null : originalRole.get(id)?.rId ?? null;
  };

  let out = '';
  let at = 0;
  sectionSpans(xml).forEach((span, section) => {
    // Only the section's own references: a sectPrChange's old properties carry none.
    const change = /<w:sectPrChange\b[\s\S]*?<\/w:sectPrChange>/.exec(span.xml)?.[0] ?? '';
    let own = change === '' ? span.xml : span.xml.replace(change, '\u0000');
    const present = new Set<string>();
    own = own.replace(/<w:(header|footer)Reference\b[^>]*>/g, (tag) => {
      const ref = sectionReferences(`<w:sectPr>${tag}</w:sectPr>`)[0];
      if (ref === undefined) return tag;
      present.add(`${ref.kind}:${ref.type}`);
      const rId = target(ref.kind, section, ref.type);
      return rId === null ? '' : tag.replace(/(\sr:id=")[^"]*(")/, `$1${rId}$2`);
    });
    // A role only the original references: keep it shown in the redline.
    const added: string[] = [];
    for (const kind of ['header', 'footer'] as const) {
      for (const type of TYPES) {
        if (present.has(`${kind}:${type}`) || !originalRole.has(roleId(kind, section, type))) continue;
        const rId = target(kind, section, type);
        if (rId !== null) added.push(`<w:${kind}Reference w:type="${type}" r:id="${rId}"/>`);
      }
    }
    if (added.length > 0) {
      own = own.endsWith('/>') && !own.includes('</w:sectPr>')
        ? `${own.slice(0, -2)}>${added.join('')}</w:sectPr>`
        : own.replace(/^<w:sectPr\b[^>]*>/, (open) => `${open}${added.join('')}`);
    }
    if (change !== '') own = own.replace('\u0000', change);
    out += xml.slice(at, span.start) + own;
    at = span.end;
  });
  return out + xml.slice(at);
}
