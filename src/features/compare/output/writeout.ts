/**
 * Feature 08 — Compare: materialise the authored redline (C2).
 *
 * Assembles the authored `<w:body>` into a clone of the original package and
 * delegates materialisation to the active HostPort (`getActiveHost().writeBack`).
 * No host bridge here — the office adapter (OOXML Service) owns Office JS; the
 * web adapter serialises to a `.docx` Blob download (key constraint 2; ADR-0022).
 *
 * Default output is a standalone new document (office: `createDocument().open()`),
 * because an in-place insert can drop package-level constructs (§9.13). The
 * insert-into-active path is the fallback and validates the round-trip via the
 * service's artifact detection (office only — a no-op distinction on web).
 *
 * Spec: spec/features/08-compare.md §5.5, §9.13.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import { getActiveHost } from '@shared/host/index.js';

import { sectionSpans } from '../io/parts.js';

export interface WriteoutOptions {
  /** Default 'newDocument'. Insert-into-active is the fallback. */
  mode: 'newDocument' | 'insertIntoActive';
  /** Required when mode === 'insertIntoActive'. */
  onArtifactDetected?: 'fallbackToNewDoc' | 'throw';
  author: string;
  date: string;
}

/**
 * Replace the `<w:body>…</w:body>` subtree of a full `word/document.xml` string
 * with `authoredBodyXml`, preserving the original `<w:document>` root (and its
 * namespace declarations) and anything after the body (e.g. trailing
 * `<w:sectPr>` already inside the authored body, or document-level elements).
 *
 * Pure string surgery to preserve byte fidelity of the root declarations
 * (key constraint 6).
 */
export function replaceBody(originalDocumentXml: string, authoredBodyXml: string): string {
  const bodyOpen = originalDocumentXml.search(/<w:body(?:\s[^>]*)?>/);
  const bodyCloseIdx = originalDocumentXml.lastIndexOf('</w:body>');

  if (bodyOpen === -1 || bodyCloseIdx === -1) {
    throw new Error('replaceBody: <w:body> not found in original document.xml');
  }

  const before = originalDocumentXml.slice(0, bodyOpen);
  const after = originalDocumentXml.slice(bodyCloseIdx + '</w:body>'.length);

  // authoredBodyXml already includes its own <w:body>…</w:body> wrapper.
  return `${before}${authoredBodyXml}${after}`;
}

/** A part the original package does not have (ADR-0047). */
export interface NewPart {
  /** Package path, e.g. `word/header4.xml`. */
  path: string;
  xml: string;
  /** Relationship type's last segment (`header`, `footer`, `footnotes`, `endnotes`). */
  relType: string;
  /** The relationship id `document.xml.rels` gives it — allocated by the caller when the body must reference it. */
  rId: string;
  /** The part's own relationships (a header's logo), copied from the package it came from. */
  rels?: PartRels;
}

/** A copied part's own `.rels` and the internal files it targets. */
export interface PartRels {
  /** The `.rels` XML as the source package has it. */
  xml: string;
  files: Array<{
    /** The `Target` attribute as written in `xml`. */
    target: string;
    /** The file's path in the source package. */
    path: string;
    bytes: Uint8Array;
    /** The source package's `Default` content type for its extension, if any. */
    contentType?: string;
  }>;
}

/**
 * What the redline package carries beyond the authored body (ADR-0047): the
 * authored headers, footers and notes, and the body's last `<w:sectPr>`.
 */
export interface AuxiliaryParts {
  /** Existing parts replaced: package path → authored XML. */
  replace?: ReadonlyMap<string, string>;
  /** Parts the original lacks, wired into `document.xml.rels` and `[Content_Types].xml`. */
  add?: readonly NewPart[];
  /**
   * The body's last `<w:sectPr>` — page setup and the last section's header and
   * footer references. The authored body has none (`authorBody` writes
   * paragraphs and tables), so without one Word falls back to its defaults and
   * shows no header or footer in the last section. Absent ⇒ the original's.
   */
  trailingSectPr?: string;
}

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const DOC_RELS = 'word/_rels/document.xml.rels';

/**
 * The body-level `<w:sectPr>` of a `document.xml` — the last section's — or
 * null when the body ends without one.
 */
export function trailingSectPr(documentXml: string): string | null {
  const last = sectionSpans(documentXml).at(-1);
  if (last === undefined) return null;
  return /^\s*<\/w:body>/.test(documentXml.slice(last.end)) ? last.xml : null;
}

/** Relationship ids not yet used in a `.rels` part: `rIdDp1`, `rIdDp2` … */
export function relationshipIdAllocator(relsXml: string | undefined): () => string {
  const used = new Set([...(relsXml ?? '').matchAll(/\sId="([^"]*)"/g)].map((m) => m[1]));
  let n = 1;
  return () => {
    while (used.has(`rIdDp${n}`)) n++;
    used.add(`rIdDp${n}`);
    return `rIdDp${n}`;
  };
}

/**
 * Assemble the authored body into a clone of the original package, WITHOUT
 * handing it to a host.
 *
 * Split out of {@link materialise} for batch compare (ADR-0037), which needs the
 * redline package as bytes it can hold as a `Blob` — one artefact per variant,
 * materialised as each pair completes — rather than as a download or an
 * insert-into-active. Same assembly, no side effect.
 *
 * ADR-0047: the authored auxiliary parts replace the original's, parts the
 * original lacks are added with their relationship and content type, and the
 * body gets its last `<w:sectPr>` back.
 */
export function assembleRedlinePackage(
  authoredBodyXml: string,
  originalBundle: DocxBundle,
  aux: AuxiliaryParts = {},
): PizZipBundle {
  const documentXml = originalBundle.part('word/document.xml');
  if (documentXml === undefined) {
    throw new Error('assembleRedlinePackage: original bundle has no word/document.xml');
  }

  // Clone the original package so styles/numbering/section properties carry
  // over. PizZipBundle clones by re-zipping the original's bytes.
  const bundle = cloneBundle(originalBundle);
  const sectPr = aux.trailingSectPr ?? trailingSectPr(documentXml);
  const body = sectPr === null ? authoredBodyXml : authoredBodyXml.replace(/<\/w:body>\s*$/, `${sectPr}</w:body>`);
  bundle.setPart('word/document.xml', replaceBody(documentXml, body));
  for (const [path, xml] of aux.replace ?? []) bundle.setPart(path, xml);
  if (aux.add !== undefined && aux.add.length > 0) addParts(bundle, aux.add);
  return bundle;
}

/** Write new parts with their relationships and content-type overrides. */
function addParts(bundle: PizZipBundle, parts: readonly NewPart[]): void {
  let rels = bundle.part(DOC_RELS) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let types = bundle.part('[Content_Types].xml');
  for (const part of parts) {
    bundle.setPart(part.path, part.xml);
    if (part.rels !== undefined) types = copyPartRels(bundle, part, part.rels, types);
    const target = part.path.startsWith('word/') ? part.path.slice('word/'.length) : `/${part.path}`;
    rels = rels.replace('</Relationships>',
      `<Relationship Id="${part.rId}" Type="${REL_BASE}${part.relType}" Target="${target}"/></Relationships>`);
    if (types !== undefined && !types.includes(`PartName="/${part.path}"`)) {
      types = types.replace('</Types>',
        `<Override PartName="/${part.path}" ContentType="${CT_BASE}${part.relType}+xml"/></Types>`);
    }
  }
  bundle.setPart(DOC_RELS, rels);
  if (types !== undefined) bundle.setPart('[Content_Types].xml', types);
}

/**
 * Copy a new part's own relationships and the files they target. A target the
 * output already holds with the same bytes is shared; one it holds with other
 * bytes (the original's `image1.png` is not the revised one's) is copied under
 * a fresh name and the relationship pointed at it.
 */
function copyPartRels(bundle: PizZipBundle, part: NewPart, rels: PartRels, types: string | undefined): string | undefined {
  let xml = rels.xml;
  for (const file of rels.files) {
    let path = file.path;
    const existing = bundle.partBytes(path);
    if (existing !== undefined && !sameBytes(existing, file.bytes)) {
      const dot = path.lastIndexOf('.');
      let n = 1;
      while (bundle.hasPart(`${path.slice(0, dot)}-dp${n}${path.slice(dot)}`)) n++;
      path = `${path.slice(0, dot)}-dp${n}${path.slice(dot)}`;
      const renamed = file.target.slice(0, file.target.lastIndexOf('/') + 1) + path.slice(path.lastIndexOf('/') + 1);
      xml = xml.replace(`Target="${file.target}"`, `Target="${renamed}"`);
    }
    if (!bundle.hasPart(path)) bundle.setPartBytes(path, file.bytes);
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    if (types !== undefined && file.contentType !== undefined && !new RegExp(`Extension="${ext}"`, 'i').test(types)) {
      types = types.replace('</Types>', `<Default Extension="${ext}" ContentType="${file.contentType}"/></Types>`);
    }
  }
  const slash = part.path.lastIndexOf('/');
  bundle.setPart(`${part.path.slice(0, slash)}/_rels/${part.path.slice(slash + 1)}.rels`, xml);
  return types;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

/**
 * Assemble the authored body into a clone of the original package and
 * materialise it via the OOXML Service.
 *
 * Steps (§5.5):
 *   1. Clone the original bundle (preserves styles.xml, numbering.xml, etc.).
 *   2. Replace the <w:body> of word/document.xml with the authored body.
 *   3. Delegate to ooxmlService.writeBack.
 *
 * Default: mode 'newDocument' → writeBack(..., { mode: 'newDocument' }).
 * Insert:  mode 'insertIntoActive' → writeBack(..., { mode: 'replace', onArtifactDetected }).
 */
export async function materialise(
  authoredBodyXml: string,
  originalBundle: DocxBundle,
  opts: WriteoutOptions,
  aux: AuxiliaryParts = {},
): Promise<void> {
  // 1-2. Clone the original package and swap in the authored body (and parts).
  const bundle = assembleRedlinePackage(authoredBodyXml, originalBundle, aux);

  // 3. Delegate to the active HostPort (office: insert in place via the OOXML
  //    Service; web: serialise the bundle to a .docx Blob download).
  const host = getActiveHost();
  if (opts.mode === 'newDocument') {
    await host.writeBack(bundle, { mode: 'newDocument' });
    return;
  }

  // insertIntoActive → 'replace' with round-trip validation / fallback (§9.13).
  await host.writeBack(bundle, {
    mode: 'replace',
    onArtifactDetected: opts.onArtifactDetected ?? 'fallbackToNewDoc',
  });
}

/**
 * Clone a DocxBundle into a fresh, independently-mutable PizZipBundle.
 * If the bundle is already a PizZipBundle we round-trip its bytes; otherwise we
 * rebuild from its listed parts.
 */
function cloneBundle(source: DocxBundle): PizZipBundle {
  if (source instanceof PizZipBundle) {
    return new PizZipBundle(source.toUint8Array());
  }
  const clone = new PizZipBundle();
  for (const path of source.listParts()) {
    const xml = source.part(path);
    if (xml !== undefined) clone.setPart(path, xml);
  }
  return clone;
}
