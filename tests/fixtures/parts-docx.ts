/**
 * Synthetic packages with headers, footers, footnotes and endnotes
 * (ADR-0047).
 *
 * A package is described by its body, its sections' header/footer references
 * and its part files, and is wired the way Word wires one: every reference goes
 * through `word/_rels/document.xml.rels`, every part has a content-type
 * override, and the body's last `<w:sectPr>` is a child of `<w:body>`. Every
 * string is invented — no real document content.
 */
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import type { CompareConfig } from '@shared/config/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import type { CompareParts } from '@features/compare/io/parts.js';
import { diffAndAuthorParts, type PartsResult } from '@features/compare/parts-diff.js';
import { planHeaderFooters } from '@features/compare/output/package-parts.js';
import { relationshipIdAllocator } from '@features/compare/output/index.js';

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W_NS}" xmlns:r="${R_NS}"`;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** One run of plain text. */
export const run = (text: string): string => `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;

/** A paragraph of plain text. */
export const para = (text: string): string => `<w:p>${text === '' ? '' : run(text)}</w:p>`;

/** A footnote reference run (Word's shape). */
export const fnRef = (id: number): string =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;

/** An endnote reference run. */
export const enRef = (id: number): string =>
  `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;

/** A body paragraph from text and reference runs, e.g. `p('The Buyer', fnRef(1), ' shall pay.')`. */
export function p(...pieces: string[]): string {
  return `<w:p>${pieces.map((x) => (x.startsWith('<w:r>') ? x : run(x))).join('')}</w:p>`;
}

/** A PAGE field (complex), its cached result `result`. */
export const pageField = (result: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

export type RefType = 'default' | 'first' | 'even';

/** One section: its references (kind + type → part file name) and whether it is the last. */
export interface SectionSpec {
  headers?: Partial<Record<RefType, string>>;
  footers?: Partial<Record<RefType, string>>;
}

/** A note: its paragraphs' texts. */
export type NoteSpec = readonly string[];

export interface PackageSpec {
  /**
   * Body paragraphs. A section break is the string `'§'`, which closes the
   * current section with an in-paragraph `<w:sectPr>`; the body's last section
   * is closed by a trailing `<w:sectPr>`.
   */
  body: readonly string[];
  /** One spec per section, in order (missing ⇒ no references). */
  sections?: readonly SectionSpec[];
  /** Header/footer files by name (`header1.xml` …) → their paragraphs (raw XML). */
  parts?: Record<string, readonly string[]>;
  /** Footnotes by id (1, 2 …) → paragraph texts. Separators -1/0 are added. */
  footnotes?: Record<number, NoteSpec>;
  endnotes?: Record<number, NoteSpec>;
}

function sectPr(spec: SectionSpec | undefined, rIdOf: (file: string) => string): string {
  const refs: string[] = [];
  for (const kind of ['header', 'footer'] as const) {
    const map = kind === 'header' ? spec?.headers : spec?.footers;
    for (const [type, file] of Object.entries(map ?? {})) {
      refs.push(`<w:${kind}Reference w:type="${type}" r:id="${rIdOf(file)}"/>`);
    }
  }
  return `<w:sectPr>${refs.join('')}<w:pgSz w:w="11906" w:h="16838"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>';
}

/** A note paragraph: the note's own number mark first, as Word writes it. */
function noteParagraph(kind: 'footnote' | 'endnote', text: string, first: boolean): string {
  const mark = first
    ? `<w:r><w:rPr><w:rStyle w:val="${kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference'}"/></w:rPr><w:${kind}Ref/></w:r>${run(' ')}`
    : '';
  return `<w:p><w:pPr><w:pStyle w:val="${kind === 'footnote' ? 'FootnoteText' : 'EndnoteText'}"/></w:pPr>${mark}${text === '' ? '' : run(text)}</w:p>`;
}

/** A whole notes part (`w:footnotes` / `w:endnotes`), separators included. */
export function notesXml(kind: 'footnote' | 'endnote', notes: Record<number, NoteSpec>): string {
  const sep =
    `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
  const body = Object.entries(notes)
    .map(([id, paras]) =>
      `<w:${kind} w:id="${id}">${paras.map((t, i) => noteParagraph(kind, t, i === 0)).join('')}</w:${kind}>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s ${NS}>${sep}${body}</w:${kind}s>`;
}

/** Build the package. */
export function partsPackage(spec: PackageSpec): PizZipBundle {
  const bundle = synthesiseDocx('placeholder');
  const rels: string[] = [
    `<Relationship Id="rId1" Type="${REL}styles" Target="styles.xml"/>`,
    `<Relationship Id="rId2" Type="${REL}settings" Target="settings.xml"/>`,
  ];
  const overrides: string[] = [];
  let next = 10;
  const rIds = new Map<string, string>();
  const rIdOf = (file: string): string => {
    let id = rIds.get(file);
    if (id === undefined) {
      id = `rId${next++}`;
      rIds.set(file, id);
      const kind = file.startsWith('header') ? 'header' : 'footer';
      rels.push(`<Relationship Id="${id}" Type="${REL}${kind}" Target="${file}"/>`);
      overrides.push(`<Override PartName="/word/${file}" ContentType="${CT}${kind}+xml"/>`);
    }
    return id;
  };

  const sections = spec.sections ?? [];
  let section = 0;
  const blocks: string[] = [];
  for (const b of spec.body) {
    if (b === '§') {
      blocks.push(`<w:p><w:pPr>${sectPr(sections[section], rIdOf)}</w:pPr></w:p>`);
      section++;
    } else {
      blocks.push(b.startsWith('<') ? b : para(b));
    }
  }
  const trailing = sectPr(sections[section], rIdOf);
  bundle.setPart(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${blocks.join('')}${trailing}</w:body></w:document>`,
  );

  for (const [file, paras] of Object.entries(spec.parts ?? {})) {
    const root = file.startsWith('header') ? 'w:hdr' : 'w:ftr';
    bundle.setPart(`word/${file}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} ${NS}>${paras.join('')}</${root}>`);
    rIdOf(file); // an unreferenced part still gets a relationship, as Word keeps one
  }
  for (const kind of ['footnote', 'endnote'] as const) {
    const notes = kind === 'footnote' ? spec.footnotes : spec.endnotes;
    if (notes === undefined) continue;
    bundle.setPart(`word/${kind}s.xml`, notesXml(kind, notes));
    rels.push(`<Relationship Id="rId${next++}" Type="${REL}${kind}s" Target="${kind}s.xml"/>`);
    overrides.push(`<Override PartName="/word/${kind}s.xml" ContentType="${CT}${kind}s+xml"/>`);
  }

  bundle.setPart(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
  );
  const types = bundle.part('[Content_Types].xml')!;
  bundle.setPart('[Content_Types].xml', types.replace('</Types>', `${overrides.join('')}</Types>`));
  return bundle;
}

/** Copy a bundle (so a test can mutate one side). */
export function copyBundle(b: PizZipBundle): PizZipBundle {
  return new PizZipBundle(b.toUint8Array());
}

/**
 * Compare two packages' parts as the driver does (plan, then author), with an
 * empty body — the headers and footers only.
 */
export async function compareHeadersFooters(
  original: CompareParts,
  revised: CompareParts,
  config: CompareConfig,
  author = 'Compare',
  date = '2026-10-26T00:00:00Z',
): Promise<PartsResult> {
  const newRelationshipId = relationshipIdAllocator(undefined);
  const plan = planHeaderFooters(original, revised, config, [], newRelationshipId);
  return diffAndAuthorParts({
    originalParts: original,
    revisedParts: { ...revised, documentXml: plan.revisedDocumentXml },
    bodyXml: '<w:body/>',
    config, author, date, newRelationshipId, plan,
  });
}
