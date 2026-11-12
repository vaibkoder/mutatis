/**
 * Feature 08 — Compare: parts diff for headers, footers, footnotes (I2).
 * Spec: spec/features/08-compare.md §9.7, §9.8, §7 (config keys excludeHeaders/Footers, footnotes);
 * ADR-0047.
 *
 * T10 acceptance: with `config.excludeHeaders` true there are zero DiffOps for header
 * parts. When enabled, each header/footer change is its own DiffOp, its `partId`
 * the role id, classified by content like a body change, and authored into the
 * redline's own part (ADR-0047).
 *
 * Headers and footers are paired by ROLE (ADR-0047): the section and the
 * reference type, never the file name. A role one document has and the other
 * does not is compared with what the other document SHOWS there — the part it
 * inherits from an earlier section ("link to previous") — or, failing that, with
 * nothing, so its paragraphs read as inserted or deleted. It is never skipped.
 */
import type { ChangeSummary, DiffOp, ParaRecord } from './model.js';
import type { CompareParts, HeaderFooterKind, HeaderFooterType, PartRef } from './io/index.js';
import { effectivePart, roleId, parseRelationships, resolveTarget } from './io/parts.js';
import type { CompareConfig } from '@shared/config/index.js';
import { preprocessSide } from './preprocess/index.js';
import { analyseRecords } from './pipeline.js';
import { makeAuthorContext, type AuthorContext } from './output/author.js';
import { authorNotes, type NoteChangeSet, type NoteKind } from './output/author-notes.js';
import { analyseAndAuthorPart, prefixPartOps } from './output/author-part.js';
import type { HeaderFooterPlan } from './output/package-parts.js';
import { trailingSectPr, type AuxiliaryParts, type NewPart, type PartRels } from './output/writeout.js';
import type { DocxBundle } from '@shared/ooxml/index.js';
import type { SearchScope } from '@shared/host/index.js';

export type PartKind = HeaderFooterKind | NoteKind;

/** One part's records and ops (ADR-0047). */
export interface PartChangeSet {
  /**
   * The part's id: a header/footer role id (`header:s0:default`), a note
   * (`footnote:3` — its number in the document), or `footnotes`/`endnotes` for
   * a notes part compared whole because its notes could not be authored.
   */
  partId: string;
  kind: PartKind;
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  /** Every op, `equal` included; ids prefixed with the part id. */
  ops: DiffOp[];
  /** The changes are tracked in the redline's own part (false: listed, not marked). */
  authored: boolean;
  /** Where the change is: the change list's badge, the export's locator, the host's search scope. */
  location: PartLocation;
}

/** A part change's place in the document (ADR-0047). */
export interface PartLocation {
  /** e.g. "Footer", "First-page header, section 2", "Footnote 3 (added)". */
  label: string;
  /** What "go to change" can and cannot do there — shown as the row's tooltip. */
  hint: string;
  /** The story the host searches. */
  scope: SearchScope;
}

const TYPE_LABEL: Record<HeaderFooterType, string> = { default: '', first: 'First-page ', even: 'Even-page ' };
const NOTE_HINT =
  'Go to change searches the notes where Word supports it (WordApi 1.5); otherwise it searches the main text.';

/** A header or footer role's location; the section is named only when the document has more than one. */
export function headerFooterLocation(
  kind: HeaderFooterKind,
  section: number,
  type: HeaderFooterType,
  sectionCount: number,
): PartLocation {
  const name = `${TYPE_LABEL[type]}${kind}`;
  const label = `${name[0]!.toUpperCase()}${name.slice(1)}${sectionCount > 1 ? `, section ${section + 1}` : ''}`;
  return { label, hint: `Go to change searches this ${kind} in Word.`, scope: { story: kind, section, type } };
}

/** A note's location: its number in the document, and whether it was added or removed. */
export function noteLocation(kind: NoteKind, ordinal: number | null, status: NoteChangeSet['status'] | null): PartLocation {
  const noun = kind === 'footnote' ? 'Footnote' : 'Endnote';
  const label = ordinal === null
    ? `${noun}s`
    : `${noun} ${ordinal}${status === 'inserted' ? ' (added)' : status === 'deleted' ? ' (removed)' : ''}`;
  return { label, hint: NOTE_HINT, scope: { story: kind } };
}

export interface PartsDiffResult {
  /** DiffOps for header/footer/footnote parts (empty arrays if excluded). */
  ops: DiffOp[];
  /** The same ops grouped by part, with each part's records. */
  parts: PartChangeSet[];
  /** Counts for summary reporting. */
  headerChanges: number;
  footerChanges: number;
  footnoteChanges: number;
}

/** Inputs to {@link diffAndAuthorParts}. */
export interface PartsInput {
  originalParts: CompareParts;
  revisedParts: CompareParts;
  /** The authored body (`authoredBodyXml`) — the notes follow its references. */
  bodyXml: string;
  config: CompareConfig;
  author: string;
  date: string;
  /** Relationship ids free in the original package (a part the original lacks needs one). */
  newRelationshipId: () => string;
  /**
   * Where each header and footer goes ({@link planHeaderFooters}, run BEFORE the
   * pipeline so the revised body's section references point into the output).
   * Absent ⇒ no header or footer is compared.
   */
  plan?: HeaderFooterPlan;
  /** The two packages, for a new part's own relationships (a logo in an added header). */
  bundles?: { original: DocxBundle; revised: DocxBundle };
}

/** {@link PartsDiffResult}, plus what the redline package carries (ADR-0047). */
export interface PartsResult extends PartsDiffResult {
  /** The authored body, every note reference pointing at its own note. */
  bodyXml: string;
  /** The authored parts for `assembleRedlinePackage`. */
  aux: AuxiliaryParts;
}

/** The next `w:id` after every one in `xml`, so a part's revisions never repeat a body id. */
export function nextRevisionId(xml: string): number {
  let max = 0;
  for (const m of xml.matchAll(/\sw:id="(\d+)"/g)) max = Math.max(max, Number(m[1]));
  return max + 1;
}

/**
 * Compare and author the auxiliary parts (ADR-0047): footnotes and endnotes
 * (`compare.footnotes`), headers and footers (`compare.excludeHeaders`/
 * `excludeFooters`). Returns the listed ops, the authored body (its note
 * references retargeted) and the parts for the package.
 */
export async function diffAndAuthorParts(input: PartsInput): Promise<PartsResult> {
  const { originalParts, config, author, date } = input;
  const ctx = makeAuthorContext(author, date, nextRevisionId(input.bodyXml));
  const parts: PartChangeSet[] = [];
  const replace = new Map<string, string>();
  const add: NewPart[] = [];
  let bodyXml = input.bodyXml;
  let footnoteChanges = 0;
  let headerChanges = 0;
  let footerChanges = 0;
  const sectionCount = Math.max(originalParts.sectionCount, input.revisedParts.sectionCount);

  for (const planned of input.plan?.parts ?? []) {
    const { pair } = planned;
    const result = await analyseAndAuthorPart({
      partId: pair.roleId,
      originalXml: pair.original?.xml ?? null,
      revisedXml: pair.revised?.xml ?? null,
      config, author, date, ctx,
    });
    const changed = result.ops.filter(op => op.type !== 'equal').length;
    if (pair.kind === 'header') headerChanges += changed;
    else footerChanges += changed;
    if (result.changed) {
      parts.push({
        partId: pair.roleId, kind: pair.kind, originalParas: result.originalParas,
        revisedParas: result.revisedParas, ops: result.ops, authored: typeof result.blocksXml === 'string',
        location: headerFooterLocation(pair.kind, pair.section, pair.type, sectionCount),
      });
    }
    // The part as it goes out: authored; or, unchanged or refused by its check,
    // the original's part where it has one at this path (as before ADR-0047),
    // else the part as the side that has it wrote it.
    const shell = planned.isNew ? (pair.revised ?? pair.original)! : (pair.original ?? pair.revised)!;
    const xml = typeof result.blocksXml === 'string' ? withBlocks(shell.xml, result.blocksXml) : shell.xml;
    if (!planned.isNew) {
      if (typeof result.blocksXml === 'string') replace.set(planned.path, xml);
      continue;
    }
    const from = pair.revised !== null ? input.bundles?.revised : input.bundles?.original;
    const rels = from === undefined ? undefined : partRels(from, shell.path);
    add.push({ path: planned.path, xml, relType: pair.kind, rId: planned.rId, ...(rels ? { rels } : {}) });
  }

  if (config.footnotes === 'compare') {
    for (const kind of ['footnote', 'endnote'] as const) {
      const notes = await authorNotesOf(kind, bodyXml, input, ctx);
      bodyXml = notes.bodyXml;
      footnoteChanges += notes.parts.reduce((n, s) => n + s.ops.filter(op => op.type !== 'equal').length, 0);
      parts.push(...notes.parts);
      if (notes.write === null) continue;
      const path = originalParts.notesPaths?.[`${kind}s`] ?? null;
      if (path !== null) replace.set(path, notes.write);
      else add.push({ path: `word/${kind}s.xml`, xml: notes.write, relType: `${kind}s`, rId: input.newRelationshipId() });
    }
  }

  // The body's last sectPr, from the revised document as its references were
  // remapped into the output package (ADR-0047).
  const sectPr = input.plan === undefined ? null : trailingSectPr(input.plan.revisedDocumentXml);
  return {
    ops: parts.flatMap(s => s.ops),
    parts,
    headerChanges,
    footerChanges,
    footnoteChanges,
    bodyXml,
    aux: { replace, add, ...(sectPr !== null ? { trailingSectPr: sectPr } : {}) },
  };
}

/**
 * One kind of note: authored (each changed note one part), or — when the
 * self-check refuses — the pre-ADR-0047 behaviour: the original part kept, the
 * references unchanged, and the part compared whole so its changes are still
 * listed (`authored: false`).
 */
async function authorNotesOf(
  kind: NoteKind,
  bodyXml: string,
  input: PartsInput,
  ctx: AuthorContext,
): Promise<{ bodyXml: string; write: string | null; parts: PartChangeSet[] }> {
  const { originalParts, revisedParts, config, author, date } = input;
  const xmlOf = (p: CompareParts): string | null => (kind === 'footnote' ? p.footnotesXml : p.endnotesXml);
  const authored = await authorNotes({
    kind, bodyXml, config, author, date, ctx,
    original: { documentXml: originalParts.documentXml, notesXml: xmlOf(originalParts) },
    revised: { documentXml: revisedParts.documentXml, notesXml: xmlOf(revisedParts) },
  });
  if (authored !== null) {
    return {
      bodyXml: authored.bodyXml,
      write: authored.notesXml,
      parts: authored.changed.map(c => ({
        partId: c.partId, kind, originalParas: c.originalParas, revisedParas: c.revisedParas,
        ops: c.ops, authored: c.authored, location: noteLocation(kind, c.ordinal, c.status),
      })),
    };
  }
  // Compared whole, by the body's analysis (so classified by content), not marked.
  const o = xmlOf(originalParts);
  const r = xmlOf(revisedParts);
  const partId = `${kind}s`;
  const originalParas = o === null ? [] : preprocessSide(o, partId, config.flattenOriginal, config).paras;
  const revisedParas = r === null ? [] : preprocessSide(r, partId, config.flattenRevised, config).paras;
  const { ops: raw } = await analyseRecords({ originalParas, revisedParas, config, author, date });
  const ops = prefixPartOps(raw, partId);
  const changed = ops.some(op => op.type !== 'equal');
  return {
    bodyXml, write: null,
    parts: changed
      ? [{ partId, kind, originalParas, revisedParas, ops, authored: false, location: noteLocation(kind, null, null) }]
      : [],
  };
}

/** A header/footer role and what each document shows for it. */
export interface RolePair {
  roleId: string;
  kind: HeaderFooterKind;
  section: number;
  type: HeaderFooterType;
  original: PartRef | null;
  revised: PartRef | null;
}

/**
 * Pair the two documents' headers (or footers) by role. Every role either
 * document references is one pair; a side without a reference of its own
 * contributes the part it inherits there, or `null`. Two roles that resolve to
 * the same two parts are one pair (the first), so a part shared by sections is
 * compared once.
 */
export function pairRoles(
  original: CompareParts,
  revised: CompareParts,
  kind: HeaderFooterKind,
): RolePair[] {
  const seen = new Set<string>();
  const keys = [...original.roles, ...revised.roles]
    .filter((r) => r.kind === kind)
    .sort((a, b) => a.section - b.section || TYPE_ORDER[a.type] - TYPE_ORDER[b.type]);
  const pairs: RolePair[] = [];
  const shown = new Set<string>();
  for (const { section, type } of keys) {
    const id = roleId(kind, section, type);
    if (seen.has(id)) continue;
    seen.add(id);
    const o = effectivePart(original, kind, section, type)?.part ?? null;
    const r = effectivePart(revised, kind, section, type)?.part ?? null;
    const both = `${o?.path ?? ''}|${r?.path ?? ''}`;
    if (shown.has(both)) continue;
    shown.add(both);
    pairs.push({ roleId: id, kind, section, type, original: o, revised: r });
  }
  return pairs;
}

const TYPE_ORDER: Record<HeaderFooterType, number> = { default: 0, first: 1, even: 2 };

/** A part's root open tag (`<w:hdr …>` with its namespaces) wrapped round authored blocks. */
function withBlocks(partXml: string, blocks: string): string {
  const open = /^[\s\S]*?<w:(hdr|ftr)\b[^>]*>/.exec(partXml);
  if (open === null) return partXml;
  return `${open[0]}${blocks}</w:${open[1]}>`;
}

/** A part's own relationships and the internal files they target, from its package. */
function partRels(bundle: DocxBundle, path: string): PartRels | undefined {
  const slash = path.lastIndexOf('/');
  const dir = path.slice(0, slash);
  const xml = bundle.part(`${dir}/_rels/${path.slice(slash + 1)}.rels`);
  if (xml === undefined) return undefined;
  const types = bundle.part('[Content_Types].xml') ?? '';
  const files: PartRels['files'] = [];
  for (const rel of parseRelationships(xml)) {
    if (rel.external) continue;
    const target = resolveTarget(dir, rel.target);
    const bytes = bundle.partBytes?.(target);
    if (bytes === undefined) continue;
    const ext = target.slice(target.lastIndexOf('.') + 1);
    const contentType = new RegExp(`<Default Extension="${ext}" ContentType="([^"]*)"`, 'i').exec(types)?.[1];
    files.push({ target: rel.target, path: target, bytes, ...(contentType ? { contentType } : {}) });
  }
  return { xml, files };
}

/**
 * Fold the auxiliary parts' changes into the body summary's counts (§8.2, §9.8).
 *
 * `summary.total` is *"ALWAYS the full count"* (§8.2) and `byCategory` partitions
 * it, so a change found outside the body belongs in both. Until this existed the
 * driver kept only the three integers below, which meant a **footnote content
 * change — substantive, and compared BY DEFAULT (§9.8)** — reached the reader as
 * a line in the "Other parts" block and was absent from the headline number and
 * from every category count. §9.7's *"reported separately"* licence covers
 * headers and footers; §9.8 grants no such thing, and says the opposite.
 *
 * The part ops are ALSO rows of the navigable change list (ADR-0047): the
 * driver lists them after the body's and keeps each part's records under its
 * `partId` (`CompareDriverState.parts`), because a part op's range indexes its
 * OWN part's paragraphs — resolved against the body's arrays, the same integer
 * would land on an unrelated body clause. Every reader of a range goes through
 * that map (`driver.ts recordsFor`, `change-set-export.ts`).
 */
export function mergePartsIntoSummary(
  summary: ChangeSummary,
  parts: PartsDiffResult,
): ChangeSummary {
  const byCategory = { ...summary.byCategory };
  let total = summary.total;
  for (const op of parts.ops) {
    if (op.type === 'equal') continue;
    byCategory[op.classification] = (byCategory[op.classification] ?? 0) + 1;
    total += 1;
  }
  return {
    ...summary,
    total,
    byCategory,
    partChanges: {
      headers: parts.headerChanges,
      footers: parts.footerChanges,
      footnotes: parts.footnoteChanges,
    },
  };
}
