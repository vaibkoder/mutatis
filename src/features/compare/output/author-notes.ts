/**
 * Feature 08 — Compare: footnotes and endnotes authored into the redline, every
 * reference pointing at its own note (ADR-0047; retires ADR-0045 § Accepted
 * limits 9(a)).
 *
 * Until ADR-0047 the redline package was the ORIGINAL's with only the body
 * replaced: a note whose text changed kept its original text, a note added had
 * no text at all, and a reference the body took from the revised document
 * pointed at whatever the original's notes part held under that id.
 *
 * **Notes are matched by place, read from the authored body itself.** Every
 * reference in the authored body is in the accept-all view, the reject-all view,
 * or both. The accept-all view's references are the revised document's, in
 * order; the reject-all view's are the original's (the author's own self-checks
 * hold every paragraph to that). So a reference in BOTH views pairs the original
 * note at its reject-all position with the revised note at its accept-all
 * position — kept, and its note diffed; one in accept-all only is a note
 * inserted; one in reject-all only a note struck. The body's diff has already
 * decided which references are kept, so the notes follow it and cannot disagree
 * with it. Two references that swapped notes (`Seller¹ shall³` → `Buyer³
 * shall¹`) pair each place's two notes, and the note text under the place is
 * diffed: the change is shown (ADR-0045 limit 9(a), retired).
 *
 * **Ids are re-dealt.** A kept or inserted note keeps its REVISED id, so a
 * reference the author took from the revised paragraph needs no change; a
 * struck note gets a fresh id above every id in either part, and its reference
 * — taken from the original paragraph, carrying the original's id, which in the
 * revised id space can be another note — is retargeted to it. Separator notes
 * (`w:type`) are the original's. Footnote numbers are not compared (§9.8).
 *
 * **Self-check (ADR-0008).** Each note is held to the 08i yardstick by
 * `analyseAndAuthorPart`; then the emitted body and notes part are read back as
 * a reader would: in each view, every reference must resolve to exactly one
 * note, and the note under the n-th reference must read as the note under the
 * n-th reference of that view's document. Anything else returns `null` and the
 * caller keeps the pre-ADR-0047 output (the original's notes part, references
 * unchanged) — recorded, never a silent loss.
 *
 * Spec: spec/features/08-compare.md §5.5, §9.8.
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { DiffOp, ParaRecord } from '../model.js';
import { flattenTrackedChanges, normaliseText } from '../preprocess/index.js';
import type { AuthorContext } from './author.js';
import { analyseAndAuthorPart } from './author-part.js';
import { paragraphViews } from './reconstruction.js';

export type NoteKind = 'footnote' | 'endnote';

/** One side's input: its body and its notes part, both as the pipeline compared them. */
export interface NoteSide {
  /** `word/document.xml` (any revisions it carries are resolved by the side's policy). */
  documentXml: string;
  /** The notes part (`w:footnotes` / `w:endnotes`), or null when the package has none. */
  notesXml: string | null;
}

/** One authored note and how it got there. */
export interface NoteChangeSet {
  /** `footnote:4` — the note's id in the authored notes part (unique in the redline). */
  partId: string;
  kind: NoteKind;
  /** 1-based: the note's place among the references of the revised document (the original for a struck note). */
  ordinal: number;
  status: 'kept' | 'inserted' | 'deleted';
  originalParas: ParaRecord[];
  revisedParas: ParaRecord[];
  /** Every op, ids prefixed with `partId`. */
  ops: DiffOp[];
  /** The note's changes are tracked in the output (false: listed only). */
  authored: boolean;
}

export interface AuthoredNotes {
  kind: NoteKind;
  /** The authored body with every reference of this kind pointing at its note. */
  bodyXml: string;
  /** The authored notes part; null when neither body references a note (keep the original's). */
  notesXml: string | null;
  /** Every note that is not unchanged, in body order. */
  changed: NoteChangeSet[];
}

export interface NoteAuthoringInput {
  kind: NoteKind;
  /** The authored `<w:body>…</w:body>`. */
  bodyXml: string;
  original: NoteSide;
  revised: NoteSide;
  config: CompareConfig;
  author: string;
  date: string;
  ctx: AuthorContext;
}

/** A note element of a notes part. */
interface NoteElement {
  id: string;
  /** Separator, continuation separator or notice — never referenced from the body. */
  special: boolean;
  xml: string;
  open: string;
}

const TAG = 'dpnote';

/** `<mc:Fallback>` repeats its Choice's content for old readers; it is not what Word shows. */
const stripFallback = (xml: string): string => xml.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, '');

const referenceRe = (kind: NoteKind): RegExp => new RegExp(`<w:${kind}Reference\\b[^>]*>`, 'g');

const idOf = (tag: string): string | undefined => /\sw:id="([^"]*)"/.exec(tag)?.[1];

/** The note ids a body references, in document order (Fallback content skipped). */
export function referenceIds(bodyXml: string, kind: NoteKind): string[] {
  return [...stripFallback(bodyXml).matchAll(referenceRe(kind))].map((m) => idOf(m[0]) ?? '');
}

/** The notes of a notes part, in part order. */
function noteElements(notesXml: string | null, kind: NoteKind): NoteElement[] {
  if (notesXml === null) return [];
  const re = new RegExp(`<w:${kind}\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/w:${kind}>)`, 'g');
  const out: NoteElement[] = [];
  for (const m of notesXml.matchAll(re)) {
    const open = /^<[^>]*>/.exec(m[0])![0];
    out.push({ id: idOf(open) ?? '', special: /\sw:type="/.test(open), xml: m[0], open });
  }
  return out;
}

/** Set the `w:id` of an element's open tag. */
const withId = (tag: string, id: string): string => tag.replace(/(\sw:id=")[^"]*(")/, `$1${id}$2`);

/** The notes part's prolog and root open tag, and its close tag. */
function rootOf(notesXml: string, kind: NoteKind): { open: string; close: string } {
  const m = new RegExp(`^[\\s\\S]*?<w:${kind}s\\b[^>]*>`).exec(notesXml);
  return { open: m?.[0] ?? `<w:${kind}s>`, close: `</w:${kind}s>` };
}

/**
 * Author both sides' notes of one kind into one notes part and retarget the
 * body's references to it. `null` when the self-check fails (keep the pre-ADR
 * output).
 */
export async function authorNotes(input: NoteAuthoringInput): Promise<AuthoredNotes | null> {
  const { kind, config, author, date, ctx } = input;
  const origDoc = flattenTrackedChanges(input.original.documentXml, config.flattenOriginal);
  const revDoc = flattenTrackedChanges(input.revised.documentXml, config.flattenRevised);
  const origSeq = referenceIds(origDoc, kind);
  const revSeq = referenceIds(revDoc, kind);
  // Neither body references a note: nothing to author, the original part stays.
  if (origSeq.length === 0 && revSeq.length === 0) {
    return { kind, bodyXml: input.bodyXml, notesXml: null, changed: [] };
  }
  const origNotes = new Map(noteElements(input.original.notesXml, kind).map((n) => [n.id, n]));
  const revNotes = new Map(noteElements(input.revised.notesXml, kind).map((n) => [n.id, n]));

  // Tag every reference of the authored body, then read each tag's place in both views.
  const sourceIds: string[] = [];
  const tagged = input.bodyXml.replace(referenceRe(kind), (tag) => {
    sourceIds.push(idOf(tag) ?? '');
    return withId(tag, `${TAG}${sourceIds.length - 1}`);
  });
  const positions = (policy: 'acceptAll' | 'rejectAll'): Map<number, number> => {
    const tags = referenceIds(flattenTrackedChanges(tagged, policy), kind);
    return new Map(tags.map((t, i) => [Number(t.slice(TAG.length)), i]));
  };
  const accepted = positions('acceptAll');
  const rejected = positions('rejectAll');

  // Pair each tag with its notes; deal the output ids.
  const fresh = freshIdFrom([...origNotes.keys(), ...revNotes.keys(), ...origSeq, ...revSeq]);
  interface Slot { outId: string; status: NoteChangeSet['status']; o: string | null; r: string | null; ordinal: number }
  const slots = new Map<string, Slot>();
  const struckIds = new Map<string, string>();
  const targets: string[] = [];
  for (let k = 0; k < sourceIds.length; k++) {
    const a = accepted.get(k);
    const j = rejected.get(k);
    if (a === undefined && j === undefined) {
      targets.push(sourceIds[k]!); // in neither view (e.g. Fallback): left as it was
      continue;
    }
    // A view with more references than its document: the read-back check
    // below refuses it; until then the slot simply has no note to author.
    const r = a === undefined ? null : revSeq[a] ?? '';
    const o = j === undefined ? null : origSeq[j] ?? '';
    let outId: string;
    if (r !== null) {
      outId = r;
    } else {
      outId = struckIds.get(o!) ?? String(fresh());
      struckIds.set(o!, outId);
    }
    const slot: Slot = {
      outId, o, r,
      status: r === null ? 'deleted' : o === null ? 'inserted' : 'kept',
      ordinal: (a ?? j!) + 1,
    };
    // One revised note referenced twice is one note, authored once from its
    // first pairing; if the second pairing differs, the read-back check refuses.
    if (!slots.has(outId)) slots.set(outId, slot);
    targets.push(outId);
  }

  // Author every note.
  const changed: NoteChangeSet[] = [];
  const authoredNotes: string[] = [];
  for (const slot of slots.values()) {
    const on = slot.o === null ? undefined : origNotes.get(slot.o);
    const rn = slot.r === null ? undefined : revNotes.get(slot.r);
    if ((slot.o !== null && on === undefined) || (slot.r !== null && rn === undefined)) return null;
    // The OUTPUT id: unique even where a note struck and a note inserted share a number.
    const partId = `${kind}:${slot.outId}`;
    const result = await analyseAndAuthorPart({
      partId,
      originalXml: on?.xml ?? null,
      revisedXml: rn?.xml ?? null,
      config, author, date, ctx,
    });
    const shell = rn ?? on!;
    const open = withId(shell.open, slot.outId);
    if (result.blocksXml === undefined) {
      // Unchanged: the revised note as the pipeline read it.
      authoredNotes.push(withId(flattenTrackedChanges(rn!.xml, config.flattenRevised), slot.outId));
      continue;
    }
    if (result.blocksXml === null) return null;
    authoredNotes.push(`${open}${result.blocksXml}</w:${kind}>`);
    changed.push({
      partId, kind, ordinal: slot.ordinal, status: slot.status,
      originalParas: result.originalParas, revisedParas: result.revisedParas,
      ops: result.ops, authored: true,
    });
  }

  let i = 0;
  const bodyXml = tagged.replace(referenceRe(kind), (tag) => withId(tag, targets[i++]!));
  const base = input.original.notesXml ?? input.revised.notesXml;
  if (base === null) return null;
  const separators = noteElements(base, kind).filter((n) => n.special).map((n) => n.xml);
  const root = rootOf(base, kind);
  const notesXml = `${root.open}${separators.join('')}${authoredNotes.join('')}${root.close}`;

  const check = (policy: 'acceptAll' | 'rejectAll', doc: string, notes: string | null, exact: boolean): boolean =>
    sameNotes(resolvedNotes(bodyXml, notesXml, kind, policy), resolvedNotes(doc, notes, kind, policy), exact);
  if (!check('acceptAll', revDoc, flattenNotes(input.revised.notesXml, config.flattenRevised), true)) return null;
  if (!check('rejectAll', origDoc, flattenNotes(input.original.notesXml, config.flattenOriginal), false)) return null;
  return { kind, bodyXml, notesXml, changed };
}

const flattenNotes = (xml: string | null, policy: CompareConfig['flattenOriginal']): string | null =>
  xml === null ? null : flattenTrackedChanges(xml, policy);

/** An id generator above every numeric id in sight. */
function freshIdFrom(ids: readonly string[]): () => number {
  let next = Math.max(0, ...ids.map(Number).filter(Number.isFinite)) + 1;
  return () => next++;
}

/**
 * What a reader sees in one view: for each reference of the body, in order, the
 * text of the note it resolves to (its paragraphs joined by a newline), or
 * `null` where it resolves to no note or to more than one.
 */
export function resolvedNotes(
  bodyXml: string,
  notesXml: string | null,
  kind: NoteKind,
  policy: 'acceptAll' | 'rejectAll',
): (string | null)[] {
  const byId = new Map<string, NoteElement[]>();
  for (const n of noteElements(notesXml === null ? null : flattenTrackedChanges(notesXml, policy), kind)) {
    if (n.special) continue;
    byId.set(n.id, [...(byId.get(n.id) ?? []), n]);
  }
  return referenceIds(flattenTrackedChanges(bodyXml, policy), kind).map((id) => {
    const notes = byId.get(id);
    if (notes?.length !== 1) return null;
    return paragraphViews(notes[0]!.xml).map((v) => v.text).join('\n');
  });
}

/** Two views' notes agree: exactly (accept-all), or under §5.2's fold (reject-all — 08i limit 14). */
function sameNotes(actual: (string | null)[], expected: (string | null)[], exact: boolean): boolean {
  if (actual.length !== expected.length) return false;
  return actual.every((a, i) => {
    const e = expected[i];
    if (a === null || e === null || e === undefined) return false;
    return exact ? a === e : normaliseText(a) === normaliseText(e);
  });
}
