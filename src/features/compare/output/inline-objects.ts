/**
 * Feature 08 — Compare: author a changed paragraph without losing what is not
 * text (ADR-0045).
 *
 * `emitReplace` rebuilds a changed paragraph from its word diff, and a word
 * diff is text. Before ADR-0045 everything in the paragraph that was not text
 * went with the rebuild: a footnote reference vanished, a hyperlink became
 * plain words, a REF field became its frozen result, a text box was poured into
 * the sentence it was anchored in, and a bookmark another clause's REF field
 * pointed at was dropped. This module puts those back.
 *
 * The rule, per object (a unit from `preprocess/inline-units.ts`):
 *
 * - **Its text is untouched by the diff, and it lines up with an object of the
 *   same sort on the other side** → it is re-emitted VERBATIM from the revised
 *   paragraph, as unchanged content around the changed text.
 * - **Otherwise** → it is emitted WHOLE: the revised object as one tracked
 *   insertion and the original object as one tracked deletion, insertion first
 *   (ADR-0043). Every character of both sides is still on the page, inside one
 *   of the two objects, so nothing the diff found is hidden (ADR-0008).
 * - **Zero-width** objects (a footnote reference, an image) have no text to
 *   line up by. When both sides carry the same sequence of sorts they are
 *   re-emitted from the revised side; when not, every one of them is emitted as
 *   a tracked insertion or deletion — over-marking, never under-marking.
 * - **Bookmarks** are re-emitted from the revised side and never revised.
 *
 * Ordinary text is rebuilt exactly as before, so a changed paragraph with no
 * objects in it never reaches this module and is authored byte for byte as it
 * always was.
 *
 * Every result is checked before it is used: the authored paragraph must read,
 * with all its revisions accepted, as the revised text, and with all of them
 * rejected, as the original. A paragraph that fails is authored the old way.
 *
 * Spec: spec/features/08-compare.md §5.5.
 */
import type { ParaRecord, WordDiff, ZeroTextObject } from '../model.js';
import { computeLCS } from '../diff/lcs.js';
import { flattenTrackedChanges } from '../preprocess/flatten.js';
import {
  paragraphUnits, paragraphNode, paragraphText, buildNodes, parseNodes, isPlainTextNode, collectText,
  zeroTextObjects, textLeaves, type InlineUnit, type TextLeaf,
} from '../preprocess/inline-units.js';
import { normaliseText, normaliseTextWithOffsets, stripNumberingPrefix } from '../preprocess/normalise.js';
import { tagOf, childrenOf, attrOf, type PoNode } from '@shared/ooxml/po-xml.js';

import { wrapIns, wrapDel } from './revision-markup.js';
import { reconstruct, isClean } from './reconstruction.js';

/** The slice of `AuthorContext` this module needs (kept local to avoid a cycle). */
export interface RevisionIds {
  author: string;
  date: string;
  nextId: () => number;
}

type Kind = 'ins' | 'del';

// ---------------------------------------------------------------------------
// Revising content that is not a plain run
// ---------------------------------------------------------------------------

/**
 * Elements a `<w:ins>`/`<w:del>` may not wrap, or that Word records a revision
 * INSIDE of. `w:hyperlink` and `w:fldSimple` are paragraph-level content
 * (`EG_PContent`) and are not permitted inside `CT_RunTrackChange` at all; a
 * content control's revision belongs on its content; an equation's on its math
 * runs, which is where Word puts it. Each is kept and its content revised.
 */
const CONTAINERS = new Set([
  'w:hyperlink', 'w:fldSimple', 'w:smartTag', 'w:customXml', 'w:dir', 'w:bdo',
  'w:sdt', 'm:oMath', 'm:oMathPara',
]);

/** A container's own property children, which stay outside any revision. */
const CONTAINER_PROPS = new Set(['w:smartTagPr', 'w:customXmlPr', 'w:sdtPr', 'w:sdtEndPr', 'm:oMathParaPr']);

/** Does this paragraph hold content a whole-paragraph `<w:ins>`/`<w:del>` would mis-wrap? */
export function needsStructuredRevision(paraXml: string): boolean {
  return /<(?:w:hyperlink|w:fldSimple|m:oMath|w:txbxContent)[\s/>]/.test(paraXml);
}

/**
 * Mark every run of a paragraph inserted or deleted, descending into links,
 * simple fields, content controls, equations and text boxes instead of
 * wrapping them. The paragraph MARK is the caller's business (ADR-0035).
 */
export function reviseParagraphRuns(paraXml: string, kind: Kind, ids: RevisionIds): string {
  const para = paragraphNode(paraXml);
  if (para === undefined) return paraXml;
  const children = childrenOf(para);
  const pPr = children.filter((c) => tagOf(c) === 'w:pPr');
  const rest = children.filter((c) => tagOf(c) !== 'w:pPr');
  const [open, close] = shell(para);
  return `${open}${buildNodes(pPr)}${reviseNodes(rest, kind, ids)}${close}`;
}

/** Revise a list of sibling nodes: plain content grouped into one wrapper, containers descended into. */
function reviseNodes(nodes: readonly PoNode[], kind: Kind, ids: RevisionIds): string {
  const out: string[] = [];
  let group: PoNode[] = [];
  const flush = (): void => {
    if (group.length === 0) return;
    const plain = buildNodes(group);
    if (plain.trim() === '') {
      out.push(plain);
      group = [];
      return;
    }
    // The wrapper's id first, then any text box's own paragraphs: ids are
    // minted in the order their elements open.
    const id = ids.nextId();
    const xml = buildNodes(
      group.map((n) => (hasDescendant(n, 'w:txbxContent') ? reviseTextBoxes(n, kind, ids) : n)),
    );
    group = [];
    out.push(kind === 'ins' ? wrapIns(xml, id, ids.author, ids.date) : wrapDel(xml, id, ids.author, ids.date));
  };
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag !== undefined && CONTAINERS.has(tag)) {
      flush();
      out.push(reviseContainer(node, tag, kind, ids));
      continue;
    }
    group.push(node);
  }
  flush();
  return out.join('');
}

function reviseContainer(node: PoNode, tag: string, kind: Kind, ids: RevisionIds): string {
  const [open, close] = shell(node);
  const children = childrenOf(node);
  if (tag === 'w:sdt' || tag === 'm:oMathPara') {
    // The revisable content is one level down: `w:sdtContent`, or each `m:oMath`.
    const inner = children.map((c) => {
      const t = tagOf(c);
      if (t === 'w:sdtContent') {
        const [o, cl] = shell(c);
        return `${o}${reviseNodes(childrenOf(c), kind, ids)}${cl}`;
      }
      if (t === 'm:oMath') return reviseContainer(c, t, kind, ids);
      return buildNodes([c]);
    });
    return `${open}${inner.join('')}${close}`;
  }
  const props = children.filter((c) => CONTAINER_PROPS.has(tagOf(c) ?? ''));
  const rest = children.filter((c) => !CONTAINER_PROPS.has(tagOf(c) ?? ''));
  return `${open}${buildNodes(props)}${reviseNodes(rest, kind, ids)}${close}`;
}

/**
 * A run anchoring a text box: the paragraphs INSIDE the box are revised too, so
 * the box's own words show as inserted or struck and come back on a reject. The
 * run itself is then wrapped with its siblings like any other.
 */
function reviseTextBoxes(node: PoNode, kind: Kind, ids: RevisionIds): PoNode {
  const tag = tagOf(node);
  if (tag === undefined) return node;
  const children = childrenOf(node).map((child) => {
    if (tag === 'w:txbxContent' && tagOf(child) === 'w:p') {
      return parseNodes(reviseParagraphRuns(buildNodes([child]), kind, ids))[0] ?? child;
    }
    return hasDescendant(child, 'w:txbxContent') || tagOf(child) === 'w:txbxContent'
      ? reviseTextBoxes(child, kind, ids)
      : child;
  });
  return { ...node, [tag]: children };
}

function hasDescendant(node: PoNode, tag: string): boolean {
  for (const child of childrenOf(node)) {
    if (tagOf(child) === tag || hasDescendant(child, tag)) return true;
  }
  return false;
}

/** The open and close tags of an element, attributes kept, built by the shared builder. */
function shell(node: PoNode): [string, string] {
  const tag = tagOf(node)!;
  const MARK = 'MUTATISSHELLMARK';
  const built = buildNodes([{ ...node, [tag]: [{ '#text': MARK }] }]);
  const at = built.indexOf(MARK);
  return [built.slice(0, at), built.slice(at + MARK.length)];
}

// ---------------------------------------------------------------------------
// The object-aware replace
// ---------------------------------------------------------------------------

/** One side of the pair: its units, and which unit each body character came from. */
interface SideMap {
  units: InlineUnit[];
  /** Unit index for each character of `normText`, or `FREE`. */
  charUnit: number[];
  /** Units with no character in `normText`, at the body position they sit before. */
  zeroWidth: { unit: number; pos: number; place: number }[];
  /**
   * Units holding a real space credited to nobody in them: collapsed into a
   * neighbour's space, or trimmed at the paragraph's edge. The unit's XML spells
   * more than its share of the text, so emitted verbatim it would carry that
   * space into a view where it is not collapsed or trimmed away.
   */
  shared: Set<number>;
  /**
   * The stripped number as the document spells it: without its separator when
   * that space is only the boundary of an equation or text box right after it.
   */
  prefix: string;
  /** Offset in the paragraph's raw text (`collectText`) of each `normText` character. */
  rawAt: number[];
  /** Offset in the raw text at which each unit starts, and the raw text's length. */
  rawStart: number[];
  rawLength: number;
  /** The unit each raw character belongs to. */
  rawUnit: number[];
  /**
   * The run each raw character was read from (`null` for a boundary space), or
   * `null` altogether when the leaves and the text disagree — the paragraph is
   * then never rebuilt from its source runs.
   */
  leaves: (TextLeaf | null)[] | null;
}

/**
 * `ZeroTextObject.at` / `SideMap.zeroWidth.place` for an object standing in a
 * paragraph's manual number rather than after it: `-2` before the number, `-1`
 * inside it. Not a character count, so a renumbered clause (`9.` → `10.`) does
 * not move it, and not `0`, which is the first character after the number — an
 * object that moved across the number moved.
 */
export function manualNumberPlace(charsBefore: number): number {
  return charsBefore === 0 ? -2 : -1;
}

/** `SideMap.charUnit` for a boundary-only space: no unit owns it. */
const FREE = -1;

/** Does this record carry anything the text rebuild would lose? */
function hasNonText(units: readonly InlineUnit[]): boolean {
  return units.some((u) => u.kind !== 'text');
}

function sideMap(rec: ParaRecord): SideMap | null {
  const units = paragraphUnits(rec.rawXml);
  if (units === null) return null;
  let raw = '';
  const rawUnit: number[] = [];
  const rawStart: number[] = [];
  const boundary: boolean[] = [];
  units.forEach((u, i) => {
    rawStart.push(raw.length);
    raw += u.text;
    for (let k = 0; k < u.text.length; k++) {
      rawUnit.push(i);
      boundary.push(u.boundaries.has(k));
    }
  });
  const { text, offsets } = normaliseTextWithOffsets(raw);
  // Held to the extractor: if the mapped normaliser and the one that built
  // `normText` ever disagree, nothing below can be trusted.
  if (text !== normaliseText(raw)) return null;
  const prefixLen = rec.numberingPrefix.length;
  if (text.slice(0, prefixLen) !== rec.numberingPrefix || text.slice(prefixLen) !== rec.normText) return null;
  // An object inside the stripped number would be emitted twice — once in the
  // restored prefix run, once verbatim. Leave that paragraph to the old path.
  for (let i = 0; i < prefixLen; i++) {
    // A boundary-only space after the number (`4.2` then an equation) is the
    // number's separator and no character of the equation.
    if (boundary[offsets[i]!]) continue;
    if (units[rawUnit[offsets[i]!]!]!.kind !== 'text') return null;
  }
  // A collapsed whitespace run is credited to its FIRST character, and that can
  // be a BOUNDARY space — one `collectText` reads around an equation or text
  // box, which is a character of the paragraph text but not of the document.
  // Credited to it, a real space just after would be authored by nobody: the
  // object is emitted from its own XML, and whatever follows — a text unit, or
  // a link diffed inside its shell — from the character after, so the words
  // fuse (`"r=3%per annum"`, `"x[the Buyer portal]"`). A run holding a real
  // space credits the first real one instead.
  for (let i = 0; i < text.length; i++) {
    const j = offsets[i]!;
    if (text[i] !== ' ' || !boundary[j]) continue;
    for (let k = j + 1; k < raw.length && /\s/.test(raw[k]!); k++) {
      if (!boundary[k]) {
        offsets[i] = k;
        break;
      }
    }
  }
  // A space that is ONLY a boundary — no real space in its run — belongs to no
  // unit. It is how the diff sees an equation or text box as its own word; the
  // document has no character there, so inserting, deleting or keeping it is
  // nobody's edit, and attributing it to the object turned a word fused or
  // split at the object's edge into a change TO the object.
  const charUnit = offsets.slice(prefixLen).map((o) => (boundary[o] ? FREE : rawUnit[o]!));
  const credited = new Set(offsets);
  const shared = new Set<number>();
  for (let a = 0; a < raw.length; ) {
    if (!/\s/.test(raw[a]!)) {
      a++;
      continue;
    }
    let b = a;
    while (b < raw.length && /\s/.test(raw[b]!)) b++;
    const run = [...Array(b - a).keys()].map((k) => a + k);
    const owner = run.find((k) => credited.has(k));
    // A run with no credited character was trimmed at the paragraph's edge —
    // and the other view can put text before it, where it would read as a
    // space. Beside a boundary it reads as one anyway, so that is no share.
    const trimmed = owner === undefined && !run.some((k) => boundary[k]);
    for (const k of run) {
      if (credited.has(k) || boundary[k]) continue;
      if (owner !== undefined ? rawUnit[k] !== rawUnit[owner] : trimmed) shared.add(rawUnit[k]!);
    }
    a = b;
  }
  const bearing = new Set(charUnit);
  const zeroWidth: SideMap['zeroWidth'] = [];
  units.forEach((u, i) => {
    if (u.kind === 'text' || bearing.has(i)) return;
    let pos = 0;
    while (pos < charUnit.length && offsets[prefixLen + pos]! < rawStart[i]!) pos++;
    let before = 0;
    while (before < prefixLen && offsets[before]! < rawStart[i]!) before++;
    zeroWidth.push({ unit: i, pos, place: before < prefixLen ? manualNumberPlace(before) : pos });
  });
  const prefix =
    prefixLen > 0 && boundary[offsets[prefixLen - 1]!] ? rec.numberingPrefix.slice(0, -1) : rec.numberingPrefix;
  let leaves: (TextLeaf | null)[] | null = [];
  for (const u of units) {
    const own = textLeaves(u.nodes);
    if (own.length !== u.text.length || own.some((l, k) => (l === null) !== u.boundaries.has(k))) {
      leaves = null;
      break;
    }
    leaves.push(...own);
  }
  return {
    units, charUnit, zeroWidth, shared, prefix, rawAt: offsets.slice(prefixLen), rawStart, rawLength: raw.length,
    rawUnit, leaves,
  };
}

// ---------------------------------------------------------------------------
// Rebuilding text from its source runs
// ---------------------------------------------------------------------------

/**
 * The source leaves body character `i` of `side` stands for. A `normText`
 * character is a stretch of the document: a collapsed whitespace run is ONE
 * character for several (all of them emitted, tabs and breaks as themselves),
 * and an em dash is FOUR characters for one (emitted once, by the first). A
 * stretch never reaches into another object's characters — those are that
 * object's to emit — but runs on across ordinary text units, since a bookmark
 * between two runs of text splits a whitespace run in two.
 *
 * The whitespace trimmed at the paragraph's edges belongs to no character. The
 * first character carries the leading (a paragraph with no manual number,
 * which owns everything before its body), with whatever revision it has. The
 * last carries the trailing only when `trailing` — when it is revised: an
 * UNCHANGED last character that carried it put a revised paragraph's final
 * space before the original's struck continuation, which brings its own
 * (`"Price.  [-The…-]"`); the caller emits it at the end instead.
 */
function spanLeaves(side: SideMap, i: number, trailing: boolean): TextLeaf[] | null {
  const leaves = side.leaves;
  const at = side.rawAt[i];
  if (leaves === null || at === undefined) return null;
  if (i > 0 && side.rawAt[i - 1] === at) return []; // an em dash's second, third and fourth
  let end = trailing ? side.rawLength : at + 1;
  for (let j = i + 1; j < side.rawAt.length; j++) {
    if (side.rawAt[j]! > at) {
      end = side.rawAt[j]!;
      break;
    }
  }
  return leavesIn(side, i === 0 && side.prefix === '' ? 0 : at, end, at);
}

/** Is body character `i` the last of its side (the one trailing whitespace follows)? */
function isLastOf(side: SideMap, i: number): boolean {
  const at = side.rawAt[i];
  return at !== undefined && at === side.rawAt[side.rawAt.length - 1]; // offsets never decrease
}

/**
 * Leaves of raw offsets `[start, end)` — of `own`'s unit, and of every text
 * unit when `own`'s is one (or when there is no `own`).
 */
function leavesIn(side: SideMap, start: number, end: number, own = -1): TextLeaf[] {
  const unit = own >= 0 ? side.rawUnit[own]! : -1;
  const text = (u: number): boolean => side.units[u]!.kind === 'text';
  const out: TextLeaf[] = [];
  for (let k = start; k < end; k++) {
    const leaf = side.leaves?.[k];
    const u = side.rawUnit[k]!;
    if (leaf === null || leaf === undefined) continue;
    if (k !== own && u !== unit && !(text(u) && (unit < 0 || text(unit)))) continue;
    out.push(leaf);
  }
  return out;
}

/** The whitespace trimmed at the end of a paragraph, from its text runs. */
function trailingLeaves(side: SideMap): TextLeaf[] {
  const last = side.rawAt[side.rawAt.length - 1];
  return last === undefined ? [] : leavesIn(side, last + 1, side.rawLength);
}

/** The source leaves of the manual number the side's body is stripped of (all before the body). */
function prefixLeaves(side: SideMap): TextLeaf[] | null {
  if (side.leaves === null) return null;
  return side.prefix === '' ? [] : leavesIn(side, 0, side.rawAt[0] ?? side.rawLength);
}

/**
 * Runs rebuilt from source leaves: one `<w:r>` per stretch of one source run,
 * carrying that run's properties and attributes, its characters as they are in
 * the document — a curly quote, a non-breaking space, a tab or break as the
 * element it was. A run split at a diff boundary keeps its properties on both
 * pieces.
 */
function leafRuns(leaves: readonly TextLeaf[]): string {
  const out: PoNode[] = [];
  /** A whole source run, its formatting revision dropped. */
  const verbatim = (run: PoNode): PoNode => {
    const tag = tagOf(run)!;
    const copy: PoNode = { [tag]: childrenOf(run).map((c) => (tagOf(c) === 'w:rPr' ? withoutFormattingRevision(c) : c)) };
    if (run[':@'] !== undefined) copy[':@'] = run[':@'];
    return copy;
  };
  for (let i = 0; i < leaves.length; ) {
    out.push(...leaves[i]!.runsBefore.map(verbatim));
    const run = leaves[i]!.run;
    const tag = tagOf(run)!;
    // A formatting revision an input document carried (`w:rPrChange`) is not
    // copied: its `w:id` would repeat in the redline (§5.5: ids unique) and the
    // other party's revision would stay live in it. The run keeps its formatting.
    const props = childrenOf(run)
      .filter((c) => tagOf(c) === 'w:rPr' || tagOf(c) === 'm:rPr')
      .map((c) => withoutFormattingRevision(c));
    const content: PoNode[] = [];
    let text = '';
    const flush = (): void => {
      if (text === '') return;
      // `xml:space="preserve"` where an edge space needs it, as the text rebuild writes it.
      const leaf: PoNode = { [tag === 'm:r' ? 'm:t' : 'w:t']: [{ '#text': text }] };
      if (/^\s|\s$/.test(text)) leaf[':@'] = { '@_xml:space': 'preserve' };
      content.push(leaf);
      text = '';
    };
    let after: PoNode[] = [];
    for (; i < leaves.length && leaves[i]!.run === run; i++) {
      const leaf = leaves[i]!;
      after = leaf.runsAfter;
      if (leaf.before.length > 0) {
        flush();
        content.push(...leaf.before);
      }
      if (leaf.item === null) {
        text += leaf.char;
      } else {
        flush();
        content.push(leaf.item);
      }
      if (leaf.after.length > 0) {
        flush();
        content.push(...leaf.after);
      }
    }
    flush();
    const attrs = run[':@'];
    out.push(attrs === undefined ? { [tag]: [...props, ...content] } : { [tag]: [...props, ...content], ':@': attrs });
    out.push(...after.map(verbatim));
  }
  return buildNodes(out);
}

/** `rPr` without its `w:rPrChange` child, its attributes kept. */
function withoutFormattingRevision(rPr: PoNode): PoNode {
  const tag = tagOf(rPr)!;
  const kids = childrenOf(rPr);
  if (!kids.some((c) => tagOf(c) === 'w:rPrChange')) return rPr;
  const out: PoNode = { [tag]: kids.filter((c) => tagOf(c) !== 'w:rPrChange') };
  if (rPr[':@'] !== undefined) out[':@'] = rPr[':@'];
  return out;
}

/** `xml` wrapped as `kind`, or as it is for unchanged text. */
function revise(xml: string, kind: WordDiff['op'], ids: RevisionIds): string {
  if (xml === '' || kind === 'eq') return xml;
  return kind === 'ins' ? wrapIns(xml, ids.nextId(), ids.author, ids.date) : wrapDel(xml, ids.nextId(), ids.author, ids.date);
}

/**
 * The manual number of an edited clause, from its source runs (ADR-0028). The same number on both sides is the revised one,
 * unmarked. A number that changed is the revised one inserted and the original
 * struck, insertion first (ADR-0043) — otherwise rejecting every change would
 * keep the new number on the old clause. `null` when a number has no leaves.
 */
function numberFromSource(o: SideMap, orig: ParaRecord, r: SideMap, rev: ParaRecord, ids: RevisionIds): string | null {
  const rl = prefixLeaves(r);
  const ol = prefixLeaves(o);
  if (rl === null || ol === null) return null;
  return markedNumber(rl, ol, orig.numberingPrefix === rev.numberingPrefix, ids);
}

/** The rule of {@link numberFromSource}, on the two numbers' leaves. */
function markedNumber(revised: readonly TextLeaf[], original: readonly TextLeaf[], same: boolean, ids: RevisionIds): string {
  if (same) return leafRuns(revised);
  return revise(leafRuns(revised), 'ins', ids) + revise(leafRuns(original), 'del', ids);
}

/**
 * Does the authored paragraph reproduce both documents' paragraphs as a reader
 * sees them — characters and run formatting (`reconstruction.ts`, the same
 * comparison the test yardstick makes)? The rebuild from source runs is held
 * to this; the text rebuild it replaces could not be.
 */
function reconstructsExactly(paraXml: string, orig: ParaRecord, rev: ParaRecord): boolean {
  return reconstructsRange(paraXml, [orig], [rev]);
}

function reconstructsRange(xml: string, origRecs: readonly ParaRecord[], revRecs: readonly ParaRecord[]): boolean {
  const body = (recs: readonly ParaRecord[]): string => recs.map((rec) => rec.rawXml).join('');
  const report = reconstruct(xml, body(origRecs), body(revRecs));
  return isClean(report.accept) && isClean(report.reject);
}

/** One paragraph of a merge or split, as {@link authorRangeFromSource} hands it to the caller. */
export interface SeamParagraph {
  /** 0-based position among the authored paragraphs. */
  k: number;
  /** Whether it is the last (it takes the revised record's properties). */
  last: boolean;
  /** Its manual number's XML, revised as its place in the join requires. */
  number: string;
  runs: string;
  /** The revision on the paragraph mark that closes it, if any. */
  mark?: 'ins' | 'del';
}

/**
 * A merge or split (ADR-0032) rebuilt from its source runs, exactly as a
 * one-to-one replace is: the word diff spans several records on one
 * side, joined by {@link PARA_SEAM}, and each character is traced to its record
 * and its run. The paragraphs are cut at the seams; `paragraph` wraps each
 * (properties, mark). The first paragraph's number follows ADR-0028 as amended
 * by task 4 (the same number unmarked, a changed one inserted and struck); a
 * later paragraph's number rides the revision of the mark before it, as the
 * text rebuild's does. `paragraph` takes its ids from the `ids` it is handed —
 * a scratch counter on the dry run.
 *
 * `null` — a record holding anything but text (its objects would be lost, as
 * they are in the text rebuild, ADR-0045 limit 3, which the caller then takes),
 * a character with no source, or a result that does not reconstruct both
 * ranges as a reader sees them.
 */
export function authorRangeFromSource(
  wordDiffs: readonly WordDiff[],
  origRecs: readonly ParaRecord[],
  revRecs: readonly ParaRecord[],
  ids: RevisionIds,
  seam: string,
  paragraph: (p: SeamParagraph, ids: RevisionIds) => string,
): string | null {
  const o = rangeText(origRecs);
  const r = rangeText(revRecs);
  if (o === null || r === null) return null;
  const many = origRecs.length > revRecs.length ? o : r;
  const build = (revIds: RevisionIds): string | null => {
    const groups: { pieces: { op: WordDiff['op']; leaves: TextLeaf[] }[]; mark?: 'ins' | 'del' }[] = [{ pieces: [] }];
    let jr = 0;
    let jo = 0;
    for (const wd of wordDiffs) {
      for (const c of wd.text) {
        const group = groups[groups.length - 1]!;
        if (c === seam) {
          if (wd.op !== 'eq') group.mark = wd.op;
          groups.push({ pieces: [] });
        } else {
          const leaves = wd.op === 'del' ? o.leaves(jo) : r.leaves(jr);
          if (leaves === null) return null;
          const last = group.pieces[group.pieces.length - 1];
          if (last !== undefined && last.op === wd.op) last.leaves.push(...leaves);
          else group.pieces.push({ op: wd.op, leaves: [...leaves] });
        }
        if (wd.op !== 'del') jr++;
        if (wd.op !== 'ins') jo++;
      }
    }
    const out: string[] = [];
    for (const [k, group] of groups.entries()) {
      let number: string | null;
      if (k === 0) {
        const rn = r.number(0);
        const on = o.number(0);
        if (rn === null || on === null) return null;
        number = markedNumber(rn, on, origRecs[0]?.numberingPrefix === revRecs[0]?.numberingPrefix, revIds);
      } else {
        const n = many.number(k);
        if (n === null) return null;
        number = revise(leafRuns(n), groups[k - 1]!.mark ?? 'eq', revIds);
      }
      const runs = group.pieces.map((piece) => revise(leafRuns(piece.leaves), piece.op, revIds)).join('');
      out.push(paragraph({ k, last: k === groups.length - 1, number, runs, ...(group.mark ? { mark: group.mark } : {}) }, revIds));
    }
    return out.join('');
  };
  let scratch = 1;
  const dry = build({ ...ids, nextId: () => scratch++ });
  if (dry === null || !reconstructsRange(dry, origRecs, revRecs)) return null;
  return build(ids);
}

/** A record range's characters, joined as the word diff joined them, each traced to its source. */
interface RangeText {
  /** Leaves of joined character `j` — `null` for a seam or a character with no source. */
  leaves(j: number): TextLeaf[] | null;
  /** The manual number of the range's `k`-th record, from its runs. */
  number(k: number): TextLeaf[] | null;
}

function rangeText(records: readonly ParaRecord[]): RangeText | null {
  const maps: SideMap[] = [];
  for (const rec of records) {
    const m = sideMap(rec);
    if (m === null || m.leaves === null || m.units.some((u) => u.kind === 'object')) return null;
    maps.push(m);
  }
  const at: ([number, number] | null)[] = [];
  records.forEach((rec, k) => {
    if (k > 0) at.push(null);
    for (let i = 0; i < rec.normText.length; i++) at.push([k, i]);
  });
  return {
    leaves: (j) => {
      const e = at[j];
      return e === null || e === undefined ? null : spanLeaves(maps[e[0]]!, e[1], true);
    },
    number: (k) => (maps[k] === undefined ? null : prefixLeaves(maps[k]!)),
  };
}

/** How {@link authorObjectAwareRuns} authors ordinary text. */
export interface TextAuthoring {
  /**
   * `true`: every character from its own source run — a kept or
   * inserted one from the revised paragraph's, a struck one from the
   * original's — formatting, typography, tabs and breaks kept, the manual
   * number too; the result held to `reconstruction.ts` as well as to the
   * text-and-census check. Taken by EVERY one-to-one pair, objects or not.
   *
   * `false` (default): the caller's `emitText` over the word diff's
   * normalised text — the pre-08i rebuild, objects kept (ADR-0045); a pair
   * with no objects returns `null`, since the caller's own text rebuild is
   * then the same thing.
   */
  source?: boolean;
  /**
   * The manual number for the text rebuild, as XML: the revised and original
   * spellings, whether the two numbers are the same (`numberingPrefix`), and
   * the ids any revision on it takes. Unused when `source` (the number is
   * rebuilt from its runs).
   */
  number?: (revised: string, original: string, same: boolean, ids: RevisionIds) => string;
}

/**
 * The runs of a changed one-to-one paragraph, objects kept, or `null` when the
 * result fails its own check — or, rebuilding from text, when the pair has no
 * objects (the caller's text rebuild is then exact).
 *
 * `emitText` is the caller's text-run emitter, so ordinary text rebuilt from
 * the diff is authored by the one function that always authored it.
 * `prefix` is the manual number's XML, its revision ids minted first, so ids
 * still open in document order.
 */
export function authorObjectAwareRuns(
  wordDiffs: readonly WordDiff[],
  orig: ParaRecord,
  rev: ParaRecord,
  ids: RevisionIds,
  emitText: (wd: WordDiff, ids: RevisionIds) => string,
  wrapParagraph: (runs: string, prefix: string) => string,
  authoring: TextAuthoring = {},
): { runs: string; prefix: string } | null {
  const o = sideMap(orig);
  const r = sideMap(rev);
  if (o === null || r === null) return null;
  const source = authoring.source === true;
  if (source && (o.leaves === null || r.leaves === null)) return null;
  if (!source && !hasNonText(o.units) && !hasNonText(r.units)) return null;

  // The stream must spell both sides, or a character cannot be placed.
  let revText = '';
  let origText = '';
  for (const wd of wordDiffs) {
    if (wd.op !== 'del') revText += wd.text;
    if (wd.op !== 'ins') origText += wd.text;
  }
  if (revText !== rev.normText || origText !== orig.normText) return null;

  // A split changes what is unchanged, so the pairing can move; each pass
  // splits at least one `eq` character, so this ends.
  let stream: readonly WordDiff[] = wordDiffs;
  for (let next = splitStraySpaces(stream, o, r); next !== null; next = splitStraySpaces(stream, o, r)) stream = next;

  const number = (revIds: RevisionIds): string | null =>
    source
      ? numberFromSource(o, orig, r, rev, revIds)
      : (authoring.number?.(r.prefix, o.prefix, orig.numberingPrefix === rev.numberingPrefix, revIds) ?? '');
  /** The number first, then the runs: ids are minted in document order. */
  const author = (revIds: RevisionIds): { runs: string; prefix: string } | null => {
    const prefix = number(revIds);
    if (prefix === null) return null;
    const runs = walk(stream, o, r, revIds, emitText, source);
    return runs === null ? null : { runs, prefix };
  };
  // Dry run on a scratch counter, so a result that fails its check mints no id.
  let scratch = 1;
  const dry = author({ ...ids, nextId: () => scratch++ });
  if (dry === null) return null;
  const para = wrapParagraph(dry.runs, dry.prefix);
  if (!readsAs(para, orig, rev)) return null;
  if (source && !reconstructsExactly(para, orig, rev)) return null;
  return author(ids);
}

/**
 * A text-identical pair whose zero-text objects differ — a footnote or endnote
 * reference, a picture, an embedded object added, removed or moved, and not a
 * word changed (ADR-0045 § Amendment). The REVISED paragraph is emitted as it
 * is, every run with its own formatting; only the objects that differ are
 * marked: a revised one with no counterpart is inserted where it sits, an
 * original one with no counterpart is struck at its own place, a run of text
 * split there if need be.
 *
 * Not the replace author: that rebuilds every word from the word diff (limit 5),
 * which would take the formatting off a clause whose words did not change.
 *
 * Objects are aligned on sort and place (longest common subsequence), so one
 * footnote removed of three marks that one alone. `null` — no unit model, an
 * object that holds a space, a place inside a link or equation, or a result
 * that fails its check (both views' text, census AND the places of their
 * zero-text objects) — and the caller falls back.
 */
export function authorObjectChange(
  orig: ParaRecord,
  rev: ParaRecord,
  ids: RevisionIds,
  wrapParagraph: (runs: string) => string,
): string | null {
  const o = sideMap(orig);
  const r = sideMap(rev);
  if (o === null || r === null || orig.normText !== rev.normText) return null;
  const oz = o.zeroWidth.filter((z) => isObject(o, z.unit));
  const rz = r.zeroWidth.filter((z) => isObject(r, z.unit));
  if (oz.some((z) => o.shared.has(z.unit)) || rz.some((z) => r.shared.has(z.unit))) return null;
  const key = (side: SideMap, z: { unit: number; place: number }): string =>
    `${side.units[z.unit]!.objectKind}:${census(buildNodes(side.units[z.unit]!.nodes), false)}@${z.place}`;
  const kept = computeLCS(oz.map((z) => key(o, z)), rz.map((z) => key(r, z)), (k) => k);
  /** Kept revised unit → its original counterpart. */
  const counterpart = new Map(kept.map(([i, j]) => [rz[j]!.unit, oz[i]!.unit]));
  const keptOrig = new Set(kept.map(([i]) => i));
  const struck = oz.filter((_, i) => !keptOrig.has(i));

  // Where each struck original goes in the revised paragraph's raw text: before
  // the character at its place, or at the end.
  //
  // A place on a space inside an object with text — the boundary read at a text
  // box's or equation's edge, or a real space a collapsed run was credited to —
  // with only spaces between it and the object's edge, is that edge: the same
  // place once whitespace is collapsed, and one a struck object can sit at.
  const target = (z: { pos: number; place: number }): number => {
    if (z.place === -2) return 0; // before the manual number
    const at = r.rawAt[z.pos];
    if (at === undefined) return r.rawLength;
    const k = r.units.findIndex((u, i) => u.kind !== 'text' && r.rawStart[i]! < at && at < r.rawStart[i]! + u.text.length);
    if (k < 0) return at;
    const start = r.rawStart[k]!;
    const text = r.units[k]!.text;
    if (/^\s*$/.test(text.slice(0, at - start))) return start;
    if (/^\s*$/.test(text.slice(at - start))) return start + text.length;
    return at;
  };
  for (const z of struck) {
    if (z.place === -1) return null; // inside the manual number: no place to strike it
    const at = target(z);
    // Strictly inside an object that has text — a link, a field, an equation:
    // the struck one cannot go there without opening it.
    const inside = r.units.findIndex((u, k) => u.kind !== 'text' && r.rawStart[k]! < at && at < r.rawStart[k]! + u.text.length);
    if (inside >= 0) return null;
  }

  const build = (revIds: RevisionIds): string => {
    const out: string[] = [];
    let next = 0;
    const strike = (z: { unit: number }): void => {
      out.push(reviseNodes(o.units[z.unit]!.nodes, 'del', revIds));
    };
    /** Struck originals that belong at raw offset `at` and before original unit `before`. */
    const flush = (at: number, before = Infinity): void => {
      while (next < struck.length && target(struck[next]!) <= at && struck[next]!.unit < before) strike(struck[next++]!);
    };
    r.units.forEach((u, k) => {
      const start = r.rawStart[k]!;
      if (u.kind === 'text') {
        // Struck originals whose place falls inside this run of text split it.
        const cuts: number[] = [];
        for (let n = next; n < struck.length && target(struck[n]!) < start + u.text.length; n++) {
          const c = target(struck[n]!) - start;
          if (c > 0 && !cuts.includes(c)) cuts.push(c);
        }
        flush(start);
        splitText(u.nodes, cuts).forEach((piece, i) => {
          out.push(buildNodes(piece));
          if (i < cuts.length) flush(start + cuts[i]!);
        });
        return;
      }
      const mate = counterpart.get(k);
      if (mate !== undefined) {
        // A kept object: the struck originals that came before its counterpart
        // come before it, or rejecting would swap two objects in one place.
        flush(start, mate);
        out.push(buildNodes(u.nodes));
      } else if (rz.some((z) => z.unit === k)) {
        // An inserted one goes first among the changes at its place (ADR-0043).
        flush(start - 1);
        out.push(reviseNodes(u.nodes, 'ins', revIds));
      } else if (u.kind === 'marker') {
        // A bookmark is never revised and has no width: a struck object at its
        // offset waits past it, so an insertion just after the bookmark still
        // reads before the deletion it sits beside (ADR-0043).
        flush(start - 1);
        out.push(buildNodes(u.nodes));
      } else {
        // An object with text, whose first character is the place of the
        // struck ones at its offset.
        flush(start);
        out.push(buildNodes(u.nodes));
      }
    });
    while (next < struck.length) strike(struck[next++]!);
    return wrapParagraph(out.join(''));
  };

  let scratch = 1;
  const dry = build({ ...ids, nextId: () => scratch++ });
  if (!readsAs(dry, orig, rev) || !placesAs(dry, orig, rev)) return null;
  return build(ids);
}

/**
 * The zero-text objects of each view sit where they sit in the two documents.
 * Exported for the caller's second try, which the replace author's own check
 * (text and census) cannot hold to this: a census does not see a place.
 */
export function placesAs(paraXml: string, orig: ParaRecord, rev: ParaRecord): boolean {
  const same = (a: readonly ZeroTextObject[], b: readonly ZeroTextObject[]): boolean =>
    a.length === b.length && a.every((x, i) => x.sort === b[i]!.sort && x.at === b[i]!.at && x.within === b[i]!.within);
  const view = (policy: 'acceptAll' | 'rejectAll', rec: ParaRecord): readonly ZeroTextObject[] =>
    zeroTextObjects(flattenTrackedChanges(paraXml, policy), rec.numberingPrefix);
  return same(view('acceptAll', rev), zeroTextObjects(rev.rawXml, rev.numberingPrefix)) &&
    same(view('rejectAll', orig), zeroTextObjects(orig.rawXml, orig.numberingPrefix));
}

/**
 * Cut a text unit's nodes at character offsets of its text (ascending, each
 * strictly inside it), splitting a run — and the `<w:t>` in it — where a cut
 * falls inside one. Each piece keeps the run's properties. Returns
 * `cuts.length + 1` pieces.
 */
function splitText(nodes: readonly PoNode[], cuts: readonly number[]): PoNode[][] {
  const pieces: PoNode[][] = [[]];
  let at = 0;
  let c = 0;
  const lengthOf = (n: readonly PoNode[]): number => {
    const parts: string[] = [];
    collectText(n, parts);
    return parts.join('').length;
  };
  for (const node of nodes) {
    const length = lengthOf([node]);
    if (c >= cuts.length || at + length <= cuts[c]! || tagOf(node) !== 'w:r') {
      // Wholly before the next cut (or no cut left): it goes in the current piece.
      if (c < cuts.length && at + length === cuts[c]) {
        pieces[pieces.length - 1]!.push(node);
        pieces.push([]);
        c++;
      } else {
        pieces[pieces.length - 1]!.push(node);
      }
      at += length;
      continue;
    }
    // A run a cut falls inside: rebuild it child by child.
    const attrs = node[':@'];
    const props = childrenOf(node).filter((ch) => tagOf(ch) === 'w:rPr');
    const runOf = (children: PoNode[]): PoNode => ({ 'w:r': [...props, ...children], ...(attrs ? { ':@': attrs } : {}) });
    let children: PoNode[] = [];
    for (const child of childrenOf(node)) {
      if (tagOf(child) === 'w:rPr') continue;
      let rest: PoNode | null = child;
      while (rest !== null) {
        const length = lengthOf([rest]);
        if (c >= cuts.length || at + length < cuts[c]! || (at + length === cuts[c]! && length === 0)) {
          children.push(rest);
          at += length;
          rest = null;
        } else if (at + length === cuts[c]!) {
          children.push(rest);
          at += length;
          rest = null;
          pieces[pieces.length - 1]!.push(runOf(children));
          children = [];
          pieces.push([]);
          c++;
        } else if (tagOf(rest) === 'w:t') {
          const text = childrenOf(rest).map((t) => (typeof t['#text'] === 'string' ? t['#text'] : '')).join('');
          const k = cuts[c]! - at;
          children.push(textLeaf(text.slice(0, k)));
          at += k;
          pieces[pieces.length - 1]!.push(runOf(children));
          children = [];
          pieces.push([]);
          c++;
          rest = textLeaf(text.slice(k));
        } else {
          // A one-character child (a tab or break) the cut cannot fall inside.
          children.push(rest);
          at += length;
          rest = null;
        }
      }
    }
    if (children.length > 0) pieces[pieces.length - 1]!.push(runOf(children));
  }
  while (pieces.length < cuts.length + 1) pieces.push([]);
  return pieces;
}

function textLeaf(text: string): PoNode {
  return { 'w:t': [{ '#text': text }], ':@': { '@_xml:space': 'preserve' } };
}

/**
 * Accepting every revision gives the revised paragraph; rejecting every one
 * gives the original — compared two ways, because each catches what the other
 * cannot:
 *
 * - the TEXT, as the diff itself reads it: `normText`, boundary spaces around
 *   an equation or text box included. Not the document's own spelling, though
 *   that was tried — a check stricter than the diff it checks rejects what the
 *   diff calls equal (a space added or removed right beside an equation, which
 *   the boundary space absorbs), and every rejection is a paragraph sent back
 *   to the lossy rebuild. A space LOST beside an equation by this author is
 *   prevented where it would arise (`sideMap`'s re-credit) and is pinned by the
 *   tests, which do compare the document's spelling;
 * - a CENSUS of the objects that carry no text or whose text a rebuild would
 *   keep: footnote and endnote references, links, fields, equations,
 *   drawings, embedded objects, content controls. A footnote dropped from inside a rebuilt link, or a link emitted
 *   twice, leaves the text intact and the census wrong.
 *
 * Bookmarks are counted on the accepted view only: they come from the revised
 * paragraph and are never revised, so the rejected view carries the revised
 * side's bookmarks by design.
 */
function readsAs(paraXml: string, orig: ParaRecord, rev: ParaRecord): boolean {
  const accepted = flattenTrackedChanges(paraXml, 'acceptAll');
  const rejected = flattenTrackedChanges(paraXml, 'rejectAll');
  const body = (xml: string): string | null => {
    const para = paragraphNode(xml);
    return para === undefined ? null : stripNumberingPrefix(normaliseText(paragraphText(para))).body;
  };
  if (body(accepted) !== rev.normText || body(rejected) !== orig.normText) return false;
  return census(accepted, true) === census(rev.rawXml, true) && census(rejected, false) === census(orig.rawXml, false);
}

const CENSUS: readonly RegExp[] = [
  /<w:footnoteReference[\s/>]/g,
  /<w:endnoteReference[\s/>]/g,
  /w:fldCharType="begin"/g,
  ...['w:hyperlink', 'w:fldSimple', 'm:oMath', 'w:drawing', 'w:pict', 'w:object', 'w:txbxContent', 'w:sdtContent'].map(
    (tag) => new RegExp(`<${tag}[\\s/>]`, 'g'),
  ),
];

function census(xml: string, bookmarks: boolean): string {
  const pruned = buildNodes(pruneEmptyContainers(parseNodes(xml)));
  const counts = CENSUS.map((re) => (pruned.match(re) ?? []).length);
  if (bookmarks) counts.push((pruned.match(/<w:bookmarkStart[\s/>]/g) ?? []).length);
  return counts.join(',');
}

/**
 * Drop every container left with nothing but its own properties — the shell a
 * whole-object deletion leaves once accepted (`<w:hyperlink …/>`, an `m:oMath`
 * with no runs, a content control with an empty `w:sdtContent`). Revisions go
 * INSIDE these elements (`CONTAINERS`), so flattening removes the content and
 * keeps the shell; counting the shell would read one object as two.
 */
function pruneEmptyContainers(nodes: readonly PoNode[]): PoNode[] {
  const out: PoNode[] = [];
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) {
      out.push(node);
      continue;
    }
    const children = pruneEmptyContainers(childrenOf(node));
    const empty = children.every((c) => CONTAINER_PROPS.has(tagOf(c) ?? '') || tagOf(c) === undefined);
    if ((CONTAINERS.has(tag) || tag === 'w:sdtContent') && empty) continue;
    out.push(Array.isArray(node[tag]) ? { ...node, [tag]: children } : node);
  }
  return out;
}

/**
 * Objects whose words can be diffed INSIDE them: a container of ordinary runs.
 * An equation (structure, not a run of words), a footnote reference and a
 * drawing are replaced whole instead — and so is a container holding anything
 * that is NOT an ordinary run (a footnote inside a link, a PAGEREF inside a TOC
 * entry's link, an equation inside a content control), because rebuilding its
 * words from the diff would drop that content (see `insideShell`).
 */
const DIFFABLE_INSIDE = new Set([
  'w:hyperlink', 'w:fldSimple', 'w:smartTag', 'w:customXml', 'w:dir', 'w:bdo', 'w:sdt', 'field',
]);

/**
 * An object's own open and close, for diffing text inside it: a container's
 * tags and properties, a content control down to `w:sdtContent`, a complex
 * field's `begin…separate` and its `end`. Plus the run properties of its first
 * text run, so a rebuilt link still LOOKS like a link.
 *
 * `null` when the content between open and close is anything but ordinary runs
 * — the part that is rebuilt from the word diff must hold nothing else — or the
 * shape is not one this can open: a nested field, a field with no result.
 */
function insideShell(unit: InlineUnit): { open: string; close: string; rPr: string } | null {
  const rPr = firstTextRunProps(unit.nodes);
  if (unit.objectKind === 'field') {
    let depth = 0;
    let sep = -1;
    for (const [i, node] of unit.nodes.entries()) {
      for (const child of childrenOf(node)) {
        if (tagOf(child) !== 'w:fldChar') continue;
        const type = attrOf(child, '@_w:fldCharType');
        if (type === 'begin' && ++depth > 1) return null;
        if (type === 'separate' && depth === 1) sep = i;
      }
    }
    const last = unit.nodes.length - 1;
    if (sep < 0 || sep >= last) return null;
    if (!unit.nodes.slice(sep + 1, last).every(isPlainTextNode)) return null;
    return { open: buildNodes(unit.nodes.slice(0, sep + 1)), close: buildNodes(unit.nodes.slice(last)), rPr };
  }
  const node = unit.nodes[0]!;
  const [open, close] = shell(node);
  const children = childrenOf(node);
  if (unit.objectKind === 'w:sdt') {
    const content = children.find((c) => tagOf(c) === 'w:sdtContent');
    if (content === undefined || !childrenOf(content).every(isPlainTextNode)) return null;
    const [co, cc] = shell(content);
    const props = children.filter((c) => tagOf(c) !== 'w:sdtContent');
    return { open: `${open}${buildNodes(props)}${co}`, close: `${cc}${close}`, rPr };
  }
  const props = children.filter((c) => CONTAINER_PROPS.has(tagOf(c) ?? ''));
  if (!children.filter((c) => !CONTAINER_PROPS.has(tagOf(c) ?? '')).every(isPlainTextNode)) return null;
  return { open: `${open}${buildNodes(props)}`, close, rPr };
}

/** The serialised `<w:rPr>` of the first run under `nodes` that carries text, or ''. */
function firstTextRunProps(nodes: readonly PoNode[]): string {
  for (const node of nodes) {
    if (tagOf(node) === 'w:r' && childrenOf(node).some((c) => tagOf(c) === 'w:t')) {
      return buildNodes(childrenOf(node).filter((c) => tagOf(c) === 'w:rPr'));
    }
    const found = firstTextRunProps(childrenOf(node));
    if (found !== '') return found;
  }
  return '';
}

const isObject = (side: SideMap, u: number): boolean => u >= 0 && side.units[u]!.kind === 'object';

interface Pairing {
  /** Revised unit → its original counterpart, emitted verbatim. */
  verbatim: Map<number, number>;
  /** Revised unit → its counterpart and the shell the word diff runs inside. */
  inside: Map<number, { uo: number; shell: NonNullable<ReturnType<typeof insideShell>> }>;
  /** Original unit → its revised counterpart, for `inside`. */
  insideOrig: Map<number, number>;
}

/** Which objects correspond across the two sides, and how each pair is authored. */
function pairObjects(wordDiffs: readonly WordDiff[], o: SideMap, r: SideMap): Pairing {

  // Which units each object's `eq` characters meet on the other side, and which
  // objects the diff touched at all.
  const revMeets = new Map<number, Set<number>>();
  const origMeets = new Map<number, Set<number>>();
  const changedRev = new Set<number>();
  const changedOrig = new Set<number>();
  {
    let ro = 0;
    let rr = 0;
    for (const wd of wordDiffs) {
      for (let k = 0; k < wd.text.length; k++) {
        if (wd.op === 'eq') {
          const ur = r.charUnit[rr++]!;
          const uo = o.charUnit[ro++]!;
          // A space at an object's edge may sit inside it on one side and
          // outside on the other (`[the portal ]now` → `[the portal] now`) with
          // nothing visible changed. It does not decide what lines up.
          if (wd.text[k] === ' ' || ur === FREE || uo === FREE) continue;
          (revMeets.get(ur) ?? revMeets.set(ur, new Set()).get(ur)!).add(uo);
          (origMeets.get(uo) ?? origMeets.set(uo, new Set()).get(uo)!).add(ur);
        } else if (wd.op === 'ins') {
          const ur = r.charUnit[rr++]!;
          if (ur !== FREE) changedRev.add(ur);
        } else {
          const uo = o.charUnit[ro++]!;
          if (uo !== FREE) changedOrig.add(uo);
        }
      }
    }
  }
  // Two objects CORRESPOND when they are the same sort and each one's `eq`
  // characters meet the other and nothing else. Unchanged → verbatim; changed
  // and diffable inside → the revised shell around a word diff; anything else →
  // each object whole, as one insertion and one deletion.
  const verbatim = new Map<number, number>();
  const inside = new Map<number, { uo: number; shell: NonNullable<ReturnType<typeof insideShell>> }>();
  const insideOrig = new Map<number, number>(); // orig unit → its rev counterpart
  for (const [ur, meets] of revMeets) {
    if (!isObject(r, ur) || meets.size !== 1) continue;
    const uo = [...meets][0]!;
    if (!isObject(o, uo) || origMeets.get(uo)?.size !== 1) continue;
    if (o.units[uo]!.objectKind !== r.units[ur]!.objectKind) continue;
    // Equal text is not enough for verbatim: a link that gained or lost a
    // footnote inside it has the same words and a different census, and taking
    // the revised one unmarked would drop the change (and fail the check).
    const sameInside =
      census(buildNodes(r.units[ur]!.nodes), false) === census(buildNodes(o.units[uo]!.nodes), false);
    // And neither may share a space with a neighbour: emitted verbatim, its XML
    // would carry that space into both views, beside the neighbour's own.
    const ownSpaces = !r.shared.has(ur) && !o.shared.has(uo);
    if (!changedRev.has(ur) && !changedOrig.has(uo) && sameInside && ownSpaces) {
      verbatim.set(ur, uo);
      continue;
    }
    if (!DIFFABLE_INSIDE.has(r.units[ur]!.objectKind)) continue;
    const sh = insideShell(r.units[ur]!);
    // The ORIGINAL's content is rebuilt too (its struck words), so it must be
    // plain as well, or a footnote only the original had would vanish unmarked.
    if (sh === null || insideShell(o.units[uo]!) === null) continue;
    inside.set(ur, { uo, shell: sh });
    insideOrig.set(uo, ur);
  }
  // A pair is diffed inside only if its characters form ONE stretch of the
  // stream. A plain-text edit falling between them (`Pay the [Seller portal]` →
  // `Pay [Buyer portal]` strikes "the " between "Buyer" and "Seller") would
  // otherwise close the shell and open it again — the link, or the whole REF
  // field, emitted twice. Such a pair is replaced whole instead — and that can
  // break another: a character shared by a dropped pair's revised side and a
  // kept pair's original belonged to the dropped one and now belongs to the
  // other, wherever it sits. So until nothing more breaks.
  for (let again = true; again; ) {
    const broken = new Set<number>();
    const seen = new Set<number>();
    let last = -1;
    let qo = 0;
    let qr = 0;
    for (const wd of wordDiffs) {
      for (let k = 0; k < wd.text.length; k++) {
        const ur = wd.op === 'del' ? -1 : r.charUnit[qr++]!;
        const uo = wd.op === 'ins' ? -1 : o.charUnit[qo++]!;
        const owner = inside.has(ur) ? ur : (insideOrig.get(uo) ?? -1);
        if (owner === last) continue;
        if (owner >= 0 && seen.has(owner)) broken.add(owner);
        if (owner >= 0) seen.add(owner);
        last = owner;
      }
    }
    for (const ur of broken) {
      insideOrig.delete(inside.get(ur)!.uo);
      inside.delete(ur);
    }
    again = broken.size > 0;
  }

  return { verbatim, inside, insideOrig };
}

/**
 * An `eq` space the walk cannot author from either object's XML. The meets
 * pass lets an object's edge space line up with a character outside it
 * (`[the portal ]now` → `[the portal] now`), which is right only when that
 * character sits next to the object in the stream: the kept object's own space
 * then stands where the other side's did. Lined up across a deletion or
 * insertion, or with a character of another object that does not correspond,
 * the object's XML carries the space to its own edge — a space shown twice in
 * one view, or moved across the words between (`(the "Portal")` beside a link
 * opening with the space). Such a space is split into the deletion of the one
 * and the insertion of the other, and the pairing is worked out again. `null`
 * when there is none.
 */
function splitStraySpaces(wordDiffs: readonly WordDiff[], o: SideMap, r: SideMap): WordDiff[] | null {
  const { verbatim, inside, insideOrig } = pairObjects(wordDiffs, o, r);
  const revCounterpart = (ur: number): number | undefined => verbatim.get(ur) ?? inside.get(ur)?.uo;
  const origCounterpart = new Map<number, number>(insideOrig);
  for (const [ur, uo] of verbatim) origCounterpart.set(uo, ur);

  const chars: { op: WordDiff['op']; c: string; ur: number; uo: number }[] = [];
  let po = 0;
  let pr = 0;
  for (const wd of wordDiffs) {
    for (const c of wd.text) {
      const ur = wd.op === 'del' ? FREE : r.charUnit[pr++]!;
      const uo = wd.op === 'ins' ? FREE : o.charUnit[po++]!;
      chars.push({ op: wd.op, c, ur, uo });
    }
  }
  const nextTo = (i: number, side: 'rev' | 'orig', u: number): boolean =>
    [chars[i - 1], chars[i + 1]].some((x) =>
      x !== undefined && (side === 'rev' ? x.op !== 'del' && x.ur === u : x.op !== 'ins' && x.uo === u),
    );
  const stray = (i: number): boolean => {
    const { op, c, ur, uo } = chars[i]!;
    if (op !== 'eq' || c !== ' ') return false;
    const revStray = isObject(r, ur) && revCounterpart(ur) !== uo && (isObject(o, uo) || !nextTo(i, 'rev', ur));
    const origStray = isObject(o, uo) && origCounterpart.get(uo) !== ur && (isObject(r, ur) || !nextTo(i, 'orig', uo));
    return revStray || origStray;
  };

  const out: WordDiff[] = [];
  const push = (op: WordDiff['op'], c: string): void => {
    const last = out[out.length - 1];
    if (last !== undefined && last.op === op) out[out.length - 1] = { op, text: last.text + c };
    else out.push({ op, text: c });
  };
  let split = false;
  chars.forEach((x, i) => {
    if (stray(i)) {
      push('del', x.c);
      push('ins', x.c);
      split = true;
    } else {
      push(x.op, x.c);
    }
  });
  return split ? out : null;
}

/**
 * One emitted stretch of a paragraph's runs: wholly inserted, wholly deleted,
 * a bookmark (in both views, placed from the revised side), or anything else.
 */
interface Piece {
  xml: string;
  kind: 'ins' | 'del' | 'zw' | 'bar';
}

const REVISION_ID = /(<w:(?:ins|del) w:id=")(\d+)(")/g;

/**
 * ADR-0043's order, restored after the walk. Stage 6.5 put each changed region's
 * insertion first in the word diff, but the walk makes revisions of its own —
 * a stretch plain on one side and inside a changed object on the other becomes
 * an insertion beside a whole-object deletion, and a split space is a deletion
 * then an insertion — so a deletion can open a region. Within a run of
 * consecutive inserted and deleted pieces, with nothing unchanged between,
 * the deletions that precede the region's last insertion are moved after it.
 * A bookmark in the region does not end it, and keeps its place among the
 * insertions: the accepted view is untouched — the bookmark still opens where
 * it did — and in the rejected view only struck text moves past it, where the
 * revised side's bookmarks are placed anyway (§2). A kept footnote reference
 * or image DOES end it: it belongs to the original's view as much as the
 * revised, and moving struck text past it misplaces it there — so a deletion
 * before one and an insertion after it keep that order (ADR-0045 § Accepted
 * limits 13). Each view reads the same text either way, and the
 * revision ids, minted in emission order, are re-dealt in the new order so
 * they still open in document order.
 */
function insertionsFirst(pieces: readonly Piece[]): string {
  const out: string[] = [];
  for (let a = 0; a < pieces.length; ) {
    if (pieces[a]!.kind === 'bar') {
      out.push(pieces[a++]!.xml);
      continue;
    }
    let b = a;
    while (b < pieces.length && pieces[b]!.kind !== 'bar') b++;
    const region = pieces.slice(a, b);
    const last = region.map((p) => p.kind).lastIndexOf('ins');
    const head = region.slice(0, last + 1);
    const ordered = [...head.filter((p) => p.kind !== 'del'), ...head.filter((p) => p.kind === 'del'), ...region.slice(last + 1)];
    const ids = region.flatMap((p) => [...p.xml.matchAll(REVISION_ID)].map((m) => Number(m[2])));
    const dealt = [...ids].sort((x, y) => x - y);
    let next = 0;
    for (const p of ordered) out.push(p.xml.replace(REVISION_ID, (_m, open: string, _id: string, close: string) => `${open}${dealt[next++]}${close}`));
    a = b;
  }
  return out.join('');
}

/**
 * The paragraph's runs, its objects put back around its text. With `source`,
 * text is rebuilt from its own runs (`spanLeaves`), and `null` means a
 * character had no source to rebuild it from.
 */
function walk(
  wordDiffs: readonly WordDiff[],
  o: SideMap,
  r: SideMap,
  ids: RevisionIds,
  emitText: (wd: WordDiff, ids: RevisionIds) => string,
  source = false,
): string | null {
  const { verbatim, inside, insideOrig } = pairObjects(wordDiffs, o, r);

  // Zero-width objects: kept from the revised side when both sides carry the
  // same sequence of sorts, otherwise all marked. Bookmarks: revised side, never marked.
  const zwObjects = (side: SideMap) => side.zeroWidth.filter((z) => isObject(side, z.unit));
  // A sort is its kind AND what it holds: a DrawingML shape with a VML fallback
  // and a plain inline picture are both `w:drawing`, and are not the same object.
  const sorts = (side: SideMap) =>
    zwObjects(side)
      .map((z) => `${side.units[z.unit]!.objectKind}:${census(buildNodes(side.units[z.unit]!.nodes), false)}`)
      .join('\u0001');
  // One that holds a space all the same — a link of only a space, its space
  // collapsed or trimmed away — carries that space to wherever it is emitted,
  // so the revised one cannot stand in for the original: both are marked.
  const spaceBearing = (side: SideMap) => zwObjects(side).some((z) => side.shared.has(z.unit));
  const zwSame = sorts(o) === sorts(r) && !spaceBearing(o) && !spaceBearing(r);
  // Even then, one revised object stands for both sides only where ONE point
  // of the stream is its place in both: after the characters before it on each
  // side, before the characters after it. Words replaced on both sides of an
  // unmoved footnote put the original's place inside the struck text and the
  // revised one inside the insertion; a footnote that moved has two places.
  // Kept there, rejecting every change would misplace it — so such a pair is
  // marked, the revised one inserted where it is and the original struck where
  // it was. The revised one is emitted just before its next character (or
  // after a deletion that opens a container), which is the point tested.
  const revAt: number[] = [];
  const origAt: number[] = [];
  let streamLength = 0;
  for (const wd of wordDiffs) {
    for (let k = 0; k < wd.text.length; k++, streamLength++) {
      if (wd.op !== 'del') revAt.push(streamLength);
      if (wd.op !== 'ins') origAt.push(streamLength);
    }
  }
  /** Kept revised zero-width object → its original counterpart. */
  const keptZw = new Map<number, number>();
  const origZw: SideMap['zeroWidth'] = [];
  const rz = zwObjects(r);
  zwObjects(o).forEach((oz, k) => {
    const rev = zwSame ? rz[k] : undefined;
    if (rev !== undefined) {
      const cut = revAt[rev.pos] ?? streamLength;
      const after = oz.pos > 0 ? origAt[oz.pos - 1]! + 1 : 0;
      const before = origAt[oz.pos] ?? streamLength;
      if (after <= cut && cut <= before) {
        keptZw.set(rev.unit, oz.unit);
        return;
      }
    }
    origZw.push(oz);
  });
  const revZw = r.zeroWidth;

  const out: Piece[] = [];
  const emittedRev = new Set<number>();
  const emittedOrig = new Set<number>();
  let buf = '';
  let bufOp: WordDiff['op'] = 'eq';
  /** With `source`: where each buffered character comes from — a struck one the original, else the revised. */
  let bufSrc: TextLeaf[] = [];
  /** The current stream character's index on each side (-1 where it has none). */
  let here = { rev: -1, orig: -1 };
  let unsourced = false;
  /** Has the revised paragraph's trailing whitespace gone out with an inserted last character? */
  let revTrailing = false;
  /** The revised unit of the object currently open for an inside diff, or -1. */
  let openRev = -1;
  /**
   * Whole-object deletions waiting for the insertion beside them to finish.
   * An object is emitted when the walk first touches it, and when an `eq`
   * stretch is plain text on one side and an object on the other, the deletion
   * is touched first. Held here, it lands after the insertion, which is the
   * order every other changed region has (ADR-0043).
   */
  let pendingDel: (() => string)[] = [];
  const flushPendingDel = (): void => {
    // Built only now, so its ids are minted in document order (§5.5).
    for (const build of pendingDel) out.push({ xml: build(), kind: 'del' });
    pendingDel = [];
  };
  const flushBuf = (): void => {
    if (buf === '') return;
    let run: string;
    if (source) {
      // Each character from its own run — inside an open object, the object's.
      run = revise(leafRuns(bufSrc), bufOp, ids);
    } else {
      run = emitText({ op: bufOp, text: buf }, ids);
      // Text inside an open object wears that object's run properties — on the
      // struck text too, which `emitText` has already rewritten to `w:delText`.
      if (openRev >= 0) {
        const rPr = inside.get(openRev)!.shell.rPr;
        if (rPr !== '') run = run.split('<w:r><w:').join(`<w:r>${rPr}<w:`);
      }
    }
    if (run !== '') out.push({ xml: run, kind: bufOp === 'eq' ? 'bar' : bufOp });
    buf = '';
    bufSrc = [];
  };
  /** Settle everything buffered before a non-insertion is emitted. */
  const settle = (): void => {
    flushBuf();
    flushPendingDel();
  };
  const text = (op: WordDiff['op'], c: string): void => {
    if (op !== bufOp) flushBuf();
    if (op !== 'ins') flushPendingDel();
    bufOp = op;
    buf += c;
    if (source) {
      // A revised last character takes its side's trailing whitespace with it.
      if (op === 'ins' && isLastOf(r, here.rev)) revTrailing = true;
      const leaves = op === 'del' ? spanLeaves(o, here.orig, true) : spanLeaves(r, here.rev, op === 'ins');
      if (leaves === null) unsourced = true;
      else bufSrc.push(...leaves);
    }
  };
  const closeInside = (): void => {
    if (openRev < 0) return;
    settle();
    out.push({ xml: inside.get(openRev)!.shell.close, kind: 'bar' });
    openRev = -1;
  };
  const openInside = (ur: number): void => {
    if (openRev === ur) return;
    closeInside();
    settle();
    out.push({ xml: inside.get(ur)!.shell.open, kind: 'bar' });
    openRev = ur;
  };
  const whole = (side: SideMap, u: number, kind: Kind, emitted: Set<number>): void => {
    if (emitted.has(u)) return;
    emitted.add(u);
    if (kind === 'ins') {
      flushBuf();
      out.push({ xml: reviseNodes(side.units[u]!.nodes, 'ins', ids), kind: 'ins' });
      return;
    }
    // A deletion keeps its place among deletions, and waits behind an insertion.
    if (bufOp !== 'ins') flushBuf();
    pendingDel.push(() => reviseNodes(side.units[u]!.nodes, 'del', ids));
  };
  /** An unchanged object: once, verbatim, from the revised paragraph. */
  const verbatimOrig = new Set(verbatim.values());
  const lastRevChar = new Map<number, number>();
  r.charUnit.forEach((u, i) => lastRevChar.set(u, i));
  const keep = (ur: number): void => {
    if (emittedRev.has(ur)) return;
    settle();
    out.push({ xml: buildNodes(r.units[ur]!.nodes), kind: 'bar' });
    emittedRev.add(ur);
    emittedOrig.add(verbatim.get(ur)!);
  };
  let zr = 0;
  let zo = 0;
  const flushRevZw = (upTo: number): void => {
    while (zr < revZw.length && revZw[zr]!.pos <= upTo) {
      const z = revZw[zr++]!;
      const unit = r.units[z.unit]!;
      // A struck original that came before this one's counterpart comes before
      // it here too, or rejecting would swap two objects in one place.
      const counterpart = keptZw.get(z.unit);
      if (counterpart !== undefined) flushOrigZwBefore(counterpart);
      if (unit.kind === 'object' && !keptZw.has(z.unit)) {
        flushBuf();
        out.push({ xml: reviseNodes(unit.nodes, 'ins', ids), kind: 'ins' });
      } else {
        settle();
        // A bookmark is placed from the revised side in both views anyway (§2),
        // so struck text may move past it. A kept footnote reference or image
        // is the original's too: struck text moved past it would put it in the
        // wrong place once every change is rejected.
        out.push({ xml: buildNodes(unit.nodes), kind: unit.kind === 'marker' ? 'zw' : 'bar' });
      }
    }
  };
  const emitOrigZw = (): void => {
    const z = origZw[zo++]!;
    settle();
    out.push({ xml: reviseNodes(o.units[z.unit]!.nodes, 'del', ids), kind: 'del' });
  };
  const flushOrigZw = (upTo: number): void => {
    while (zo < origZw.length && origZw[zo]!.pos <= upTo) emitOrigZw();
  };
  /** Struck originals earlier in the paragraph than `unit` (units are in document order). */
  const flushOrigZwBefore = (unit: number): void => {
    while (zo < origZw.length && origZw[zo]!.unit < unit) emitOrigZw();
  };


  let po = 0;
  let pr = 0;
  for (const wd of wordDiffs) {
    for (let k = 0; k < wd.text.length; k++) {
      const c = wd.text[k]!;
      here = { rev: wd.op === 'del' ? -1 : pr, orig: wd.op === 'ins' ? -1 : po };
      const ru = wd.op === 'del' ? FREE : r.charUnit[pr]!;
      const ou = wd.op === 'ins' ? FREE : o.charUnit[po]!;
      // A boundary-only space (`FREE`) is no character of its side. An `eq`
      // whose one side is free is that other side's own space — an insertion
      // or a deletion of plain text; anything wholly free is authored by nobody.
      let op: WordDiff['op'] | 'none' = wd.op;
      if (wd.op === 'eq' && ru === FREE) op = ou === FREE ? 'none' : 'del';
      else if (wd.op === 'eq' && ou === FREE) op = 'ins';
      else if ((wd.op === 'ins' && ru === FREE) || (wd.op === 'del' && ou === FREE)) op = 'none';
      const ur = op === 'del' || op === 'none' ? FREE : ru;
      const uo = op === 'ins' || op === 'none' ? FREE : ou;
      // Which open-for-diffing object this character belongs to, if any. The
      // correspondence is exclusive for every character but a space at an
      // object's edge (see the meets pass), which joins the shell it touches.
      const owner = inside.has(ur) ? ur : insideOrig.has(uo) ? insideOrig.get(uo)! : -1;
      if (owner !== openRev) closeInside();
      if (wd.op !== 'del') flushRevZw(pr);
      if (wd.op !== 'ins') flushOrigZw(po);
      if (wd.op !== 'del') pr++;
      if (wd.op !== 'ins') po++;
      if (op === 'none') continue;

      if (owner >= 0) {
        // Anything zero-width sitting at the object's start belongs before its
        // shell, whichever side's character opens it — a diff that opens with
        // a deletion would otherwise carry a footnote or bookmark inside.
        if (owner !== openRev) {
          flushRevZw(wd.op === 'del' ? pr : pr - 1);
          flushOrigZw(wd.op === 'ins' ? po : po - 1);
        }
        openInside(owner);
        text(op, c);
        continue;
      }
      const revObj = isObject(r, ur);
      const origObj = isObject(o, uo);
      const revKept = revObj && verbatim.has(ur);
      const origKept = origObj && verbatimOrig.has(uo);
      if (op === 'eq') {
        if (revKept && verbatim.get(ur) === uo) {
          keep(ur);
          continue;
        }
        // A space at a kept object's edge on one side only (see the meets pass).
        // The kept object carries its own space in its XML, so the character is
        // plain text at most — and the object is emitted at its first REAL
        // character, not here: at a LEADING edge, a deletion of the original's
        // can still stand between this space and the object's words, and
        // emitting the object now would put that deletion after it.
        if (c === ' ' && (revKept || origKept)) {
          if (revKept) {
            if (lastRevChar.get(ur) === pr - 1) keep(ur);
          } else if (revObj) {
            whole(r, ur, 'ins', emittedRev);
          }
          if (origObj && !origKept) whole(o, uo, 'del', emittedOrig);
          // A kept ORIGINAL's space is not in its verbatim counterpart (that
          // would have met it above): struck beside it, unless a plain space
          // on the revised side, unmarked, already stands for both.
          if (!revObj) text(origObj && !origKept ? 'ins' : 'eq', c);
          else if (!origObj || origKept) text('del', c);
          continue;
        }
        if (revObj) whole(r, ur, 'ins', emittedRev);
        if (origObj) whole(o, uo, 'del', emittedOrig);
        // Text on one side that sits inside a changed object on the other: the
        // object carries its own copy, so the plain side shows as the change.
        if (!revObj && !origObj) text('eq', c);
        else if (!revObj) text('ins', c);
        else if (!origObj) text('del', c);
      } else if (op === 'ins') {
        // Only a re-attributed edge space can reach a kept object here.
        if (revKept) {
          if (lastRevChar.get(ur) === pr - 1) keep(ur);
        } else if (revObj) {
          whole(r, ur, 'ins', emittedRev);
        } else {
          text('ins', c);
        }
      } else if (origKept) {
        // Only a space can reach a kept original here: its own edge space, lined
        // up with a boundary the revised side reads where it has none (`[Buyer ]
        // within [box]` → `[Buyer][box]`). The verbatim counterpart lacks it, so
        // it is struck as plain text beside the object.
        text('del', c);
      } else if (origObj) {
        whole(o, uo, 'del', emittedOrig);
      } else {
        text('del', c);
      }
    }
    flushBuf();
  }
  closeInside();
  flushRevZw(Number.POSITIVE_INFINITY);
  flushOrigZw(Number.POSITIVE_INFINITY);
  settle();
  // Otherwise the revised paragraph's trailing whitespace ends it, unmarked:
  // the original's cannot also be had (the two-sided whitespace hazard).
  const trailing = source && !revTrailing ? trailingLeaves(r) : [];
  if (trailing.length > 0) out.push({ xml: leafRuns(trailing), kind: 'bar' });
  if (unsourced) return null;
  // Only an object makes revisions of the walk's own; text alone is the word
  // diff's revisions in the word diff's order, which the author never re-sorts
  // (§5.5 — stage 6.5 decided it).
  const objects = [o, r].some((side) => side.units.some((u) => u.kind === 'object'));
  return objects ? insertionsFirst(out) : out.map((piece) => piece.xml).join('');
}
