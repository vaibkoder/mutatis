/**
 * The inline content of one paragraph: how it becomes the text Compare diffs,
 * and the units the replace author puts back around a word diff (ADR-0045).
 * Spec: spec/features/08-compare.md §5.0, §5.5.
 *
 * ONE home for both, because they must agree character for character. The
 * replace author maps every word-diff character back to the unit that produced
 * it, and it can do that only if the text of a paragraph is exactly the
 * concatenation of the text of its units — which is true here by construction,
 * since `collectText` is the extractor `extractParaRecords` uses and it is a
 * plain left-to-right concatenation over child nodes.
 *
 * Pure. preserveOrder nodes throughout, built with the shared parser options.
 */
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

import { PRESERVE_ORDER_OPTS, tagOf, childrenOf, attrOf, type PoNode } from '@shared/ooxml/po-xml.js';

import type { ZeroTextObject, ZeroTextSort } from '../model.js';
import { normaliseTextWithOffsets } from './normalise.js';

/** The builder options every Compare module serialises preserveOrder nodes with. */
export const BUILDER_OPTS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  suppressEmptyNode: true,
  format: false,
} as const;

/**
 * Elements whose content is a separate run of words, so a word on either side
 * of one must not fuse with a word inside it. A display equation, a text box and
 * the paragraphs inside a text box each read as their own words; before ADR-0045
 * a text box's text was glued to the sentence it was anchored in
 * (`"the Price.Boxed note"`). Read as a space, which `normaliseText` collapses.
 */
const WORD_BOUNDARY_TAGS = new Set(['m:oMath', 'w:txbxContent', 'w:p']);

/**
 * A mark a caller's `visit` may push into `parts` to find a place in the text
 * once it is complete — the boundary around an equation or text box is spliced
 * in AFTER its content is read, so a length taken during the walk is short.
 * U+FFFF cannot occur in XML 1.0 text, and a mark never makes an empty equation
 * or box read as non-empty.
 */
export const PLACE_MARK = '\uFFFF';

/**
 * Append the visible text of `nodes` to `parts`, in document order.
 *
 * - `<w:t>` and **`<m:t>`** contribute their text. `<m:t>` is an Office Math
 *   (OMML) run's text: before ADR-0045 it was skipped, so an equation was
 *   invisible to the diff and a changed formula reported no change at all.
 * - `<w:tab>`, `<w:br>`, `<w:cr>` contribute a space, so adjacent words do not
 *   fuse.
 * - `<w:instrText>` contributes nothing: a field is compared by its RESULT, the
 *   text the reader sees, never by its instruction.
 * - `<mc:Fallback>` contributes nothing. Word writes a text box (and most
 *   drawings) as `<mc:AlternateContent>`: a DrawingML `<mc:Choice>` and a VML
 *   `<mc:Fallback>` copy of the SAME content for older readers, so reading both
 *   read every word in the box twice. The choice is read; the fallback is kept
 *   in the XML, untouched.
 * - Everything else is descended into.
 */
export function collectText(
  nodes: readonly PoNode[],
  parts: string[],
  boundary = ' ',
  visit?: (tag: string, node: PoNode) => void,
): void {
  for (const node of nodes) {
    const tag = tagOf(node);
    if (tag === undefined) continue; // a bare text node is read via its <w:t>/<m:t>
    if (tag === 'w:t' || tag === 'm:t') {
      parts.push(leafText(node));
      continue;
    }
    if (tag === 'w:tab' || tag === 'w:br' || tag === 'w:cr') {
      parts.push(' ');
      continue;
    }
    if (tag === 'mc:Fallback') continue;
    visit?.(tag, node);
    if (boundary !== '' && WORD_BOUNDARY_TAGS.has(tag)) {
      // An equation or text box with no text has no words to keep apart from
      // its neighbours, so it reads as nothing — as an emptied one must: a
      // shell a revision leaves behind would otherwise put a space between two
      // words that were never apart.
      const at = parts.length;
      collectText(childrenOf(node), parts, boundary, visit);
      const inside = parts.slice(at);
      if (inside.every((p) => p === '' || p === PLACE_MARK)) {
        parts.length = at;
        parts.push(...inside.filter((p) => p === PLACE_MARK)); // a caller's marks stay
      } else {
        parts.splice(at, 0, boundary);
        parts.push(boundary);
      }
      continue;
    }
    collectText(childrenOf(node), parts, boundary, visit);
  }
}

/**
 * Where one character of {@link collectText}'s output came from: the
 * run it sits in and, for a tab or break, the element itself — so a changed
 * paragraph can be rebuilt from its own runs, formatting and all, instead of
 * from the word diff's normalised text.
 */
export interface TextLeaf {
  /** The `<w:r>` (or `<m:r>`) holding the character. */
  run: PoNode;
  /** The character as `collectText` read it (a tab or break reads as a space). */
  char: string;
  /** The `<w:tab/>`, `<w:br/>` or `<w:cr/>` element it was read from, or null for text. */
  item: PoNode | null;
  /**
   * Run content `collectText` reads as NOTHING that stands just before this
   * character in its run — a non-breaking or soft hyphen — and, on the run's
   * last character, just after it. Re-emitted with the character.
   */
  before: PoNode[];
  after: PoNode[];
  /**
   * Whole runs of nothing but silent content (a run holding only a non-breaking
   * hyphen, as Word writes one), re-emitted as they are before this character
   * — or, at the end of the paragraph, after it.
   */
  runsBefore: PoNode[];
  runsAfter: PoNode[];
}

/** Run children `collectText` reads as nothing but the document shows. */
const SILENT_RUN_CONTENT = new Set(['w:noBreakHyphen', 'w:softHyphen']);

/**
 * The source of every character {@link collectText} reads from `nodes`, in the
 * same order and one for one — a boundary character (the space read around an
 * equation or text box) is `null`, being no character of the document. A
 * caller checks the length against the text before trusting it.
 */
export function textLeaves(nodes: readonly PoNode[]): (TextLeaf | null)[] {
  const out: (TextLeaf | null)[] = [];
  /** Runs of only silent content, waiting for the next character. */
  let silentRuns: PoNode[] = [];
  const walk = (list: readonly PoNode[], run: PoNode | null): void => {
    /** Silent content waiting for the next character of this run. */
    let pending: PoNode[] = [];
    let last: TextLeaf | null = null;
    const push = (leaf: Pick<TextLeaf, 'run' | 'char' | 'item'>): void => {
      last = { ...leaf, before: pending, after: [], runsBefore: silentRuns, runsAfter: [] };
      pending = [];
      silentRuns = [];
      out.push(last);
    };
    for (const node of list) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:t' || tag === 'm:t') {
        // Per UTF-16 code unit, as `collectText`'s text is indexed: a character
        // outside the Basic Multilingual Plane is two leaves, rejoined on output
        // when the word diff keeps them together (it may not: spec 08 §12 item 12).
        const text = leafText(node);
        for (let k = 0; k < text.length; k++) if (run !== null) push({ run, char: text[k]!, item: null });
        continue;
      }
      if (tag === 'w:tab' || tag === 'w:br' || tag === 'w:cr') {
        if (run !== null) push({ run, char: ' ', item: node });
        continue;
      }
      if (SILENT_RUN_CONTENT.has(tag) && run !== null) {
        pending.push(node);
        continue;
      }
      if (tag === 'mc:Fallback') continue;
      if (WORD_BOUNDARY_TAGS.has(tag)) {
        const at = out.length;
        walk(childrenOf(node), tag === 'w:p' ? null : run);
        if (out.length > at) {
          out.splice(at, 0, null);
          out.push(null);
        }
        continue;
      }
      walk(childrenOf(node), tag === 'w:r' || tag === 'm:r' ? node : run);
    }
    // Silent content after the run's last character stays with it; a run of
    // nothing else travels whole with the next character.
    if (pending.length > 0 && last !== null) (last as TextLeaf).after.push(...pending);
    else if (pending.length > 0 && run !== null && list === childrenOf(run)) silentRuns.push(run);
  };
  walk(nodes, null);
  const final = [...out].reverse().find((leaf) => leaf !== null);
  if (silentRuns.length > 0 && final) final.runsAfter.push(...silentRuns);
  return out;
}

/**
 * The visible text of a `<w:p>` node (its children; `<w:pPr>` contributes none).
 *
 * `boundaries: false` leaves out the spaces read around an equation or a text
 * box, giving only the document's own characters. Tests compare in that form:
 * a boundary space can stand in for a REAL space the author dropped, so two
 * texts that agree with boundaries can disagree without them — a lost space
 * after an equation read `"r=3%per annum"` and still matched. The author's own
 * self-check does NOT (ADR-0045 §3): it must accept what the diff calls equal.
 */
export function paragraphText(paraNode: PoNode, boundaries = true): string {
  const parts: string[] = [];
  collectText(childrenOf(paraNode), parts, boundaries ? ' ' : '');
  return parts.join('');
}

/**
 * Is this paragraph child ordinary text content — a run of nothing but text,
 * tabs and breaks, or proofing noise? A container holding only such children
 * can have its words rebuilt from a word diff and lose nothing.
 */
export function isPlainTextNode(node: PoNode): boolean {
  const tag = tagOf(node);
  if (tag === undefined) return true;
  if (TEXT_NOISE.has(tag)) return true;
  if (tag !== 'w:r') return false;
  return childrenOf(node).every((c) => {
    const t = tagOf(c);
    return t === undefined || TEXT_RUN_CHILDREN.has(t);
  });
}

/** The string content of a text leaf (`<w:t>` / `<m:t>`): its `#text` children. */
function leafText(node: PoNode): string {
  let s = '';
  for (const child of childrenOf(node)) {
    const t = child['#text'];
    if (typeof t === 'string') s += t;
  }
  return s;
}

// ---------------------------------------------------------------------------
// Units
// ---------------------------------------------------------------------------

/**
 * One unit of a paragraph's inline content.
 *
 * - `text`   — ordinary runs. The replace author REBUILDS these from the word
 *              diff, exactly as it always has.
 * - `object` — anything that is not ordinary text: a hyperlink, a field, a
 *              footnote reference, an equation, a drawing or text box, a content
 *              control. The replace author re-emits these VERBATIM when the diff
 *              left their text untouched, and as a whole tracked deletion and
 *              insertion when it did not (ADR-0045).
 * - `marker` — a zero-width position marker (a bookmark). Never compared, never
 *              revised: the revised side's markers are re-emitted where they sit.
 */
export interface InlineUnit {
  kind: 'text' | 'object' | 'marker';
  /**
   * What sort of object, for pairing an object on one side with the object on
   * the other: an object is "unchanged" only against an object of the same
   * sort. The element name, `field` for a complex field, or `w:r/<child>` for a
   * run that carries something other than text.
   */
  objectKind: string;
  /** The top-level paragraph children making up this unit, in order. */
  nodes: PoNode[];
  /** This unit's contribution to the paragraph text (`collectText` over `nodes`). */
  text: string;
  /**
   * Offsets in `text` of the BOUNDARY spaces `collectText` reads around an
   * equation or text box — characters of the paragraph text that are not
   * characters of the document. A collapsed whitespace run must never be
   * credited to one of these when it holds a real space (`inline-objects.ts`).
   */
  boundaries: ReadonlySet<number>;
}

/** Children of an ordinary text run. Anything else makes the run an object. */
const TEXT_RUN_CHILDREN = new Set([
  'w:rPr', 'w:t', 'w:tab', 'w:br', 'w:cr', 'w:lastRenderedPageBreak',
  'w:softHyphen', 'w:noBreakHyphen',
]);

/**
 * Paragraph children with no content of their own that the text rebuild has
 * always dropped. Kept inside a text unit so they do not split it.
 */
const TEXT_NOISE = new Set(['w:proofErr', 'w:permStart', 'w:permEnd']);

/** Zero-width position markers, re-emitted from the revised side. */
const MARKERS = new Set(['w:bookmarkStart', 'w:bookmarkEnd']);

const parser = new XMLParser(PRESERVE_ORDER_OPTS);
const builder = new XMLBuilder(BUILDER_OPTS);

/** Serialise preserveOrder nodes exactly as `extractParaRecords` built `rawXml`. */
export function buildNodes(nodes: readonly PoNode[]): string {
  return builder.build(nodes) as string;
}

/** Parse an XML fragment into preserveOrder nodes with the shared options. */
export function parseNodes(xml: string): PoNode[] {
  return parser.parse(xml) as PoNode[];
}

/** The first `<w:p>` node of a paragraph's XML, or undefined. */
export function paragraphNode(paraXml: string): PoNode | undefined {
  return parseNodes(paraXml).find((n) => tagOf(n) === 'w:p');
}

/**
 * Cut a paragraph's inline content into units, or `null` when it holds a shape
 * the unit model does not cover — a complex field that opens or closes outside
 * this paragraph (a multi-paragraph TOC, say). A `null` sends the caller back to
 * the text-only rebuild, which is what every paragraph got before ADR-0045.
 */
export function paragraphUnits(paraXml: string): InlineUnit[] | null {
  const para = paragraphNode(paraXml);
  if (para === undefined) return null;
  const units: InlineUnit[] = [];
  let text: PoNode[] | null = null;
  let field: PoNode[] | null = null;
  let fieldDepth = 0;

  const push = (kind: InlineUnit['kind'], objectKind: string, nodes: PoNode[]): void => {
    const parts: string[] = [];
    collectText(nodes, parts);
    // The same walk with a marker for each boundary: one character either way,
    // so offsets line up. NUL cannot occur in XML 1.0 text.
    const marked: string[] = [];
    collectText(nodes, marked, '\u0000');
    const boundaries = new Set<number>();
    const m = marked.join('');
    for (let i = 0; i < m.length; i++) if (m[i] === '\u0000') boundaries.add(i);
    units.push({ kind, objectKind, nodes, text: parts.join(''), boundaries });
  };
  const closeText = (): void => {
    if (text !== null) push('text', 'text', text);
    text = null;
  };

  for (const node of childrenOf(para)) {
    const tag = tagOf(node);
    if (tag === undefined || tag === 'w:pPr') continue;

    // Inside a complex field, every sibling up to the matching `end` belongs to
    // it — result runs, instruction runs, a bookmark around the result.
    if (field !== null) {
      field.push(node);
      fieldDepth += fieldCharDelta(node);
      if (fieldDepth === 0) {
        push('object', 'field', field);
        field = null;
      }
      continue;
    }

    if (tag === 'w:r') {
      const delta = fieldCharDelta(node);
      if (delta > 0) {
        closeText();
        field = [node];
        fieldDepth = delta;
        continue;
      }
      // A `separate` or `end` with no `begin` before it in this paragraph.
      if (delta < 0 || hasFieldChar(node)) return null;
      const other = childrenOf(node).map(tagOf).find((t) => t !== undefined && !TEXT_RUN_CHILDREN.has(t));
      if (other === undefined) {
        (text ??= []).push(node);
        continue;
      }
      closeText();
      push('object', `w:r/${drawingKind(other)}`, [node]);
      continue;
    }
    if (TEXT_NOISE.has(tag)) {
      (text ??= []).push(node);
      continue;
    }
    closeText();
    if (MARKERS.has(tag)) {
      push('marker', tag, [node]);
      continue;
    }
    push('object', tag, [node]);
  }
  if (field !== null) return null; // a field still open at the paragraph mark
  closeText();
  return units;
}

/** Modern and legacy drawings are one sort of object for pairing. */
function drawingKind(tag: string): string {
  return tag === 'mc:AlternateContent' || tag === 'w:pict' ? 'w:drawing' : tag;
}

/** Net field nesting a run opens (+1 per `begin`, -1 per `end`). */
function fieldCharDelta(node: PoNode): number {
  if (tagOf(node) !== 'w:r') return 0;
  let delta = 0;
  for (const child of childrenOf(node)) {
    if (tagOf(child) !== 'w:fldChar') continue;
    const type = attrOf(child, '@_w:fldCharType');
    if (type === 'begin') delta++;
    else if (type === 'end') delta--;
  }
  return delta;
}

function hasFieldChar(node: PoNode): boolean {
  return childrenOf(node).some((c) => tagOf(c) === 'w:fldChar');
}

// ---------------------------------------------------------------------------
// Zero-text objects (ADR-0045 § Amendment — limit 1)
// ---------------------------------------------------------------------------

/** Cheap pre-filter: a paragraph with none of these tags has no zero-text object. */
const ZERO_TEXT_TAG = /<(?:w:footnoteReference|w:endnoteReference|w:drawing|w:pict|w:object)[\s/>]/;

/**
 * The zero-text objects of a paragraph, in document order. A text box counts
 * once as a text box (its words are compared as words); what is inside it is
 * read as well. A DrawingML choice and its VML fallback are one object: the
 * fallback is skipped exactly as `collectText` skips it.
 *
 * Comments and bookmarks are not here: they are review metadata and position
 * markers, not content of the clause.
 *
 * An object inside an equation or text box also carries `within`, the ordinal
 * of the innermost one holding it: whitespace collapse can make "just before
 * the box" and "just inside its start" the same count of characters.
 */
export function zeroTextObjects(paraXml: string, numberingPrefix: string): ZeroTextObject[] {
  if (!ZERO_TEXT_TAG.test(paraXml)) return [];
  const para = paragraphNode(paraXml);
  if (para === undefined) return [];
  const parts: string[] = [];
  const found: { sort: ZeroTextSort; node: PoNode }[] = [];
  /** Equations, text box contents and their paragraphs, in document order. */
  const containers: PoNode[] = [];
  collectText(childrenOf(para), parts, ' ', (tag, node) => {
    if (WORD_BOUNDARY_TAGS.has(tag)) containers.push(node);
    const sort = sortOf(tag, node);
    if (sort === null) return;
    found.push({ sort, node });
    parts.push(PLACE_MARK);
  });
  if (found.length === 0) return [];
  // Each mark's offset in the text without the marks.
  let raw = '';
  const marks: number[] = [];
  for (const c of parts.join('')) {
    if (c === PLACE_MARK) marks.push(raw.length);
    else raw += c;
  }
  const { offsets } = normaliseTextWithOffsets(raw);
  const prefixLen = numberingPrefix.length;
  return found.map(({ sort, node }, k) => {
    let before = 0;
    while (before < offsets.length && offsets[before]! < marks[k]!) before++;
    // In the manual number rather than after it: a code, not a count, so a
    // renumbered clause does not move it and a move across the number does.
    // Same rule as the author's `manualNumberPlace` (output/inline-objects.ts).
    const at = before >= prefixLen ? before - prefixLen : before === 0 ? -2 : -1;
    // The innermost equation or text-box paragraph holding it. A collapsed
    // space makes "just before the box" and "just inside its start" the same
    // count of characters; they are not the same place.
    let within = -1;
    containers.forEach((c, i) => {
      if (contains(c, node)) within = i;
    });
    return within >= 0 ? { sort, at, within } : { sort, at };
  });
}

/** Is `node` a descendant of `ancestor` (by identity)? */
function contains(ancestor: PoNode, node: PoNode): boolean {
  return childrenOf(ancestor).some((c) => c === node || (tagOf(c) !== undefined && contains(c, node)));
}

function sortOf(tag: string, node: PoNode): ZeroTextSort | null {
  switch (tag) {
    case 'w:footnoteReference':
      return 'footnote reference';
    case 'w:endnoteReference':
      return 'endnote reference';
    case 'w:drawing':
    case 'w:pict':
      return holds(node, 'w:txbxContent') ? 'text box' : 'picture';
    case 'w:object':
      return 'embedded object';
    default:
      return null;
  }
}

function holds(node: PoNode, tag: string): boolean {
  return childrenOf(node).some((c) => tagOf(c) === tag || (tagOf(c) !== undefined && holds(c, tag)));
}
