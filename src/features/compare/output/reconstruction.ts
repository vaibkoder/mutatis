/**
 * Feature 08 — Compare: does a redline reconstruct its two documents, as a
 * reader sees them?
 *
 * Accepting every revision in an authored redline must give the revised
 * document and rejecting every one the original — and "give" means what a
 * reader of the two documents sees, not what Compare's own model reads:
 *
 *  (a) the VISIBLE CHARACTERS of each paragraph, raw: `<w:t>`/`<m:t>` text,
 *      `<w:tab/>` as `\t`, `<w:br/>`/`<w:cr/>` as `\n`, a non-breaking hyphen
 *      as U+2011 and a soft hyphen as U+00AD. NO §5.2 normalisation — a curly
 *      apostrophe written straight, a non-breaking space written as a space and
 *      a tab written as a space are all differences here, because they are
 *      differences in the document a reviewer gets back;
 *  (b) a RUN-FORMATTING SIGNATURE — for each character the canonical
 *      `<w:rPr>` of the run it sits in (child elements sorted, `w:rsid*`
 *      attributes stripped, `<w:rPrChange>` ignored);
 *  (c) the PARAGRAPH COUNT.
 *
 * Blank paragraphs (no visible character) are dropped from all three, so an
 * empty paragraph ADR-0035 had to leave behind (§5.5) is not counted.
 *
 * ONE implementation, two users: the replace author's self-check holds every
 * paragraph it rebuilds from its source runs to this (`inline-objects.ts`),
 * and the test yardstick (`tests/fixtures/reconstruct.ts`) holds whole
 * redlines to it. Both views come from the production flattener
 * (`preprocess/flatten.ts`), so this measures the author AND the flattener.
 *
 * Every paragraph element is its own paragraph, a text box's paragraphs
 * included (after the paragraph anchoring the box); `mc:Fallback` is skipped
 * as `collectText` skips it (it repeats the choice's content).
 *
 * Spec: spec/features/08-compare.md §5.5.
 */
import DiffMatchPatch from 'diff-match-patch';

import { parsePreserveOrder, tagOf, childrenOf, type PoNode } from '@shared/ooxml/po-xml.js';

import { computeLCS } from '../diff/lcs.js';
import { flattenTrackedChanges } from '../preprocess/flatten.js';
import { buildNodes } from '../preprocess/inline-units.js';
import { normaliseText } from '../preprocess/normalise.js';

/** One non-blank paragraph as a reader sees it. */
export interface ParaView {
  /** Visible characters, raw (see the module doc). */
  text: string;
  /** Canonical run properties of each character of `text`. */
  formats: string[];
  /**
   * Which character of the AUTHORED redline each character of `text` is — its
   * run's serial and its offset in that run — and whether that run sat inside
   * a revision (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`). `undefined` for
   * a source document.
   */
  origins: (Origin | undefined)[];
}

interface Origin {
  id: string;
  revised: boolean;
}

/** How one view of a redline compares with the document it should reproduce. */
export interface ViewReport {
  /** Non-blank paragraphs in the document the view should reproduce. */
  expectedParagraphs: number;
  /** Non-blank paragraphs in the view. */
  actualParagraphs: number;
  /**
   * 0-based positions, among the EXPECTED document's non-blank paragraphs, of
   * those no paragraph of the view reproduces character for character
   * (paragraphs are aligned by a longest common subsequence on their text).
   */
  textDiffers: number[];
  /**
   * The subset of `textDiffers` that IS reproduced once every whitespace run
   * (NBSP included) is read as one space. Reported (a known
   * hazard), not by itself tolerated.
   */
  whitespaceOnly: number[];
  /**
   * Reject-all only: the subset of `textDiffers` reproduced once §5.2 has
   * folded both (whitespace, curly quotes, dashes) AND whose every differing
   * character is a fold-set character the redline left UNREVISED, every
   * neighbour of a missing one unrevised too. Such a character is authored
   * once, from the revised side, for both views: where the two documents spell
   * an unchanged stretch differently in a way Compare reads as equal (`the  X`
   * against `the X`, `"Losses"` against `“Losses”`), the author cannot satisfy
   * both — so it is tolerated by `isClean` only where accept-all PROVED that
   * character right (its paragraph read as the revised one exactly). A deleted
   * character that differs is never in here.
   */
  unrevisedOnly: number[];
  /**
   * The characters `unrevisedOnly` excuses, counted across its paragraphs —
   * each a §5.2 fold-set character, printed so the count cannot drift silently.
   */
  excusedCharacters: number;
  /**
   * Positions of expected paragraphs whose characters are reproduced but whose
   * run formatting is not, character for character — not counting
   * `formattingOnly`.
   */
  formattingDiffers: number[];
  /**
   * Reject-all only: paragraphs whose every formatting difference is on a
   * character the redline left UNREVISED and that accept-all PROVED carries
   * the revised document's formatting. Such a character is authored once for
   * both views, so the difference is a formatting change between the two
   * documents on words that did not change — which Compare does not detect
   * (out of scope; spec 08 §12). Reported, not counted as lost.
   */
  formattingOnly: number[];
}

/** Accept-all against the revised document, reject-all against the original. */
export interface ReconstructionReport {
  accept: ViewReport;
  reject: ViewReport;
}

/** Characters a paragraph's runs contribute that are not `<w:t>` text. */
const LEAF_CHARS: Record<string, string> = {
  'w:tab': '\t',
  'w:br': '\n',
  'w:cr': '\n',
  'w:noBreakHyphen': '\u2011',
  'w:softHyphen': '\u00ad',
};

/** Children that carry properties or hidden content, never visible characters. */
const NOT_VISIBLE = new Set([
  'w:pPr', 'w:rPr', 'm:rPr', 'm:ctrlPr', 'w:instrText', 'w:delInstrText', 'mc:Fallback',
]);

/** Attributes the yardstick puts on every authored run before flattening. */
const SERIAL_ATTR = '@_dpRun';
const REVISED_ATTR = '@_dpRevised';

/** The non-blank paragraphs of a body (or any block XML), in document order. */
export function paragraphViews(xml: string): ParaView[] {
  const views: (ParaView | null)[] = [];
  interface RunState {
    format: string;
    serial?: string;
    revised: boolean;
    offset: number;
  }
  const push = (para: ParaView, run: RunState, chars: string): void => {
    // One entry per UTF-16 code unit, as `text` is indexed.
    for (let k = 0; k < chars.length; k++) {
      para.text += chars[k];
      para.formats.push(run.format);
      para.origins.push(run.serial === undefined ? undefined : { id: `${run.serial}:${run.offset}`, revised: run.revised });
      run.offset++;
    }
  };
  const walk = (nodes: readonly PoNode[], para: ParaView | null, run: RunState): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined || NOT_VISIBLE.has(tag)) continue;
      if (tag === 'w:p') {
        const view: ParaView = { text: '', formats: [], origins: [] };
        const slot = views.length;
        views.push(null);
        walk(childrenOf(node), view, { format: '', revised: false, offset: 0 });
        views[slot] = view.text === '' ? null : view;
        continue;
      }
      if (para === null) {
        walk(childrenOf(node), null, run);
        continue;
      }
      if (tag === 'w:r' || tag === 'm:r') {
        const rPr = childrenOf(node).find((c) => tagOf(c) === 'w:rPr');
        const attrs = (node[':@'] ?? {}) as Record<string, unknown>;
        const serial = attrs[SERIAL_ATTR];
        walk(childrenOf(node), para, {
          format: rPr === undefined ? '' : canonicalRunProperties(rPr),
          ...(typeof serial === 'string' ? { serial } : {}),
          revised: attrs[REVISED_ATTR] !== undefined,
          offset: 0,
        });
        continue;
      }
      if (tag === 'w:t' || tag === 'm:t') {
        for (const child of childrenOf(node)) {
          const t = child['#text'];
          if (typeof t === 'string') push(para, run, t);
        }
        continue;
      }
      const leaf = LEAF_CHARS[tag];
      if (leaf !== undefined) {
        push(para, run, leaf);
        continue;
      }
      walk(childrenOf(node), para, run);
    }
  };
  walk(parsePreserveOrder(xml), null, { format: '', revised: false, offset: 0 });
  return views.filter((v): v is ParaView => v !== null);
}

/**
 * `<w:rPr>` in a canonical spelling: its children sorted, each with its
 * attributes sorted and `w:rsid*` dropped, `<w:rPrChange>` (the record of the
 * OLD formatting, not the formatting) left out.
 */
export function canonicalRunProperties(rPr: PoNode): string {
  const one = (node: PoNode): string => {
    const tag = tagOf(node);
    if (tag === undefined) return '';
    const attrs = (node[':@'] ?? {}) as Record<string, unknown>;
    const spelled = Object.keys(attrs)
      .filter((k) => !/^@_w:rsid/.test(k))
      .sort()
      .map((k) => `${k.slice(2)}=${String(attrs[k])}`)
      .join(' ');
    const inner = childrenOf(node).map(one).filter((s) => s !== '').sort().join('');
    return `<${tag}${spelled === '' ? '' : ` ${spelled}`}>${inner}`;
  };
  return childrenOf(rPr)
    .filter((c) => tagOf(c) !== 'w:rPrChange')
    .map(one)
    .filter((s) => s !== '')
    .sort()
    .join('');
}

/**
 * Character pairs `[expected offset, actual offset]` of a minimal alignment of
 * two paragraph texts. Myers' O(ND) diff, by diff-match-patch with no time cap
 * (so exact, no half-match shortcut): a long paragraph with a handful of
 * differences costs a handful of passes. A table LCS cost the product of the
 * two lengths in time and memory — 1.5 s and a 121M-cell table for one
 * 11,000-character paragraph — and this runs inside the author's self-check.
 */
function characterPairs(e: string, a: string): [number, number][] {
  const pairs: [number, number][] = [];
  let i = 0;
  let j = 0;
  for (const [op, text] of DMP.diff_main(e, a, false)) {
    if (op === DiffMatchPatch.DIFF_EQUAL) for (let k = 0; k < text.length; k++) pairs.push([i + k, j + k]);
    if (op !== DiffMatchPatch.DIFF_INSERT) i += text.length;
    if (op !== DiffMatchPatch.DIFF_DELETE) j += text.length;
  }
  return pairs;
}

const DMP = ((): DiffMatchPatch => {
  const dmp = new DiffMatchPatch();
  dmp.Diff_Timeout = 0;
  return dmp;
})();

const collapse = (s: string): string => s.replace(/[\s\u00a0]+/g, ' ').trim();

/**
 * A paragraph as Compare reads it: §5.2's fold, without the characters
 * `collectText` does not read at all (a non-breaking or soft hyphen).
 */
const fold = (s: string): string => normaliseText(s.replace(/[\u2011\u00ad]/g, ''));

/** Alignment of a view with its document: each expected paragraph's counterpart, exact or loose. */
interface Alignment {
  exact: Map<number, number>;
  /** Counterparts equal once §5.2 normalisation has folded both (whitespace, quotes, dashes). */
  loose: Map<number, number>;
}

function align(actual: readonly ParaView[], expected: readonly ParaView[]): Alignment {
  const exact = new Map(computeLCS([...expected], [...actual], (v) => v.text));
  const restE = expected.map((_, i) => i).filter((i) => !exact.has(i));
  const used = new Set(exact.values());
  const restA = actual.map((_, j) => j).filter((j) => !used.has(j));
  const folded = (v: ParaView): string => fold(v.text);
  const loose = new Map(
    computeLCS(restE.map((i) => expected[i]!), restA.map((j) => actual[j]!), folded)
      .map(([k, l]) => [restE[k]!, restA[l]!] as [number, number]),
  );
  return { exact, loose };
}

/**
 * Compare one view of a redline with the paragraphs it should reproduce.
 * `proven` (reject-all only) holds the authored characters accept-all showed
 * in the revised document's own formatting — see `formattingOnly` — and turns
 * on `unrevisedOnly`.
 */
export function compareViews(
  actual: readonly ParaView[], expected: readonly ParaView[], proven?: Proven,
): ViewReport {
  const { exact, loose } = align(actual, expected);
  const formattingDiffers: number[] = [];
  const formattingOnly: number[] = [];
  /** Judge the formatting of character pairs `[expected offset, actual offset]`. */
  const judge = (i: number, e: ParaView, a: ParaView, pairs: readonly [number, number][]): void => {
    let verdict: 'same' | 'unrevised' | 'lost' = 'same';
    for (const [k, l] of pairs) {
      if (e.formats[k] === a.formats[l]) continue;
      const origin = a.origins[l];
      if (proven !== undefined && origin !== undefined && !origin.revised && proven.format.has(origin.id)) {
        verdict = 'unrevised';
        continue;
      }
      verdict = 'lost';
      break;
    }
    if (verdict === 'unrevised') formattingOnly.push(i);
    else if (verdict === 'lost') formattingDiffers.push(i);
  };
  const identity = (n: number): [number, number][] => [...Array(n).keys()].map((k) => [k, k]);
  const textDiffers: number[] = [];
  expected.forEach((e, i) => {
    const j = exact.get(i);
    if (j === undefined) textDiffers.push(i);
    else judge(i, e, actual[j]!, identity(e.text.length));
  });
  const whitespaceOnly: number[] = [];
  const unrevisedOnly: number[] = [];
  let excusedTotal = 0;
  for (const i of textDiffers) {
    const j = loose.get(i);
    if (j === undefined) continue;
    const e = expected[i]!;
    const a = actual[j]!;
    if (collapse(e.text) === collapse(a.text)) whitespaceOnly.push(i);
    const pairs = characterPairs(e.text, a.text);
    judge(i, e, a, pairs);
    const excused = proven === undefined ? null : excusedCharacters(e, a, pairs, proven);
    if (excused !== null) {
      unrevisedOnly.push(i);
      excusedTotal += excused;
    }
  }
  formattingDiffers.sort((x, y) => x - y);
  formattingOnly.sort((x, y) => x - y);
  return {
    expectedParagraphs: expected.length,
    actualParagraphs: actual.length,
    textDiffers,
    whitespaceOnly,
    unrevisedOnly,
    excusedCharacters: excusedTotal,
    formattingDiffers,
    formattingOnly,
  };
}

/**
 * §5.2's fold set: the characters `normaliseText` rewrites or collapses —
 * whitespace (a tab and a break included), the non-breaking spaces, curly and
 * straight quotes and apostrophes, en and em dashes, the hyphen. Only these are
 * ever excused (by design): never another character.
 */
const FOLDED = /^[\s\u00a0\u202f\u2007"\u201c\u201d'\u2018\u2019\u2013\u2014-]$/;

/**
 * How many characters of `e` and `a` differ, when every difference (the gaps
 * between their aligned characters) is a §5.2 fold-set character in UNREVISED
 * text accept-all proved — else `null`. Each character the view has that the
 * document does not must be such a character; a stretch the document has and
 * the view lacks OUTRIGHT (no character of the view in its gap) must sit
 * inside an unchanged stretch, every neighbour unrevised and proved (`the  Buyer`
 * authored `the Buyer`) — never beside a struck or inserted character, where it
 * may be a deleted character lost. A substitution is excused on the shown
 * character's own proof.
 */
function excusedCharacters(e: ParaView, a: ParaView, pairs: readonly [number, number][], proven: Proven): number | null {
  const good = (l: number): boolean => {
    const o = a.origins[l];
    return o !== undefined && !o.revised && proven.text.has(o.id);
  };
  let count = 0;
  let k0 = 0;
  let l0 = 0;
  for (const [k, l] of [...pairs, [e.text.length, a.text.length] as [number, number]]) {
    if (k > k0 || l > l0) {
      for (let x = k0; x < k; x++) if (!FOLDED.test(e.text[x]!)) return null;
      for (let x = l0; x < l; x++) if (!good(x) || !FOLDED.test(a.text[x]!)) return null;
      if (l === l0) {
        const left = l0 > 0 ? good(l0 - 1) : undefined;
        const right = l < a.text.length ? good(l) : undefined;
        if (left === false || right === false || (left === undefined && right === undefined)) return null;
      }
      count += k - k0 + (l - l0);
    }
    k0 = k + 1;
    l0 = l + 1;
  }
  return count;
}

/**
 * Authored characters accept-all PROVED right: `text` — in a paragraph that
 * reads as its revised paragraph character for character; `format` — and in
 * that character's own formatting there.
 */
interface Proven {
  text: ReadonlySet<string>;
  format: ReadonlySet<string>;
}

function provenCharacters(actual: readonly ParaView[], expected: readonly ParaView[]): Proven {
  const text = new Set<string>();
  const format = new Set<string>();
  const { exact } = align(actual, expected);
  for (const [i, j] of exact) {
    const e = expected[i]!;
    const a = actual[j]!;
    a.origins.forEach((o, k) => {
      if (o === undefined) return;
      text.add(o.id);
      if (a.formats[k] === e.formats[k]) format.add(o.id);
    });
  }
  return { text, format };
}

const REVISION_WRAPPERS = new Set(['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']);

/** Number every authored run, and flag every one inside a revision wrapper. */
function tagRuns(xml: string): string {
  let serial = 0;
  const visit = (nodes: PoNode[], inside: boolean): PoNode[] =>
    nodes.map((node) => {
      const tag = tagOf(node);
      if (tag === undefined) return node;
      const value = node[tag];
      const within = inside || REVISION_WRAPPERS.has(tag);
      const children = Array.isArray(value) ? visit(value as PoNode[], within) : value;
      const out: PoNode = { [tag]: children };
      const attrs = node[':@'] as Record<string, unknown> | undefined;
      if (tag === 'w:r' || tag === 'm:r') {
        out[':@'] = { ...attrs, [SERIAL_ATTR]: String(serial++), ...(within ? { [REVISED_ATTR]: '1' } : {}) };
      } else if (attrs !== undefined) {
        out[':@'] = attrs;
      }
      return out;
    });
  return buildNodes(visit(parsePreserveOrder(xml), false));
}

/**
 * Both views of `authoredXml` against the two source bodies: accept-all
 * against `revisedXml`, reject-all against `originalXml`. The sources are
 * flattened by the same policy as their view, which leaves a source with no
 * revisions in it as it is.
 */
export function reconstruct(authoredXml: string, originalXml: string, revisedXml: string): ReconstructionReport {
  const view = (xml: string, policy: 'acceptAll' | 'rejectAll'): ParaView[] =>
    paragraphViews(flattenTrackedChanges(xml, policy));
  const authored = tagRuns(authoredXml);
  const accepted = view(authored, 'acceptAll');
  const revised = view(revisedXml, 'acceptAll');
  return {
    accept: compareViews(accepted, revised),
    reject: compareViews(view(authored, 'rejectAll'), view(originalXml, 'rejectAll'), provenCharacters(accepted, revised)),
  };
}

/** Paragraphs a view gets wrong: text not reproduced (less reject-all's `unrevisedOnly`). */
export function textLost(v: ViewReport): number[] {
  return v.textDiffers.filter((i) => !v.unrevisedOnly.includes(i));
}

/** Is the view clean — every paragraph reproduced, formatting included, and the count equal? */
export function isClean(v: ViewReport): boolean {
  return textLost(v).length === 0 && v.formattingDiffers.length === 0 && v.expectedParagraphs === v.actualParagraphs;
}
