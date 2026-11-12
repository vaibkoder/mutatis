/**
 * Objective readability metrics for the redline corpus.
 * Spec: spec/features/08-compare.md §5.3–5.5. Principles: docs/compare-readability-principles.md.
 *
 * Drives the REAL Text-vs-text Compare path (extractTextRecords →
 * runComparePipeline) and then the REAL RenderCore parser (buildRenderModel), so
 * what is measured is exactly the run/revision structure the paginated preview
 * and the vector PDF paint.
 *
 * The metrics deliberately do NOT reuse the production tokenizer's idea of a word
 * for the redundancy measure: a reader's idea of a word is the yardstick, and the
 * production tokenizer's punctuation gluing is one of the things under test. They
 * do reuse `tokenize` to find atomic values, because "what is atomic" is a
 * jurisdiction/config question the corpus must not re-answer (P5).
 */
import { readFileSync } from 'node:fs';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig, FlattenPolicy } from '@shared/config/index.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { hasReadableContent } from '@features/compare/diff/atomic-tokens.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import {
  buildNumberingContext,
  flattenTableUnits,
  normaliseText,
  preprocessSide,
  stripNumberingPrefix,
} from '@features/compare/preprocess/index.js';
import type { OpType, ParaRecord } from '@features/compare/model.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';
import { tokenize, tokenizeTagged, atomicKindOf } from '@features/compare/diff/index.js';
import { buildRenderModel } from '@shared/render-core/input/index.js';
import type { RcBlock, RcRun, RcRevisionType } from '@shared/render-core/types.js';

import type { ReadabilityCase, TextReadabilityCase, IdealSeg } from './readability-corpus.js';

export const READABILITY_CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);

export type SegKind = IdealSeg['t'];

/** A maximal run of the rendered paragraph sharing one revision state. */
export interface Seg {
  t: SegKind;
  s: string;
}

export interface RenderedPara {
  /** Auto-numbering label RenderCore resolved, if any (empty in Text mode). */
  label: string;
  segs: Seg[];
  text: string;
}

/**
 * One pipeline op, reduced to what ALIGNMENT is measured from.
 *
 * Deliberately read from `result.ops` and not from the painted runs. A struck
 * paragraph does not always render as an all-`del` paragraph — ADR-0035 revises
 * the paragraph MARK too — so counting the rendering measures the renderer,
 * while counting the ops measures the pairing. Reading the ops reproduces the
 * planning probe's counts exactly; reading the rendering can come up short.
 */
export interface OpShape {
  type: OpType;
  /** ORIGINAL paragraphs this op covers (0 for a pure insert). */
  origParas: number;
  /** REVISED paragraphs this op covers (0 for a pure delete). */
  revParas: number;
}

/** What the two sides said before the diff — the yardstick reconstruction checks against. */
export interface RenderedSource {
  /** Original records' text, numbering prefix restored, whitespace-collapsed. */
  originalText: string;
  /** Revised records' text, same treatment. */
  revisedText: string;
  originalParaCount: number;
  revisedParaCount: number;
}

export interface Rendered {
  paras: RenderedPara[];
  /** Every paragraph's visible text, newline-joined (labels included). */
  text: string;
  /** Non-equal pipeline ops as `type/classification`, for triage checks (P10). */
  ops: string[];
  /** EVERY op, `equal` included, in document order (see {@link OpShape}). */
  opShapes: OpShape[];
  source: RenderedSource;
}

const OPEN: Record<RcRevisionType, string> = {
  ins: '{+',
  del: '[-',
  moveTo: '{>',
  moveFrom: '[<',
};
const CLOSE: Record<RcRevisionType, string> = {
  ins: '+}',
  del: '-]',
  moveTo: '>}',
  moveFrom: '<]',
};

/** Collapse runs into maximal same-revision segments. */
function segsOf(runs: readonly RcRun[]): Seg[] {
  const out: Seg[] = [];
  for (const run of runs) {
    const t: SegKind = run.revision?.type ?? 'eq';
    const prev = out[out.length - 1];
    if (prev && prev.t === t) prev.s += run.text;
    else out.push({ t, s: run.text });
  }
  return out.filter((s) => s.s !== '');
}

function collectParas(blocks: readonly RcBlock[], out: RenderedPara[]): void {
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      const segs = segsOf(b.paragraph.runs);
      const label = b.paragraph.numberLabel ?? '';
      out.push({ label, segs, text: label + segs.map((s) => s.s).join('') });
    } else {
      for (const row of b.table.rows) for (const c of row.cells) collectParas(c.blocks, out);
    }
  }
}

/**
 * The two sides of a case as `ParaRecord[]`, whichever the case is built from.
 *
 * `runComparePipeline` has always taken records rather than strings
 * (`pipeline.ts`), so `extractTextRecords` was only ever one adapter and the
 * docx adapter already existed — `extractParts` → `preprocessSide`, the same
 * two calls `runPipeline` makes for Word-vs-Word. There is no new capability
 * here, only the wiring the instrument never had.
 *
 * ADR-0042 added a third argument to that pair — the side's styles + numbering,
 * from which a paragraph's EFFECTIVE list level is resolved. It is passed here
 * for the same reason the rest of this function exists: the instrument must run
 * what `runPipeline` runs, or it measures a pipeline that no user has. Omitting
 * it would silently hide every style-inherited numbering change from the real
 * corpus, which on a firm template is all of them.
 *
 * This is NOT the same as passing `styles.xml` to the RENDERER, which
 * `renderCase` still deliberately does not do (see its doc comment): that would
 * change what is painted without changing a revision boundary. This changes
 * what is COMPARED.
 */
function recordsFor(c: ReadabilityCase): { original: ParaRecord[]; revised: ParaRecord[] } {
  if (c.kind !== 'docx') {
    return { original: extractTextRecords(c.original), revised: extractTextRecords(c.revised) };
  }
  const config = READABILITY_CONFIG;
  const readSide = (path: string, policy: FlattenPolicy): ParaRecord[] => {
    const parts = extractParts(new PizZipBundle(readFileSync(path)));
    return preprocessSide(
      parts.documentXml,
      'document',
      policy,
      config,
      buildNumberingContext(parts.stylesXml, parts.numberingXml),
    ).paras;
  };
  return {
    original: readSide(c.originalPath, config.flattenOriginal),
    revised: readSide(c.revisedPath, config.flattenRevised),
  };
}

/**
 * Render one case through the real pipeline into reader-visible segments.
 *
 * The render model is built from `authoredBodyXml` alone for BOTH kinds of
 * case, deliberately. Passing a real document's `styles.xml` would change what
 * the renderer paints without changing a single revision boundary — the
 * run/revision structure every metric here reads comes from the `w:ins`/`w:del`
 * markup the pipeline authored — and it would make a synthetic case and a real
 * one incomparable for no measurement gain.
 */
export async function renderCase(c: ReadabilityCase): Promise<Rendered> {
  const records = recordsFor(c);
  const result = await runComparePipeline({
    originalParas: records.original,
    revisedParas: records.revised,
    config: READABILITY_CONFIG,
    author: 'Mutatis',
    date: '2026-09-04T00:00:00Z',
  });
  const model = buildRenderModel({ authoredBodyXml: result.authoredBodyXml });
  const paras: RenderedPara[] = [];
  collectParas(model.blocks, paras);
  const span = (range: { start: number; end: number } | null): number =>
    range === null ? 0 : range.end - range.start + 1;
  return {
    paras,
    text: paras.map((p) => p.text).join('\n'),
    ops: result.ops
      .filter((o) => o.type !== 'equal')
      .map((o) => `${o.type}/${o.classification ?? '-'}`),
    opShapes: result.ops.map((o) => ({
      type: o.type,
      origParas: span(o.origRange),
      revParas: span(o.revRange),
    })),
    source: {
      originalText: sourceText(records.original),
      revisedText: sourceText(records.revised),
      originalParaCount: records.original.length,
      revisedParaCount: records.revised.length,
    },
  };
}

/**
 * A side's text as the pipeline received it — the body text of every paragraph,
 * in order.
 *
 * `normText` is already §5.2-normalised and already has its manual numbering
 * prefix stripped, and the rendered side is put through the SAME two production
 * functions before the two are compared (see {@link reconstructSide}). That is
 * what makes the check meaningful rather than merely strict:
 *
 *  - **Normalisation.** §5.2 folds curly quotes, dashes, NBSP and whitespace
 *    runs before paragraphs are matched, but the redline is authored from the
 *    source `rawXml`, so a curly apostrophe survives into the output. Comparing
 *    raw characters would report a reconstruction failure on every document
 *    that uses typographic quotes, which is a property of the normaliser, not a
 *    lost word.
 *  - **Numbering.** The prefix is stripped in preprocess precisely so a
 *    renumbering does not read as an edit, and ADR-0028 restores the REVISED
 *    one at author time for navigation (P9). So a paragraph renumbered 1.1 → 1.2
 *    shows only `1.2`, and the original's number is legitimately not in the
 *    redline. Whether the number survived is `numbered`'s question, not this
 *    one's; this one asks whether a WORD was lost.
 *
 * Both functions are imported, never re-derived — two implementations of one
 * rule is a defect, and this workstream has already
 * been bitten by a yardstick that drifted from the code it measured.
 */
function sourceText(records: readonly ParaRecord[]): string {
  // A Word table unit (ADR-0046) is read cell by cell, as the rendering is.
  return flattenTableUnits(records)
    .map((p) => p.normText)
    .filter((t) => t !== '')
    .join(' ');
}

/** Whitespace runs to one space, trimmed. */
function collapseWhitespace(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

/** Spell a segment list the way the reader sees it: brackets for strike/underline. */
export function formatSegs(segs: readonly Seg[]): string {
  return segs
    .map((s) => (s.t === 'eq' ? s.s : OPEN[s.t] + s.s + CLOSE[s.t]))
    .join('');
}

export function formatRendered(r: Rendered): string[] {
  return r.paras.map((p) => (p.label ? `${p.label}\t` : '') + formatSegs(p.segs));
}

// ── Reader tokenisation ─────────────────────────────────────────────────────

/** Punctuation a drafter reads as its own mark, not as part of the word. */
const LEAD_PUNCT = /^[([“"'‘]+/;
const TRAIL_PUNCT = /[.,;:!?)\]”'’]+$/;

/**
 * Split text into the units a *reader* treats as indivisible: atomic values whole
 * (P5), every other token stripped of leading and trailing punctuation, which
 * becomes its own unit. Word-internal apostrophes and hyphens stay inside the
 * word — `Purchaser's` and `Third-Party` are one word each to a reader, so a
 * rendering that swaps them whole is not borrowing anything (P1).
 */
export function readerTokens(text: string, config: CompareConfig): string[] {
  const out: string[] = [];
  for (const tok of tokenize(text, config)) {
    const kind = atomicKindOf(tok, config);
    if (kind === 'clauseRef') {
      // P5 makes a clause NUMBER indivisible; it does not make the noun in front
      // of it borrowable. To a reader `Clause 9.13` is the word "Clause" plus a
      // value, so a rendering that marks both sides of `Clause 9.13 → Clause
      // 9.14` has borrowed a word. Split reference words and connectives off the
      // value; the value itself stays whole.
      for (const word of tok.split(/\s+/)) if (word !== '') out.push(word);
      continue;
    }
    if (kind !== undefined) {
      // Every other atomic kind is a single value whose unit is part of it —
      // `24 months` and `S$1,250,000` change whole or not at all.
      out.push(tok);
      continue;
    }
    let rest = tok;
    const lead = LEAD_PUNCT.exec(rest)?.[0];
    if (lead !== undefined) {
      out.push(lead);
      rest = rest.slice(lead.length);
    }
    // Trailing punctuation may carry the token's trailing whitespace after it.
    const body = rest.replace(/\s+$/, '');
    const space = rest.slice(body.length);
    const trail = TRAIL_PUNCT.exec(body)?.[0];
    if (trail !== undefined) {
      const head = body.slice(0, body.length - trail.length);
      if (head !== '') out.push(head);
      out.push(trail + space);
    } else if (rest !== '') {
      out.push(rest);
    }
  }
  return out;
}

/** Longest common subsequence of two token arrays, as the shared token texts. */
/**
 * The longest run of tokens that appears CONTIGUOUSLY in both sequences.
 *
 * The cross-paragraph measure needs "verbatim", which `commonTokens` (an LCS)
 * does not give: scattered matches on drafting boilerplate — `the`, `shall`,
 * `this Deed` — let two unrelated paragraphs score highly. An honest
 * whole-paragraph deletion beside an unrelated insertion measured 0.46 that way.
 */
function longestCommonRun(a: readonly string[], b: readonly string[]): string[] {
  let bestAt = 0;
  let best = 0;
  let prev = new Array<number>(b.length + 1).fill(0);
  for (let i = 1; i <= a.length; i++) {
    const row = new Array<number>(b.length + 1).fill(0);
    for (let j = 1; j <= b.length; j++) {
      if (a[i - 1] === b[j - 1]) {
        row[j] = prev[j - 1]! + 1;
        if (row[j]! > best) {
          best = row[j]!;
          bestAt = i;
        }
      }
    }
    prev = row;
  }
  return a.slice(bestAt - best, bestAt);
}

/**
 * The longest run of CHARACTERS appearing contiguously in both strings.
 *
 * This exists because the token-level measures cannot catch a defect in the
 * tokenizer. `readerTokens` derives its yardstick from the production
 * tokenizer, so when that tokenizer emits `faith` and `faith-][+unilateral,` as
 * two tokens, the metric also sees two unrelated tokens, finds nothing shared
 * and scores **0** — on the exact rendering that was reported. Likewise
 * `[-Statement-]{+Statements+}`: two different tokens, nothing shared, 0.
 *
 * A yardstick built out of the thing it measures cannot measure that thing. So
 * this one is built out of characters and knows nothing about tokens.
 */
function longestCommonSubstring(a: string, b: string): string {
  if (a === '' || b === '') return '';
  let best = 0;
  let bestEnd = 0;
  let prev = new Array<number>(b.length + 1).fill(0);
  for (let i = 1; i <= a.length; i++) {
    const row = new Array<number>(b.length + 1).fill(0);
    for (let j = 1; j <= b.length; j++) {
      if (a[i - 1] === b[j - 1]) {
        row[j] = prev[j - 1]! + 1;
        if (row[j]! > best) {
          best = row[j]!;
          bestEnd = i;
        }
      }
    }
    prev = row;
  }
  return a.slice(bestEnd - best, bestEnd);
}

/**
 * Character-level redundancy of one replace block, 0–1: the longest verbatim
 * character run shared by the struck and inserted sides, over the shorter side.
 *
 * A run carrying no letter or digit scores nothing — a shared space or full
 * stop is not text the reader re-reads (the same `hasReadableContent` rule the
 * token measures use, and the one defect both earlier yardsticks shared).
 * There is deliberately no minimum length beyond that: `[-the-]{+they+}` scoring
 * 1.00 is the measure working, because a word gaining a suffix is exactly the
 * defect it was added for.
 */
function charRedundancyOf(del: string, ins: string): number {
  const floor = Math.min(del.length, ins.length);
  if (floor === 0) return 0;
  const run = longestCommonSubstring(del, ins);
  if (!hasReadableContent(run)) return 0;
  return run.length / floor;
}

function commonTokens(a: readonly string[], b: readonly string[]): string[] {
  const n = a.length;
  const m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i]![j] = a[i] === b[j] ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
    }
  }
  const out: string[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      out.push(a[i]!);
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) i++;
    else j++;
  }
  return out;
}

// ── Metrics ─────────────────────────────────────────────────────────────────

export interface ReadabilityMetrics {
  /** Marks the reader's eye lands on: maximal non-equal segments, all paragraphs. */
  marks: number;
  /**
   * Worst P2/P1 violation, 0–1 — text the reader reads twice to learn that
   * nothing happened to it. 0 is clean. Measured two ways, and the worse wins:
   *
   *  - **within a paragraph**: for each struck span sitting beside an inserted
   *    span (a "replace block"), the share of the shorter side that is
   *    word-for-word identical on both sides (an LCS, so it need not be
   *    contiguous);
   *  - **across paragraphs** (ADR-0032): for a paragraph struck whole, the
   *    longest CONTIGUOUS run of it that reappears among the insertions in any
   *    edited paragraph — the "insert it here, strike it there" rendering of a
   *    merge, which is clean on every within-paragraph test and reads terribly.
   *    Contiguity is what separates real duplication from shared drafting
   *    boilerplate; DISTANCE is not, because the paragraph between a merge's two
   *    halves is often itself deleted, putting the re-insertion two paragraphs
   *    away. An earlier version restricted this to adjacent paragraphs and
   *    scored exactly that rendering 0.11 where it should score 0.96.
   *
   * A paragraph rendered as a whole-for-whole swap is exempt from both: P2
   * admits that as the other honest rendering, and ADR-0026 demotes to it on
   * purpose. **Known blind spot:** that exemption also hides a near-merge which
   * `detectMergesAndSplits` declines and ADR-0026 then demotes — the reader can
   * see most of a paragraph twice and this scores 0. Not fixed here: the
   * exemption is principled, and narrowing it is a measurement-design question
   * rather than a bug. See the compare spec's open questions.
   */
  redundancy: number;
  /**
   * The same measure with a CHARACTER yardstick instead of a token one, 0–1:
   * the longest verbatim character run shared by a struck span and the inserted
   * span beside it, over the shorter side.
   *
   * Added because `redundancy` is blind to any duplication the
   * production tokenizer does not express as a shared token, and the tokenizer
   * is one of the things under test. `[-Statement-]{+Statements+}` and
   * `[-Purchaser-]{+Purchaser's agent+}` both score `redundancy` **0** and
   * `charRedundancy` **1.00**.
   *
   * It inherits `redundancy`'s exemptions on purpose — clean swaps and
   * contentless spans — so that the ONLY difference between the two numbers is
   * the yardstick. The shape neither of them can see (a whole paragraph struck
   * beside a whole paragraph inserted) is `wholeParaChurn`'s job, not this one's.
   */
  charRedundancy: number;
  /** P5: no segment boundary falls inside an atomic value. */
  atomicIntact: boolean;
  /**
   * P9: share of the revised text's manual clause numbers still visible, 0–1.
   *
   * `null` for a docx case. The measure parses clause prefixes out of the
   * revised *string*, which a real document does not have — and on these
   * fixtures every number is style-inherited (`document.xml` carries zero
   * `<w:numPr>`), so the check would be vacuously 1 rather than honest. P9 on
   * whole documents is out of this instrument's scope.
   */
  numbered: number | null;
  /**
   * Does the rendering equal the rendering the reviewer asked for (`case.ideal`)?
   * `null` when the case records no ideal — that is most of them, since an ideal
   * was only written down where the reviewer wanted something other than what
   * the pipeline produced. Compared modulo the things a reader cannot see: the
   * restored clause-number prefix (ADR-0028), the quote and dash folding §5.2
   * applies before matching, and whether a space sits inside or outside a mark.
   */
  idealMatch: boolean | null;
}

const NUMBER_PREFIX = /^\s*((?:\d+(?:\.\d+)*)|\([a-z0-9]+\)|\([ivxl]+\))\s/i;

/** Numbering prefixes a reader expects to still see, taken from the revised text. */
export function expectedNumbers(revised: string): string[] {
  const out: string[] = [];
  for (const line of revised.split('\n')) {
    const m = NUMBER_PREFIX.exec(line);
    if (m?.[1] !== undefined) out.push(m[1]);
  }
  return out;
}

/** Is this paragraph rendered as a whole-for-whole swap (P2's other honest form)? */
function isCleanSwap(segs: readonly Seg[]): boolean {
  return segs.length > 0 && !segs.some((s) => s.t === 'eq');
}

function sideText(segs: readonly Seg[], side: 'del' | 'ins'): string {
  const keep: SegKind[] = side === 'del' ? ['eq', 'del', 'moveFrom'] : ['eq', 'ins', 'moveTo'];
  return segs
    .filter((s) => keep.includes(s.t))
    .map((s) => s.s)
    .join('');
}

/** Offsets at which a rendered side changes revision state. */
function boundaryOffsets(segs: readonly Seg[], side: 'del' | 'ins'): number[] {
  const keep: SegKind[] = side === 'del' ? ['eq', 'del', 'moveFrom'] : ['eq', 'ins', 'moveTo'];
  const out: number[] = [];
  let at = 0;
  for (const s of segs) {
    if (!keep.includes(s.t)) continue;
    if (at > 0) out.push(at);
    at += s.s.length;
  }
  return out;
}

function atomicsIntactOnSide(
  segs: readonly Seg[],
  side: 'del' | 'ins',
  config: CompareConfig,
): boolean {
  const text = sideText(segs, side);
  const bounds = new Set(boundaryOffsets(segs, side));
  let at = 0;
  // Read the kind the tokenizer STAMPED. A bare reference value (`9.13`) is not
  // recognisable from its text alone, so re-deriving the kind here would leave
  // clause numbers unprotected and the P5 check quietly green.
  for (const tok of tokenizeTagged(text, config)) {
    if (tok.atomic !== undefined) {
      for (let k = at + 1; k < at + tok.text.length; k++) if (bounds.has(k)) return false;
    }
    at += tok.text.length;
  }
  return true;
}

// ── Ideal conformance ───────────────────────────────────────────────────────

/** The quote/dash folding §5.2 applies before paragraphs are matched. */
function foldPunct(s: string): string {
  return s.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
}

/**
 * Canonical form for comparing two segment sequences as a reader sees them:
 * whitespace at the edge of a mark moves outside it (strikethrough on a space
 * is invisible), adjacent same-state runs merge, empties drop.
 */
function canonicalSegs(segs: readonly Seg[]): Seg[] {
  const flat: Seg[] = [];
  for (const seg of segs) {
    const s = foldPunct(seg.s);
    if (seg.t === 'eq') {
      flat.push({ t: 'eq', s });
      continue;
    }
    const lead = /^\s*/.exec(s)?.[0] ?? '';
    const trail = /\s*$/.exec(s.slice(lead.length))?.[0] ?? '';
    const core = s.slice(lead.length, s.length - trail.length);
    if (lead) flat.push({ t: 'eq', s: lead });
    if (core) flat.push({ t: seg.t, s: core });
    if (trail) flat.push({ t: 'eq', s: trail });
  }
  const out: Seg[] = [];
  for (const seg of flat) {
    if (seg.s === '') continue;
    const last = out[out.length - 1];
    if (last && last.t === seg.t) last.s += seg.s;
    else out.push({ ...seg });
  }
  return out;
}

/**
 * True when the rendering matches `case.ideal`. The rendered first paragraph
 * carries its manual clause number (ADR-0028) and the ideal texts do not, so the
 * prefix is stripped before comparing.
 */
function matchesIdeal(c: TextReadabilityCase, r: Rendered): boolean {
  const got = canonicalSegs(r.paras.flatMap((p) => p.segs));
  const first = got[0];
  if (first?.t === 'eq') {
    const m = NUMBER_PREFIX.exec(first.s);
    if (m) first.s = first.s.slice(m[0].length);
  }
  const want = canonicalSegs((c.ideal ?? []).map((seg) => ({ t: seg.t, s: seg.s })));
  const clean = got.filter((seg) => seg.s !== '');
  return (
    clean.length === want.length &&
    clean.every((seg, i) => seg.t === want[i]!.t && seg.s === want[i]!.s)
  );
}

/**
 * A paragraph struck in its entirety, as `del` rather than as a move. Its words
 * reappearing in an insertion elsewhere is the same P2 failure as an adjacent
 * struck/inserted pair — the reader reads them twice — but no test inside one
 * paragraph can see it. `moveFrom` is exempt by construction: a move renders as
 * a move (P7), and this only ever looks at `del`.
 */
function wholeParaDeletion(p: RenderedPara): string | null {
  if (p.segs.length === 0) return null;
  if (p.segs.some((s) => s.t !== 'del')) return null;
  return p.segs.map((s) => s.s).join('');
}

/**
 * Re-exported so the yardstick and the code it measures cannot drift apart.
 * One definition, in `diff/atomic-tokens.ts`.
 */
export { hasReadableContent };

/** A struck span sitting immediately beside an inserted span. */
function replaceBlocks(segs: readonly Seg[]): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  for (let i = 0; i < segs.length - 1; i++) {
    const a = segs[i]!;
    const b = segs[i + 1]!;
    if (a.t === 'del' && b.t === 'ins') out.push([a.s, b.s]);
    else if (a.t === 'ins' && b.t === 'del') out.push([b.s, a.s]);
  }
  return out;
}

export function measure(c: ReadabilityCase, r: Rendered): ReadabilityMetrics {
  const config = READABILITY_CONFIG;
  let marks = 0;
  let redundancy = 0;
  let charRedundancy = 0;
  let atomicIntact = true;

  for (const p of r.paras) {
    marks += p.segs.filter((s) => s.t !== 'eq').length;
    if (!atomicsIntactOnSide(p.segs, 'del', config)) atomicIntact = false;
    if (!atomicsIntactOnSide(p.segs, 'ins', config)) atomicIntact = false;
    if (isCleanSwap(p.segs)) continue;
    for (const [d, i] of replaceBlocks(p.segs)) {
      if (!hasReadableContent(d)) continue;
      const dTok = readerTokens(d, config);
      const iTok = readerTokens(i, config);
      const shared = commonTokens(dTok, iTok).join('').length;
      const floor = Math.min(d.length, i.length);
      if (floor > 0) redundancy = Math.max(redundancy, shared / floor);
      charRedundancy = Math.max(charRedundancy, charRedundancyOf(d, i));
    }
  }

  // Cross-paragraph P2: a paragraph struck whole whose words are re-inserted
  // elsewhere in the same change. Rendering a paragraph join as "insert the
  // whole sentence here, strike the whole sentence there" scores clean on every
  // within-paragraph test and reads terribly.
  //
  // The shared text must be one CONTIGUOUS run. Written first as a token LCS, it
  // accumulated shared drafting boilerplate (`the`, `shall`, `this Deed`) into a
  // false positive: an honest insertion that merely reuses the document's defined
  // terms, beside an unrelated whole-paragraph deletion, scored 0.47.
  //
  // DISTANCE is deliberately not a limit. An intermediate version looked only at
  // the paragraph before and after, which reads plausibly — until the paragraph
  // between a merge's two halves is itself deleted, putting the re-insertion two
  // paragraphs away. That version scored exactly the rendering this measure
  // exists to catch 0.11 instead of 0.91. Contiguity is what separates real
  // duplication from shared vocabulary; proximity is not.
  // Each of the three decisions here is pinned by a test in
  // `compare.readability.test.ts` that fails when that decision alone is undone.
  // Only insertions inside an EDITED paragraph count. A wholly-inserted
  // paragraph beside a wholly-struck one is the clean whole-for-whole swap
  // ADR-0026 demotes to on purpose, and P2 admits it as the other honest
  // rendering — the same exemption `isCleanSwap` makes within a paragraph.
  const insertedText = r.paras
    .filter((p) => p.segs.some((s) => s.t !== 'ins'))
    .flatMap((p) => p.segs)
    .filter((s) => s.t === 'ins')
    .map((s) => s.s)
    .join(' ');
  if (insertedText !== '') {
    const iTok = readerTokens(insertedText, config);
    for (const para of r.paras) {
      const struck = wholeParaDeletion(para);
      if (struck === null || !hasReadableContent(struck)) continue;
      // A run with no letter or digit in it is not a repeated WORD. Without this,
      // a lone shared space between a struck paragraph and an unrelated
      // two-word insertion scored 0.09 (C34, ADR-0039) — the same exemption the
      // within-paragraph loop above makes with `hasReadableContent(d)`.
      const run = longestCommonRun(readerTokens(struck, config), iTok).join('');
      const shared = hasReadableContent(run) ? run.length : 0;
      const floor = Math.min(struck.length, insertedText.length);
      if (floor > 0) redundancy = Math.max(redundancy, shared / floor);
    }
  }

  // Two fields are coupled to the case being two STRINGS, and both opt out for
  // a real document rather than being given a made-up value.
  let numbered: number | null = null;
  let idealMatch: boolean | null = null;
  if (c.kind !== 'docx') {
    const wanted = expectedNumbers(c.revised);
    const present = wanted.filter((n) => r.text.includes(n)).length;
    numbered = wanted.length === 0 ? 1 : Math.round((present / wanted.length) * 100) / 100;
    idealMatch = c.ideal === undefined ? null : matchesIdeal(c, r);
  }

  return {
    marks,
    redundancy: Math.round(redundancy * 100) / 100,
    charRedundancy: Math.round(charRedundancy * 100) / 100,
    atomicIntact,
    numbered,
    idealMatch,
  };
}

// ── Companion metrics (whole documents) ──────────────────────────────────────

/**
 * The measures a whole document needs and the 30-case corpus does not.
 *
 * `redundancy` is the objective function for a synthetic case because a
 * synthetic case is one editing act: its max IS its score. Over a 26-paragraph
 * contract a max stops discriminating — one bad block pins it near 1.0 — and,
 * worse, `redundancy` is **blind by construction** to the dominant real-document
 * failure. Its within-paragraph half exempts a paragraph with no `eq` segment
 * and its cross-paragraph half ignores insertions drawn from wholly-inserted
 * paragraphs, so a counterpart pair rendered as delete-everything-here plus
 * insert-everything-there scores **0**. The objective function IMPROVES as the
 * alignment degrades. `wholeParaChurn` and `alignmentRate` exist because of
 * that, and they are the only two numbers here that can see it.
 *
 * These are reported on their own scale. Real cases are never summed into the
 * synthetic corpus's total, so its `5.04 → 0.11` history stays comparable
 * (ADR-0038).
 */
export interface CompanionMetrics {
  /**
   * ADR-0008 on real input: `del`+`eq`+`moveFrom` must rebuild the original and
   * `ins`+`eq`+`moveTo` the revised, across the whole document, whitespace
   * collapsed. **A hard assert.** A failure here outranks every other finding in
   * this workstream — it would mean the redline had lost something.
   */
  reconstruction: { original: boolean; revised: boolean };
  /**
   * The alignment-failure signature: whole-paragraph `delete` ops sitting
   * immediately beside a whole-paragraph `insert` op in document order.
   *
   * This is the shape a reader sees as "it deleted my clause and wrote it out
   * again below", and it is precisely what `redundancy` scores 0. Ratchet DOWN.
   *
   * Counted in **paragraphs, not ops**. An op-level count is not monotone under
   * the ratchet: a change that merely merged three adjacent unpaired deletes
   * into one op would drop the figure from 3 to 1 without pairing a single
   * paragraph, and a ratchet on it would read that as progress. Multi-paragraph `delete` ops are reachable —
   * `demoteDissimilarReplaces` copies a widened `origRange` onto the op it
   * demotes (`diff/replace-threshold.ts`) — so this is not hypothetical, even
   * though ordinary document pairs rarely produce one.
   */
  wholeParaChurn: number;
  /**
   * Share of ORIGINAL paragraphs the pipeline matched to a counterpart
   * (`equal` / `replace` / `move` / `property`) rather than deleting outright, 0–1.
   * Ratchet UP. On a document whose every paragraph has a counterpart, a low
   * figure is the collapse in one number.
   */
  alignmentRate: number;
  /**
   * Replace blocks at or above the 0.60 "the reader must diff this himself"
   * floor, on the TOKEN yardstick — the same floor the 4,000-pair generalisation
   * sweep counts, so the two are comparable.
   */
  offendingBlocks: number;
  /** The same, per 1,000 words of the revised document — comparable across pairs. */
  offendingBlocksPerKWord: number;
  /**
   * The same count on the CHARACTER yardstick, and it is not a redundant pair of
   * numbers. On a whole-document pair `offendingBlocks` can be **0** while
   * `charRedundancyMean` is high — the token floor never fires on a
   * defect that lives inside a token. Reporting only the token count would
   * publish a clean sheet for the worst thing these documents show.
   */
  charOffendingBlocks: number;
  /** The character-yardstick count, per 1,000 words of the revised document. */
  charOffendingBlocksPerKWord: number;
  /** Mean token-level redundancy over every scored replace block. */
  redundancyMean: number;
  /** 95th percentile (nearest-rank) of the same distribution. */
  redundancyP95: number;
  /** Mean character-level redundancy over the same blocks. */
  charRedundancyMean: number;
  /** 95th percentile (nearest-rank) of the character-level distribution. */
  charRedundancyP95: number;
  /** How many blocks the two distributions above were computed over. */
  blocks: number;
  /** Words in the revised document — the denominator for the per-1,000 figure. */
  words: number;
}

/** The "the reader must diff this block himself" floor, unchanged from the sweep. */
const OFFENDING_FLOOR = 0.6;

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/** Nearest-rank percentile over an ascending-sorted array. Empty ⇒ 0. */
function percentile(sorted: readonly number[], p: number): number {
  if (sorted.length === 0) return 0;
  const rank = Math.max(1, Math.ceil((p / 100) * sorted.length));
  return sorted[rank - 1]!;
}

/**
 * Reconstruct one side of the whole rendered document — `eq` plus that side's
 * own revision states — whitespace-collapsed, the way {@link sourceText} treats
 * the input. Segments join with nothing inside a paragraph and a newline
 * between paragraphs, because a segment boundary is not a word boundary.
 */
function reconstructSide(r: Rendered, side: 'del' | 'ins'): string {
  const keep: SegKind[] = side === 'del' ? ['eq', 'del', 'moveFrom'] : ['eq', 'ins', 'moveTo'];
  return r.paras
    .map((p) => {
      const raw = p.segs
        .filter((sg) => keep.includes(sg.t))
        .map((sg) => sg.s)
        .join('');
      // The same two production functions preprocess ran over the source, in the
      // same order, so the comparison is "same words", not "same bytes".
      return stripNumberingPrefix(normaliseText(collapseWhitespace(raw))).body;
    })
    .filter((t) => t !== '')
    .join(' ');
}

/** Is this op a whole-paragraph deletion / insertion — i.e. an unpaired paragraph? */
function isWholeParaDelete(op: OpShape): boolean {
  return op.type === 'delete' && op.origParas > 0 && op.revParas === 0;
}

function isWholeParaInsert(op: OpShape): boolean {
  return op.type === 'insert' && op.revParas > 0 && op.origParas === 0;
}

/** Every scored replace block in the rendering, as [struck, inserted] pairs. */
function scoredBlocks(r: Rendered): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  for (const p of r.paras) {
    if (isCleanSwap(p.segs)) continue;
    for (const [d, i] of replaceBlocks(p.segs)) {
      if (!hasReadableContent(d)) continue;
      out.push([d, i]);
    }
  }
  return out;
}

/**
 * Compute the companion metrics for a rendering. Works on any case; only real
 * ones are reported with them, because a one-act synthetic case has nothing for
 * a distribution to describe.
 */
export function companions(r: Rendered): CompanionMetrics {
  const config = READABILITY_CONFIG;

  const tokenScores: number[] = [];
  const charScores: number[] = [];
  for (const [d, i] of scoredBlocks(r)) {
    const floor = Math.min(d.length, i.length);
    const shared = commonTokens(readerTokens(d, config), readerTokens(i, config)).join('').length;
    tokenScores.push(floor > 0 ? shared / floor : 0);
    charScores.push(charRedundancyOf(d, i));
  }
  const sortedTokens = [...tokenScores].sort((a, b) => a - b);
  const sortedChars = [...charScores].sort((a, b) => a - b);
  const mean = (xs: readonly number[]): number =>
    xs.length === 0 ? 0 : xs.reduce((a, b) => a + b, 0) / xs.length;

  let churn = 0;
  for (let k = 0; k < r.opShapes.length; k++) {
    if (!isWholeParaDelete(r.opShapes[k]!)) continue;
    const before = r.opShapes[k - 1];
    const after = r.opShapes[k + 1];
    if ((before && isWholeParaInsert(before)) || (after && isWholeParaInsert(after))) {
      churn += r.opShapes[k]!.origParas;
    }
  }

  let matchedOriginals = 0;
  for (const op of r.opShapes) {
    // `property` (ADR-0042) is a MATCHED pair by construction: the pass converts
    // an `equal` op in place, keeping both ranges, and never re-pairs anything.
    // Counting it as unmatched would report a fall in alignment for a change
    // that made the alignment no different at all — the instrument misreading a
    // new op type as a regression, which is the failure mode ADR-0038 exists to
    // keep out of this file.
    if (
      op.type === 'equal' ||
      op.type === 'replace' ||
      op.type === 'move' ||
      op.type === 'property'
    ) {
      matchedOriginals += op.origParas;
    }
  }

  const words = r.source.revisedText === '' ? 0 : r.source.revisedText.split(/\s+/).length;
  const offending = tokenScores.filter((x) => x >= OFFENDING_FLOOR).length;
  const charOffending = charScores.filter((x) => x >= OFFENDING_FLOOR).length;
  const perKWord = (n: number): number => (words === 0 ? 0 : round2((n * 1000) / words));

  return {
    reconstruction: {
      original: reconstructSide(r, 'del') === r.source.originalText,
      revised: reconstructSide(r, 'ins') === r.source.revisedText,
    },
    wholeParaChurn: churn,
    alignmentRate:
      r.source.originalParaCount === 0
        ? 1
        : round2(matchedOriginals / r.source.originalParaCount),
    offendingBlocks: offending,
    offendingBlocksPerKWord: perKWord(offending),
    charOffendingBlocks: charOffending,
    charOffendingBlocksPerKWord: perKWord(charOffending),
    redundancyMean: round2(mean(tokenScores)),
    redundancyP95: round2(percentile(sortedTokens, 95)),
    charRedundancyMean: round2(mean(charScores)),
    charRedundancyP95: round2(percentile(sortedChars, 95)),
    blocks: tokenScores.length,
    words,
  };
}
