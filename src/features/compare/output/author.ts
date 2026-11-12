/**
 * Feature 08 — Compare: author the redline `<w:body>` from the ordered op list (C1).
 *
 * Pure functions only — no Office JS, no Date.now(), no Math.random(). The
 * author/date and a monotonic id counter are injected via AuthorContext so the
 * output is fully deterministic (key constraint 1, 5).
 *
 * Spec: spec/features/08-compare.md §5.5.
 */
import { XMLParser } from 'fast-xml-parser';
import type { ChangeClass, DiffOp, ParaRecord, WordDiff } from '../model.js';
import { tableMarksObjects } from '../diff/table-summary.js';
import { wrapDel, wrapIns, wrapMoveFrom, wrapMoveTo, wrapParaRuns, markParagraphMark, matchLeadingPPr, openSelfClosedParagraph } from './revision-markup.js';
import { PARA_SEAM } from '../diff/merge-split.js';
import { authorTableInsert, authorTableDelete, authorTableReplace } from './author-table.js';
import { authorWordTableChange, authorWordTableWhole, type WordTableAuthorDeps } from './author-word-table.js';
import {
  authorObjectAwareRuns, authorObjectChange, authorRangeFromSource, needsStructuredRevision, placesAs,
  reviseParagraphRuns, type RevisionIds, type TextAuthoring,
} from './inline-objects.js';

export interface AuthorContext {
  /** w:author for every emitted revision (from CompareConfig). */
  author: string;
  /** ISO-8601 timestamp injected by the driver (w:date). */
  date: string;
  /** Monotonic unique-id generator — keeps every w:id distinct across the doc. */
  nextId: () => number;
  /**
   * Optional, side-channel only (RenderCore feature 16, Phase 2 B1). When the
   * context is created with `captureClasses`, `authorBody` calls this before each
   * op so that every w:id `nextId()` then mints is tagged with the op's change
   * classes. **Purely observational** — it never alters the emitted XML; the
   * editable-redline output is byte-identical whether or not capture is on.
   */
  beginOp?: (classes: readonly ChangeClass[]) => void;
  /**
   * Populated only when created with `captureClasses`: emitted w:id (as a string,
   * matching the parser's `@_w:id`) → that revision's change classes. Consumed by
   * RenderCore to attach styling annotations (never to filter — ADR-0008).
   */
  classByRevId?: Map<string, readonly ChangeClass[]>;
}

/** Options for {@link makeAuthorContext}. */
export interface AuthorContextOptions {
  /**
   * Capture an emitted-w:id → change-class map as a side channel (Phase 2 B1).
   * Off by default; turning it on does not change the authored XML by one byte.
   */
  captureClasses?: boolean;
}

/** Build an AuthorContext whose ids start at `start` (default 1). */
export function makeAuthorContext(
  author: string,
  date: string,
  start = 1,
  opts: AuthorContextOptions = {},
): AuthorContext {
  let counter = start;
  if (!opts.captureClasses) {
    return { author, date, nextId: () => counter++ };
  }
  // Capture path: tag each minted id with the current op's classes.
  const classByRevId = new Map<string, readonly ChangeClass[]>();
  let current: readonly ChangeClass[] | undefined;
  return {
    author,
    date,
    classByRevId,
    beginOp: (classes) => { current = classes; },
    nextId: () => {
      const id = counter++;
      if (current && current.length) classByRevId.set(String(id), current);
      return id;
    },
  };
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
});

/**
 * Walk the ordered DiffOp list and produce a fresh `<w:body>…</w:body>` string.
 *
 * - 'equal'   → paragraph emitted unchanged (verbatim rawXml; see emitEqual).
 * - 'property'→ same as 'equal' — the text is identical; only a non-text
 *               property differs, and the revised record already carries it.
 *               Except a footnote/endnote reference, picture or object added,
 *               removed or moved: that one is marked (emitObjectChange).
 * - 'insert'  → every run wrapped in <w:ins>.
 * - 'delete'  → every run wrapped in <w:del> (<w:t> → <w:delText>).
 * - 'replace' → run-level <w:ins>/<w:del> interleaved from the WordDiff list.
 * - 'move'    → <w:moveFromRangeStart/End> + <w:moveFrom> at the source op,
 *               <w:moveToRangeStart/End> + <w:moveTo> at the destination op,
 *               sharing w:name = "move-{lowerMoveId}".
 */
export function authorBody(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  ctx: AuthorContext,
): string {
  return `<w:body>${authorBlocks(ops, original, revised, ctx)}</w:body>`;
}

/**
 * The blocks {@link authorBody} wraps — also how a Word table cell's content is
 * authored from that cell's own ops (ADR-0046), so a cell gets exactly the
 * body's rules.
 */
export function authorBlocks(
  ops: DiffOp[],
  original: ParaRecord[],
  revised: ParaRecord[],
  ctx: AuthorContext,
): string {
  const parts: string[] = [];

  for (const op of ops) {
    // Side-channel class capture (Phase 2 B1) — no-op unless the context opted in.
    ctx.beginOp?.(op.classes);
    switch (op.type) {
      case 'equal':
        parts.push(emitEqual(op, revised, original));
        break;
      // ADR-0042 — a `property` op is a matched pair with IDENTICAL text whose
      // style or numbering differs. `emitEqual` already prefers the revised
      // record, which is the one carrying the current style and numbering, so
      // the authored paragraph is exactly what it was before this op type
      // existed: the redline XML is byte-identical and only `ops` and the
      // summary grow. Painting a "Formatted:" mark on it is 08c's half.
      //
      // The one exception (ADR-0045 § Amendment): a footnote or endnote
      // reference, picture or embedded object added, removed or moved is a
      // change to the content, and the redline marks it.
      case 'property': {
        // ADR-0046: a Word table whose cells changed only what text identity
        // does not see is re-assembled, so a cell's footnote or picture change
        // is marked where it is.
        const table = op.tableDiff ? wordTablePair(op, original, revised) : null;
        // Nothing to mark (a style or a number, no object): the revised table
        // as written, exactly as a paragraph's property op is — so a table the
        // model must otherwise author whole is not struck and re-inserted for a
        // change the change list files as formatting.
        if (table !== null && !tableMarksObjects(op.tableDiff!)) {
          parts.push(table[1].rawXml);
          break;
        }
        if (table !== null) {
          parts.push(authorWordTableChange(op, table[0], table[1], ctx, tableDeps(ctx)));
          break;
        }
        parts.push(op.propertyChange?.objects === true
          ? emitObjectChange(op, revised, original, ctx)
          : emitEqual(op, revised, original));
        break;
      }
      case 'insert':
        parts.push(emitInsert(op, revised, ctx));
        break;
      case 'delete':
        parts.push(emitDelete(op, original, ctx));
        break;
      case 'replace':
        parts.push(emitReplace(op, revised, original, ctx));
        break;
      case 'move':
        parts.push(emitMove(op, original, revised, ctx));
        break;
      default: {
        const exhaustive: never = op.type;
        throw new Error(`authorBody: unsupported op type '${String(exhaustive)}'`);
      }
    }
  }

  return parts.join('');
}

/** What the Word table author borrows from this one (ADR-0046). */
function tableDeps(ctx: AuthorContext): WordTableAuthorDeps {
  return {
    authorBlocks: (ops, original, revised) => authorBlocks(ops, original, revised, ctx),
    markParagraph: (paraXml, kind) =>
      needsStructuredRevision(paraXml)
        ? reviseParagraphRuns(paraXml, kind, ctx)
        : kind === 'ins' ? insertRunsInParagraph(paraXml, ctx) : deleteRunsInParagraph(paraXml, ctx),
  };
}

/** The two records of a Word table pair, when `op` pairs two Word table units. */
function wordTablePair(op: DiffOp, original: ParaRecord[], revised: ParaRecord[]): [ParaRecord, ParaRecord] | null {
  const o = op.origRange ? original[op.origRange.start] : undefined;
  const r = op.revRange ? revised[op.revRange.start] : undefined;
  return o?.wordTable && r?.wordTable ? [o, r] : null;
}

// ---------------------------------------------------------------------------
// ParaRecord access — an op range may span one or more paragraphs
// ---------------------------------------------------------------------------

function recordsInRange(
  records: ParaRecord[],
  range: { start: number; end: number } | null,
): ParaRecord[] {
  if (range === null) return [];
  const out: ParaRecord[] = [];
  for (let i = range.start; i <= range.end; i++) {
    const rec = records[i];
    if (rec) out.push(rec);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Whole-paragraph ops (equal / insert / delete)
// ---------------------------------------------------------------------------

function emitEqual(op: DiffOp, revised: ParaRecord[], original: ParaRecord[]): string {
  // Equal paragraphs exist on both sides; prefer the revised record (it carries
  // the current numbering), falling back to original for safety. The paragraph
  // is emitted verbatim from rawXml: the numbering prefix is only stripped from
  // the *comparison* text (normText, §5.0) to keep a renumbering cascade from
  // reading as a change — rawXml still carries the manual prefix in its run text
  // (and auto-numbering has no literal prefix at all, so numberingPrefix is ''),
  // so re-injecting it here would double it (e.g. "(1) (1)  Party").
  const recs = op.revRange ? recordsInRange(revised, op.revRange) : recordsInRange(original, op.origRange);
  return recs.map((rec) => rec.rawXml).join('');
}

/**
 * A text-identical pair whose footnote or endnote references, pictures or
 * embedded objects differ (ADR-0045 § Amendment). Three tries, most faithful
 * first, and never the two answers that would hide the change — the revised
 * paragraph verbatim (the redline silent) or the text rebuild (the objects
 * gone):
 *
 * 1. the revised paragraph as it is, only the objects that differ marked;
 * 2. the object-aware replace over an all-unchanged word diff — for an object
 *    that differs INSIDE a link or text box, which it marks whole (it rebuilds
 *    the words: from their source runs where that reconstructs, else
 *    from text, when their run formatting goes, limit 5). Held to the places
 *    check as well as its own: its census cannot see an object that moved, and
 *    a container whose census matches is kept unmarked;
 * 3. ONE paragraph — the revised paragraph's properties, its runs inserted and
 *    the original's struck. Not a deleted paragraph and an inserted one: a mark
 *    that cannot carry a revision (a section break, ADR-0035) would then be
 *    authored twice.
 */
function emitObjectChange(op: DiffOp, revised: ParaRecord[], original: ParaRecord[], ctx: AuthorContext): string {
  const rev = op.revRange ? revised[op.revRange.start] : undefined;
  const orig = op.origRange ? original[op.origRange.start] : undefined;
  if (rev === undefined || orig === undefined) return emitEqual(op, revised, original);
  // Word writes an empty paragraph self-closed (`<w:p w:rsidR="…"/>`), which has
  // no open and close tags to put runs between.
  const { open, pPr, close } = splitParagraph(openParagraph(rev.rawXml));
  const spliced = authorObjectChange(orig, rev, ctx, (runs) => `${open}${pPr}${runs}${close}`);
  if (spliced !== null) return spliced;
  const unchanged: WordDiff[] = rev.normText === '' ? [] : [{ op: 'eq', text: rev.normText }];
  const wrap = (runs: string, number: string): string => `${open}${pPr}${number}${runs}${close}`;
  // From the words' own runs first, then from the diff's text.
  for (const source of [true, false]) {
    let scratch = 1;
    const authoring = textAuthoring(source);
    const dry = authorObjectAwareRuns(unchanged, orig, rev, { ...ctx, nextId: () => scratch++ }, emitWordDiffRun, wrap, authoring);
    if (dry !== null && placesAs(wrap(dry.runs, dry.prefix), orig, rev)) {
      const kept = authorObjectAwareRuns(unchanged, orig, rev, ctx, emitWordDiffRun, wrap, authoring)!;
      return wrap(kept.runs, kept.prefix);
    }
  }
  const runsOf = (rec: ParaRecord, kind: 'ins' | 'del'): string => {
    const xml = openParagraph(rec.rawXml);
    const revisedXml = needsStructuredRevision(xml)
      ? reviseParagraphRuns(xml, kind, ctx)
      : kind === 'ins' ? insertRunsInParagraph(xml, ctx) : deleteRunsInParagraph(xml, ctx);
    return splitParagraph(revisedXml).inner;
  };
  // Insertion first (ADR-0043), so ids are minted in document order.
  const inserted = runsOf(rev, 'ins');
  return `${open}${pPr}${inserted}${runsOf(orig, 'del')}${close}`;
}

/** `<w:p …/>` → `<w:p …></w:p>`; any other paragraph unchanged. */
function openParagraph(paraXml: string): string {
  return paraXml.replace(/^(\s*)<w:p(\s[^>]*?)?\s*\/>(\s*)$/, '$1<w:p$2></w:p>$3');
}

function emitInsert(op: DiffOp, revised: ParaRecord[], ctx: AuthorContext): string {
  const recs = recordsInRange(revised, op.revRange);
  // ADR-0027: a table-unit record (text-vs-text Markdown) authors a structured
  // <w:tbl> redline; ordinary paragraphs are unchanged.
  return recs
    .map((rec) => {
      if (rec.wordTable) return authorWordTableWhole(rec, 'ins', ctx, tableDeps(ctx));
      if (rec.table) return authorTableInsert(rec.table, ctx);
      // ADR-0045: a link, simple field, equation or text box takes its revision
      // INSIDE it — `<w:ins>` may not wrap a `<w:hyperlink>` at all.
      const wrapped = needsStructuredRevision(rec.rawXml)
        ? reviseParagraphRuns(rec.rawXml, 'ins', ctx)
        : insertRunsInParagraph(rec.rawXml, ctx);
      // ADR-0035, mirror direction: wrapping only the runs means REJECTING the
      // insertion strips the text but leaves the paragraph. A reviewer who
      // rejects the whole redline must get their own document back, not one
      // blank paragraph per clause the other side added.
      return paragraphMarkIsRevisable(rec, revised)
        ? markParagraphMark(wrapped, 'ins', ctx.nextId(), ctx.author, ctx.date)
        : wrapped;
    })
    .join('');
}

function emitDelete(op: DiffOp, original: ParaRecord[], ctx: AuthorContext): string {
  const recs = recordsInRange(original, op.origRange);
  return recs
    .map((rec) => {
      if (rec.wordTable) return authorWordTableWhole(rec, 'del', ctx, tableDeps(ctx));
      if (rec.table) return authorTableDelete(rec.table, ctx);
      const struck = needsStructuredRevision(rec.rawXml)
        ? reviseParagraphRuns(rec.rawXml, 'del', ctx)
        : deleteRunsInParagraph(rec.rawXml, ctx);
      // ADR-0035: striking the runs empties the paragraph but leaves it there.
      // The paragraph MARK has to be struck too, or accepting the deletion
      // leaves a blank paragraph where the clause was.
      return paragraphMarkIsRevisable(rec, original)
        ? markParagraphMark(struck, 'del', ctx.nextId(), ctx.author, ctx.date)
        : struck;
    })
    .join('');
}

/**
 * May this paragraph's own MARK carry a revision along with its runs (ADR-0035)?
 *
 * A whole-paragraph change has to mark the paragraph MARK, not just the runs
 * inside it, or the paragraph outlives the change it was part of:
 *
 *  - accepting a deletion whose mark is unmarked leaves an **empty paragraph**
 *    where the clause was;
 *  - rejecting an insertion whose mark is unmarked leaves an empty paragraph
 *    where the added clause was — one per clause, for a reviewer who rejects
 *    the whole redline and expects their own document back.
 *
 * Two shapes must keep their mark unmarked in **both** directions, because
 * revising it destroys something the change was never asked to touch:
 *
 *  - **a mark carrying `w:sectPr`.** The last paragraph of a section holds that
 *    section's page setup (size, orientation, margins, headers). Removing the
 *    mark merges this section into the next one, so resolving an ordinary clause
 *    change would silently re-paginate the document.
 *  - **the last paragraph of its part.** There is nothing to merge into: a body
 *    must end with a paragraph, and a table cell must contain at least one, so
 *    removing the final mark of either yields a package Word may refuse.
 *
 * In both cases the pre-ADR-0035 behaviour stands — an empty paragraph is left
 * behind. That is a smaller defect than a re-paginated document or an invalid
 * file, and the guard is deliberately the SAME on both sides:
 * one rule, so accept-all and reject-all cannot disagree about which paragraphs
 * survive.
 */
function paragraphMarkIsRevisable(rec: ParaRecord, all: readonly ParaRecord[]): boolean {
  if (/<w:sectPr[\s/>]/.test(rec.rawXml)) return false;
  let lastIndexInPart = -1;
  for (const other of all) {
    if (other.partId === rec.partId && other.index > lastIndexInPart) lastIndexInPart = other.index;
  }
  return rec.index !== lastIndexInPart;
}

/** Wrap every run-level child of the paragraph in a single <w:ins>. */
function insertRunsInParagraph(paraXml: string, ctx: AuthorContext): string {
  const id = ctx.nextId();
  return wrapParaRuns(paraXml, runs => (runs.trim() === '' ? runs : wrapIns(runs, id, ctx.author, ctx.date)));
}

/** Wrap every run-level child of the paragraph in a single <w:del>. */
function deleteRunsInParagraph(paraXml: string, ctx: AuthorContext): string {
  const id = ctx.nextId();
  return wrapParaRuns(paraXml, runs => (runs.trim() === '' ? runs : wrapDel(runs, id, ctx.author, ctx.date)));
}

// ---------------------------------------------------------------------------
// move ops
// ---------------------------------------------------------------------------

/**
 * Emit a move endpoint. A move is a pair of ops sharing `moveTargetId`. We use
 * the deterministic, order-independent name "move-{lowerMoveId}" so both ends
 * agree regardless of which op is authored first (key acceptance: shared w:name).
 */
function emitMove(
  op: DiffOp,
  original: ParaRecord[],
  revised: ParaRecord[],
  ctx: AuthorContext,
): string {
  const name = moveName(op);
  // wrapMoveFrom/To consume THREE consecutive ids each (rangeStart, rangeEnd,
  // run-level move element). We allocate a contiguous block of three per
  // paragraph from the shared counter so every w:id stays globally unique.
  // The source side has an origRange (text removed from Original); the
  // destination side has a revRange (text appearing in Revised).
  if (op.origRange && !op.revRange) {
    const recs = recordsInRange(original, op.origRange);
    return recs.map(rec => wrapMoveFrom(rec.rawXml, allocBlock(ctx, 3), name, ctx.author, ctx.date)).join('');
  }
  const recs = recordsInRange(revised, op.revRange);
  return recs.map(rec => wrapMoveTo(rec.rawXml, allocBlock(ctx, 3), name, ctx.author, ctx.date)).join('');
}

/**
 * Reserve a contiguous block of `n` ids and return the base. The block's ids
 * are base, base+1, … base+n-1 — matching wrapMoveFrom/To's internal use of
 * rangeId, rangeId+1, rangeId+2.
 */
function allocBlock(ctx: AuthorContext, n: number): number {
  const base = ctx.nextId();
  for (let i = 1; i < n; i++) ctx.nextId();
  return base;
}

/** Deterministic shared name derived from the lower of the two move ids. */
function moveName(op: DiffOp): string {
  const self = op.id;
  const other = op.moveTargetId ?? op.id;
  const lower = self < other ? self : other;
  return `move-${lower}`;
}

// ---------------------------------------------------------------------------
// replace ops — run-level interleaving from the WordDiff list
// ---------------------------------------------------------------------------

/**
 * Emit a replace paragraph. We rebuild the paragraph's run content from the
 * Level-2 WordDiff list: eq → plain run, ins → <w:ins>, del → <w:del>. The
 * paragraph open tag + <w:pPr> are taken from the revised record so the
 * formatting/style reflects the new version. Objects that are not text are put
 * back around the rebuilt text (`inline-objects.ts`, ADR-0045) — one-to-one
 * pairs only; a merge/split still rebuilds from text alone.
 */
function emitReplace(
  op: DiffOp,
  revised: ParaRecord[],
  original: ParaRecord[],
  ctx: AuthorContext,
): string {
  const wordDiffs = op.wordDiffs ?? [];
  const revRecs = recordsInRange(revised, op.revRange);
  const origRecs = recordsInRange(original, op.origRange);
  const revRec = revRecs[0];
  const origRec = origRecs[0];

  // ADR-0046: a Word table keeps its XML and is revised inside; one paired
  // with a paragraph is the whole table out and the paragraph in (or back).
  if (revRec?.wordTable && origRec?.wordTable) {
    return authorWordTableChange(op, origRec, revRec, ctx, tableDeps(ctx));
  }
  if (revRec?.wordTable || origRec?.wordTable) {
    // Exactly a delete op then an insert op, paragraph marks included (ADR-0035).
    return emitDelete(op, original, ctx) + emitInsert(op, revised, ctx);
  }
  // ADR-0027: table-aware replace (text-vs-text Markdown).
  if (revRec?.table && origRec?.table) {
    return authorTableReplace(origRec.table, revRec.table, ctx);
  }
  if (revRec?.table || origRec?.table) {
    // A table paired with a paragraph (kind mismatch) → clean delete + insert.
    const del = origRec
      ? origRec.table
        ? authorTableDelete(origRec.table, ctx)
        : deleteRunsInParagraph(origRec.rawXml, ctx)
      : '';
    const ins = revRec
      ? revRec.table
        ? authorTableInsert(revRec.table, ctx)
        : insertRunsInParagraph(revRec.rawXml, ctx)
      : '';
    return del + ins;
  }

  // Merge / split (ADR-0032): a `replace` whose range spans more
  // than one paragraph on one side. The word diff carries a PARA_SEAM token per
  // paragraph mark, so the mark's removal/addition is authored as a tracked,
  // rejectable revision on the mark itself.
  if (revRecs.length > 1 || origRecs.length > 1) {
    return emitMergeSplit(wordDiffs, origRecs, revRecs, ctx);
  }

  const template = revRec ?? origRec;
  if (!template) return '';

  const { open, pPr, close } = splitParagraph(template.rawXml);
  const wrap = (runs: string, number: string): string => `${open}${pPr}${number}${runs}${close}`;
  if (revRec !== undefined && origRec !== undefined) {
    // every character from its own source run — kept and inserted
    // ones from the revised paragraph, struck ones from the original — so the
    // formatting, the typography, the tabs and breaks of what did not change
    // survive, held to a check that both views read as their documents.
    // ADR-0045 next: a footnote reference, link, field, equation, text box or
    // bookmark put back around text rebuilt from the diff. `null` from both —
    // a result that failed its check, or no object for the second — falls
    // through to the text rebuild, which is then exactly what it always was.
    for (const source of [true, false]) {
      const kept = authorObjectAwareRuns(wordDiffs, origRec, revRec, ctx, emitWordDiffRun, wrap, textAuthoring(source));
      if (kept !== null) return wrap(kept.runs, kept.prefix);
    }
  }
  const number = revRec !== undefined && origRec !== undefined
    ? numberRuns(revRec.numberingPrefix, origRec.numberingPrefix, revRec.numberingPrefix === origRec.numberingPrefix, ctx)
    : numberingPrefixRun(template);
  const runs = wordDiffs.map(wd => emitWordDiffRun(wd, ctx)).join('');
  return wrap(runs, number);
}

/** The replace author's two ways with text: from source runs, or from the diff (ADR-0045). */
function textAuthoring(source: boolean): TextAuthoring {
  return { source, number: numberRuns };
}

/**
 * An edited clause's manual number, rebuilt from text (ADR-0028, as amended). The same number on both sides is the revised one,
 * unmarked. A number that changed — `(a)` now `(b)`, or removed, or added — is
 * the revised one inserted and the original struck, insertion first
 * (ADR-0043): emitted unmarked from the revised side, as it was, rejecting
 * every change kept the new number on the old clause and lost the old one.
 */
function numberRuns(revised: string, original: string, same: boolean, ids: RevisionIds): string {
  if (same) return prefixTextRun(revised);
  const ins = prefixTextRun(revised);
  const del = prefixTextRun(original);
  return (ins === '' ? '' : wrapIns(ins, ids.nextId(), ids.author, ids.date)) +
    (del === '' ? '' : wrapDel(del, ids.nextId(), ids.author, ids.date));
}

/** One emitted paragraph of a merge/split: its runs, and the revision on its mark. */
interface SeamGroup {
  diffs: WordDiff[];
  /** Revision carried by the paragraph mark that CLOSES this group (undefined = last). */
  mark?: 'ins' | 'del';
}

/**
 * Cut the word-diff stream at every {@link PARA_SEAM}. A seam may sit inside a
 * coalesced run (`"Claim.¶The"`), so the token is split, not just skipped; the
 * op it was carrying becomes the revision on that paragraph mark.
 */
function splitAtSeams(wordDiffs: readonly WordDiff[]): SeamGroup[] {
  const groups: SeamGroup[] = [{ diffs: [] }];
  const current = (): SeamGroup => groups[groups.length - 1]!;

  for (const wd of wordDiffs) {
    if (!wd.text.includes(PARA_SEAM)) {
      current().diffs.push(wd);
      continue;
    }
    const pieces = wd.text.split(PARA_SEAM);
    for (const [i, piece] of pieces.entries()) {
      if (i > 0) {
        // The seam itself: `del` = the mark is being removed (a merge), `ins` =
        // a mark is being added (a split), `eq` = a break both sides have.
        if (wd.op !== 'eq') current().mark = wd.op;
        groups.push({ diffs: [] });
      }
      if (piece !== '') current().diffs.push({ ...wd, text: piece });
    }
  }
  return groups;
}

/**
 * Author a merge or split as one `<w:p>` per seam group.
 *
 * The FIRST paragraph keeps the surviving (revised) numbering prefix, unmarked,
 * exactly as the one-to-one replace path does (ADR-0028). Every later paragraph
 * carries its own record's prefix wrapped in the SAME revision as the mark
 * before it: on a merge that number really does disappear when the change is
 * accepted, so showing it unmarked would be a fabrication; on a split it is new
 * text, so it is shown inserted.
 *
 * Paragraph properties come from the record each emitted paragraph corresponds
 * to on the many-side, except the last, which takes the revised record's — it
 * is the paragraph that survives the accept.
 */
function emitMergeSplit(
  wordDiffs: readonly WordDiff[],
  origRecs: ParaRecord[],
  revRecs: ParaRecord[],
  ctx: AuthorContext,
): string {
  const many = origRecs.length > revRecs.length ? origRecs : revRecs;
  const lastRev = revRecs[revRecs.length - 1];
  /** The record emitted paragraph `k` takes its properties from. */
  const templateOf = (k: number, isLast: boolean): ParaRecord | undefined =>
    (isLast ? lastRev : many[k]) ?? many[k] ?? origRecs[0];

  // from the records' own runs, as a one-to-one replace is.
  const fromSource = authorRangeFromSource(wordDiffs, origRecs, revRecs, ctx, PARA_SEAM, (p, ids) => {
    const template = templateOf(p.k, p.last);
    if (!template) return '';
    const { open, pPr, close } = splitParagraph(template.rawXml);
    const para = `${open}${pPr}${p.number}${p.runs}${close}`;
    return p.mark ? markParagraphMark(para, p.mark, ids.nextId(), ids.author, ids.date) : para;
  });
  if (fromSource !== null) return fromSource;

  const groups = splitAtSeams(wordDiffs);
  const out: string[] = [];
  for (const [k, group] of groups.entries()) {
    const isLast = k === groups.length - 1;
    const template = templateOf(k, isLast);
    if (!template) continue;

    const { open, pPr, close } = splitParagraph(template.rawXml);
    // k === 0 keeps the surviving number unmarked unless it changed (task 4);
    // later paragraphs' numbers ride the revision of the mark that precedes them.
    let prefix: string;
    if (k === 0) {
      const rev = revRecs[0] ?? template;
      const orig = origRecs[0] ?? template;
      prefix = numberRuns(rev.numberingPrefix, orig.numberingPrefix, rev.numberingPrefix === orig.numberingPrefix, ctx);
    } else {
      prefix = prefixRun(many[k] ?? template, groups[k - 1]?.mark, ctx);
    }
    const runs = group.diffs.map((wd) => emitWordDiffRun(wd, ctx)).join('');
    const para = `${open}${pPr}${prefix}${runs}${close}`;
    out.push(group.mark ? markParagraphMark(para, group.mark, ctx.nextId(), ctx.author, ctx.date) : para);
  }
  return out.join('');
}

/** A numbering-prefix run, optionally wrapped as an insertion or deletion. */
function prefixRun(rec: ParaRecord, mark: 'ins' | 'del' | undefined, ctx: AuthorContext): string {
  const run = numberingPrefixRun(rec);
  if (run === '' || mark === undefined) return run;
  return mark === 'ins'
    ? wrapIns(run, ctx.nextId(), ctx.author, ctx.date)
    : wrapDel(run, ctx.nextId(), ctx.author, ctx.date);
}

/**
 * A clean run carrying the paragraph's MANUAL numbering prefix, for the replace
 * path only (ADR-0028).
 *
 * `emitReplace` discards the template paragraph's own runs and rebuilds the
 * paragraph from `wordDiffs` — and those derive from `normText`, which had the
 * prefix stripped by `stripNumberingPrefix` so that a renumbering cascade would
 * not read as a change (§5.0). Without this run, an edited clause silently
 * loses the very number the reader navigates and quotes by, while untouched,
 * inserted, deleted and moved clauses (which emit `rawXml` wholesale) keep
 * theirs — so a redline's numbering disappears exactly where it is discussed.
 *
 * This is deliberately NOT done in `emitEqual`: that path emits `rawXml`, which
 * already carries the prefix in its run text, and re-injecting there is the
 * doubling bug ("(1) (1)  Party") fixed in 2fd4eb5.
 *
 * The prefix is emitted UNMARKED and taken from the template (the revised
 * record where there is one), matching `emitEqual`'s "prefer the revised
 * numbering" rule: a number that moved because clauses above it shifted is a
 * consequence of that edit, not an amendment of its own (§6.1).
 * Auto-numbered paragraphs have no literal prefix (`numberingPrefix === ''`)
 * and are unaffected — their labels come from the numbering scheme.
 */
function numberingPrefixRun(template: ParaRecord): string {
  return prefixTextRun(template.numberingPrefix);
}

/** The prefix run for a given spelling of the number (ADR-0045 may drop its separator). */
function prefixTextRun(prefix: string): string {
  if (prefix === '') return '';
  return `<w:r><w:t xml:space="preserve">${escapeText(prefix)}</w:t></w:r>`;
}

/** Build a single run for a WordDiff token, wrapped per its op. */
function emitWordDiffRun(wd: WordDiff, ctx: AuthorContext): string {
  if (wd.text === '') return '';
  const needsPreserve = /^\s|\s$/.test(wd.text);
  const spaceAttr = needsPreserve ? ' xml:space="preserve"' : '';
  switch (wd.op) {
    case 'eq':
      return `<w:r><w:t${spaceAttr}>${escapeText(wd.text)}</w:t></w:r>`;
    case 'ins':
      return wrapIns(`<w:r><w:t${spaceAttr}>${escapeText(wd.text)}</w:t></w:r>`, ctx.nextId(), ctx.author, ctx.date);
    case 'del':
      // wrapDel rewrites <w:t> → <w:delText>.
      return wrapDel(`<w:r><w:t${spaceAttr}>${escapeText(wd.text)}</w:t></w:r>`, ctx.nextId(), ctx.author, ctx.date);
    default: {
      const exhaustive: never = wd.op;
      throw new Error(`emitWordDiffRun: unsupported WordDiff op '${String(exhaustive)}'`);
    }
  }
}

// ---------------------------------------------------------------------------
// Paragraph splitting (open tag / pPr / inner / close)
// ---------------------------------------------------------------------------

interface SplitParagraph {
  open: string;
  pPr: string;
  inner: string;
  close: string;
}

/**
 * Split a `<w:p …>…</w:p>` string into its open tag, optional `<w:pPr>`, inner
 * run content, and close tag. Uses fast-xml-parser only to *validate* it is a
 * single paragraph; the actual slicing is string-based to preserve byte fidelity
 * (key constraint 6).
 */
export function splitParagraph(xml: string): SplitParagraph {
  // A self-closed `<w:p/>` (ADR-0046: the blank paragraph a cell ends with) is
  // opened first, so revised runs land INSIDE it.
  const paraXml = openSelfClosedParagraph(xml);
  const openMatch = paraXml.match(/^(\s*<w:p(?:\s[^>]*)?>)/);
  const closeMatch = paraXml.match(/(<\/w:p>\s*)$/);
  if (!openMatch || !closeMatch) {
    return { open: '', pPr: '', inner: paraXml, close: '' };
  }
  const open = openMatch[1]!;
  const close = closeMatch[1]!;
  let inner = paraXml.slice(open.length, paraXml.length - close.length);
  let pPr = '';
  const pPrMatch = matchLeadingPPr(inner);
  if (pPrMatch !== null) {
    pPr = pPrMatch;
    inner = inner.slice(pPr.length);
  }
  return { open, pPr, inner, close };
}

/** Parse-validate a paragraph fragment; exported for callers that need a tree. */
export function parseParagraph(paraXml: string): unknown {
  return parser.parse(paraXml);
}

// ---------------------------------------------------------------------------
// Text escaping
// ---------------------------------------------------------------------------

function escapeText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
