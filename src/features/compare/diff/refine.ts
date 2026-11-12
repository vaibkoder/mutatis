/**
 * Level-2 word diff · replace-block refinement.
 * Spec: spec/features/08-compare.md §5.3. ADR-0030.
 *
 * `diff_cleanupSemantic` fuses tight edits into coarse blocks, because in the
 * SENTINEL space `diff/word.ts` diffs in, its "eliminate an equality smaller
 * than or equal to the edits on both sides of it" rule measures the equality
 * with `String.length` — and there one character is one TOKEN. The tokenizer
 * emits every run of whitespace as a token of its own, so `all ` counts 2 and
 * ` ` counts 1 against a six-token deletion: real shared prose is out-weighed
 * and destroyed, the neighbouring edits grow by what they swallowed, and the
 * next equality along becomes eligible in turn. The avalanche is what turns a
 * clause with three small deletions into one whole-paragraph swap:
 *
 *   bad   … indemnify {+the Purchaser against all Losses.+}[-and hold harmless
 *                       the Purchaser against any and all Losses whatsoever.-]
 *   good  … indemnify [-and hold harmless -]the Purchaser against [-any and -]all
 *                       Losses[- whatsoever-].
 *
 * This pass takes each struck span sitting immediately beside an inserted one
 * and re-diffs the two sides on REAL tokens, keeping the finer rendering only
 * when it earns its extra marks:
 *
 *  - it must recover a shared **word**. Recovering nothing but whitespace buys
 *    the reader nothing and costs two more marks: `{+from each+}[-out of any-]`
 *    is one rewritten phrase, and splitting it at the space it happens to share
 *    into `{+from+}[-out of-] {+each+}[-any-]` reads as two acts, not one (P3).
 *  - it must not hand the reader the tool's job. A word struck inside one mark
 *    and re-inserted inside another is text the reader has to prove identical
 *    before they can dismiss it — the P2 failure that makes fragmenting a
 *    replace worse than leaving it coarse.
 *
 * It is **length-preserving**: it re-partitions the same two texts, so nothing
 * can leave the redline (ADR-0008), and it verifies that on
 * every call, returning its input unchanged if it ever failed.
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */
import DiffMatchPatch from 'diff-match-patch';

import type { CompareConfig } from '@shared/config/index.js';

import type { WordDiff } from '../model.js';

import { hasReadableContent } from './atomic-tokens.js';
import { atomicKindOf, type Token } from './atomic-tokens.js';
import { sideText } from './side-text.js';

const DIFF_DELETE = -1;
const DIFF_INSERT = 1;

/** Sentence- and clause-level punctuation a token may carry as a tail. */
const TRAILING_PUNCT = /[.,;:!?]+$/;

/** True when the token carries no letter or digit — whitespace or punctuation. */
function isNoise(text: string): boolean {
  return !hasReadableContent(text);
}

/**
 * Split a trailing-punctuation tail off each token so the word it carries can
 * line up with the same word on the other side — `Losses.` → `Losses` + `.`.
 * The tokenizer glues a full stop onto the word it follows, which is why
 * `Losses` and `Losses.` can never match and why the last word of a shortened
 * sentence gets dragged into the mark. Atomic tokens are never cut: an amount,
 * date, period or reference value changes whole or not at all (§5.3, P5). The
 * split is local to this block — the global tokenizer every classifier rule
 * reads is untouched, exactly as in `anchor.ts` (ADR-0028).
 */
function splitTails(toks: readonly WordDiff[], config: CompareConfig): Token[] {
  const out: Token[] = [];
  for (const tok of toks) {
    // The kind travels with the token: a bare reference value (`9.13`) is not
    // recognisable out of context, so it cannot be re-derived here (ADR-0029).
    const isAtomic = tok.atomic !== undefined || atomicKindOf(tok.text, config) !== undefined;
    const tail = isAtomic ? undefined : TRAILING_PUNCT.exec(tok.text)?.[0];
    const head = tail === undefined ? tok.text : tok.text.slice(0, tok.text.length - tail.length);
    if (tail === undefined || head === '') out.push(asToken(tok));
    else out.push({ text: head }, { text: tail });
  }
  return out;
}

/** A WordDiff seen as a tokenizer token, keeping any atomic kind it carries. */
function asToken(d: WordDiff): Token {
  return d.atomic === undefined ? { text: d.text } : { text: d.text, atomic: d.atomic };
}

/**
 * Sentinel key: atomic kind + text, so a reference value never aligns with an
 * identical-looking ordinary word (ADR-0029). Matches `diff/word.ts`.
 */
function keyOf(tok: Token): string {
  return `${tok.atomic ?? ''}\u0000${tok.text}`;
}

/** One candidate rendering of a replace block, as [op, tokens] runs. */
type Run = [WordDiff['op'], Token[]];

/** Re-diff two token arrays in sentinel space, with no cleanup pass of any kind. */
function tokenDiff(a: readonly Token[], b: readonly Token[]): Run[] {
  const tokenToChar = new Map<string, string>();
  const charToToken: Token[] = [{ text: '' }];
  const encode = (toks: readonly Token[]): string => {
    let s = '';
    for (const tok of toks) {
      const key = keyOf(tok);
      let ch = tokenToChar.get(key);
      if (ch === undefined) {
        ch = String.fromCharCode(charToToken.length);
        tokenToChar.set(key, ch);
        charToToken.push(tok);
      }
      s += ch;
    }
    return s;
  };
  const dmp = new DiffMatchPatch();
  return dmp.diff_main(encode(a), encode(b), false).map(([op, s]): Run => [
    op === DIFF_INSERT ? 'ins' : op === DIFF_DELETE ? 'del' : 'eq',
    [...s].map((ch) => charToToken[ch.charCodeAt(0)] ?? { text: '' }),
  ]);
}

/**
 * The other honest rendering of a replace: the deletion ENDS on the words the
 * insertion BEGINS with, so that shared run belongs to neither mark. DMP has
 * this move (`diff_commonOverlap`) but gates it on a character-count ratio that
 * a token-space diff never reaches. Returns null when there is no overlap.
 */
function overlapSplit(del: readonly Token[], ins: readonly Token[]): Run[] | null {
  const max = Math.min(del.length, ins.length) - 1;
  let k = 0;
  for (let n = max; n > 0; n--) {
    const tail = del.slice(del.length - n);
    if (tail.every((t, j) => ins[j] !== undefined && keyOf(t) === keyOf(ins[j]!))) {
      k = n;
      break;
    }
  }
  if (k === 0) return null;
  return [
    ['del', del.slice(0, del.length - k)],
    ['eq', del.slice(del.length - k)],
    ['ins', ins.slice(k)],
  ];
}

/**
 * Does this candidate hand the reader the tool's job? A word struck inside one
 * mark and re-inserted inside another is text the reader must prove identical
 * before they can dismiss it (P2). Whitespace and punctuation do not count:
 * they carry no word for the reader to chase.
 */
function selfDiffing(runs: readonly Run[]): boolean {
  const deleted = new Set<string>();
  const inserted = new Set<string>();
  for (const [op, toks] of runs) {
    for (const t of toks) {
      if (isNoise(t.text)) continue;
      if (op === 'del') deleted.add(keyOf(t));
      else if (op === 'ins') inserted.add(keyOf(t));
    }
  }
  for (const t of deleted) if (inserted.has(t)) return true;
  return false;
}

/** Accept a candidate only if it recovers a shared WORD and reads honestly. */
function acceptable(runs: readonly Run[] | null): runs is Run[] {
  if (runs === null) return false;
  const recoversWord = runs.some(
    ([op, toks]) => op === 'eq' && toks.some((t) => !isNoise(t.text)),
  );
  return recoversWord && !selfDiffing(runs);
}

/** Characters the tokenizer keeps inside one word run. Mirrors `atomic-tokens.ts`. */
const WORD_CHAR = /[A-Za-z0-9'&.\-]/;

/** Concatenated text of a token run. */
function textOf(toks: readonly Token[]): string {
  return toks.map((t) => t.text).join('');
}

/**
 * Split a token run at a character offset. Returns null when the offset falls
 * strictly inside an ATOMIC token — an amount, date, period or reference value
 * changes whole or not at all (P5), and no readability gain is worth cutting
 * one in half.
 */
function splitTokensAt(toks: readonly Token[], offset: number): [Token[], Token[]] | null {
  const before: Token[] = [];
  const after: Token[] = [];
  let at = 0;
  for (const tok of toks) {
    const end = at + tok.text.length;
    if (end <= offset) before.push(tok);
    else if (at >= offset) after.push(tok);
    else {
      // The boundary falls inside this token.
      if (tok.atomic !== undefined) return null;
      before.push({ text: tok.text.slice(0, offset - at) });
      after.push({ text: tok.text.slice(offset - at) });
    }
    at = end;
  }
  return [before, after];
}

/**
 * Is the extra material on the longer side a new word, or merely the rest of
 * the same one?
 *
 * This is the guard that separates the defect from a rendering a human reviewer
 * has already graded **good**. `Loss` → `Losses` and `Statement` → `Statements`
 * are one word changing its ending, and corpus case C02 records the verdict in
 * so many words: *"Whole word Loss/Losses swapped, not `Loss` + inserted
 * `es`."* Rendering those as `Loss{+es+}` would be the character-level
 * interleave P5 exists to forbid, bought with a metric.
 *
 * A common drafting shape is a different one: a word struck and
 * re-inserted **verbatim** because a clause was appended after it —
 * `[-faith-]{+faith-][+and honest+]+}`,
 * `[-Purchaser-]{+Purchaser's agent+}`. There the extra material is new words,
 * and the whole struck side is read twice for nothing.
 *
 * Those two, and the fixed-point example in `refineReplaceBlocks` below, are
 * written in the order THIS pass sees and emits — it runs inside `wordDiff`,
 * long before stage 6.5 puts the insertion first (ADR-0043). The example
 * predates that stage; what the reader is shown today for the second one is
 * `{+Purchaser's agent+}[-Purchaser-]`.
 *
 * So: the remainder must contain something the tokenizer would not keep inside
 * the word it is growing — whitespace, or any character outside the word-run
 * class.
 */
function addsNewWords(remainder: string): boolean {
  for (const ch of remainder) if (/\s/.test(ch) || !WORD_CHAR.test(ch)) return true;
  return false;
}

/**
 * Letters and digits — the characters that, to a reader, CONTINUE the word in
 * front of them. Deliberately narrower than {@link WORD_CHAR}: the tokenizer
 * keeps `'`, `-`, `.` and `&` inside one word run, but a reader seeing
 * `Purchaser{+'s agent+}` reads two things, not a word cut in half.
 */
const WORD_BODY = /[\p{L}\p{N}]/u;

/**
 * Does the cut fall BETWEEN words, or inside one?
 *
 * {@link addsNewWords} asks whether the extra material contains new words. It
 * scans the whole remainder, so ONE space anywhere licenses the cut — and that
 * is not enough. `Loss` → `Losses and Damages` has a space, but the cut still
 * lands after `Loss`, mid-way through `Losses`, and renders
 * `Loss{+es and Damages+}`: the character-level interleave the sibling guard's
 * own rationale forbids, reached from the other side. Same shape:
 * `Vendor` → `Vendorial agent`, `fee` → `fees of the adviser`, and mirrored,
 * `holder` → `the Shareholder` rendering `{+the Share+}holder`.
 *
 * A word is cut in half exactly when the characters on BOTH sides of the cut
 * continue a word. Testing only one side would be wrong in the other
 * direction: `pay the ` → `pay the reasonable ` cuts after a space and is
 * perfectly readable.
 *
 * Both guards are load-bearing and neither implies the other —
 * `Purchaser` → `Purchaser's` passes this one and is refused by
 * `addsNewWords`; `Loss` → `Losses and Damages` passes `addsNewWords` and is
 * refused here — so each is ablated separately.
 */
function splitsBetweenWords(short: string, remainder: string, atStart: boolean): boolean {
  const [left, right] = atStart
    ? [short[short.length - 1], remainder[0]]
    : [remainder[remainder.length - 1], short[0]];
  if (left === undefined || right === undefined) return true;
  return !(WORD_BODY.test(left) && WORD_BODY.test(right));
}

/**
 * The third branch of the acceptance rule: recover a shared character AFFIX.
 *
 * `recoversWord`, `selfDiffing` and `overlapSplit` all compare token identity
 * through `keyOf`, so they cannot see that `faith` and `faith-][+unilateral,`
 * share five characters — the two are simply different tokens. The re-diff
 * therefore recovers nothing and is declined, `overlapSplit` has no whole token
 * to pull out, and the block is left coarse, which shows the duplication
 * anyway. Both existing branches produce visible duplication and there was no
 * third option (ADR-0040).
 *
 * This one fires only in the narrow case where one side's text is wholly a
 * prefix or a suffix of the other's — so the recovered run is never a fragment
 * of a word on the shorter side, and one of the two marks disappears entirely
 * instead of being split in two.
 */
function affixSplit(del: readonly Token[], ins: readonly Token[]): Run[] | null {
  const dText = textOf(del);
  const iText = textOf(ins);
  if (dText === '' || iText === '' || dText === iText) return null;

  // Which side is wholly contained in the other, and at which end.
  const shortIsDel = dText.length < iText.length;
  const short = shortIsDel ? dText : iText;
  const longToks = shortIsDel ? ins : del;
  const longText = shortIsDel ? iText : dText;
  const longOp: WordDiff['op'] = shortIsDel ? 'ins' : 'del';

  let atStart: boolean;
  if (longText.startsWith(short)) atStart = true;
  else if (longText.endsWith(short)) atStart = false;
  else return null;

  // The recovered run must be text the reader can see; the material left over
  // must be new WORDS rather than the rest of the same word (`addsNewWords`);
  // and the boundary itself must fall between words rather than inside one
  // (`splitsBetweenWords`). Both guards are needed — see each one's docstring.
  if (!hasReadableContent(short)) return null;
  const remainder = atStart ? longText.slice(short.length) : longText.slice(0, longText.length - short.length);
  if (!hasReadableContent(remainder)) return null;
  if (!addsNewWords(remainder)) return null;
  if (!splitsBetweenWords(short, remainder, atStart)) return null;

  const split = splitTokensAt(longToks, atStart ? short.length : longText.length - short.length);
  if (split === null) return null;
  const [head, tail] = split;

  // The equality is taken from the SHORT side's own tokens, so any atomic kind
  // it carries survives; the long side contributes only the new material.
  //
  // There was a second, belt-and-braces re-check here — `atomicKindOf` over the
  // short side, in case a token arrived untagged but was atomic by the config's
  // own reckoning. It was removed as unreachable: every token in a block either
  // comes from `tokenizeTagged` already carrying its kind, or from `splitTails`,
  // which only ever splits a token `atomicKindOf` has already declined. Measured
  // rather than argued — 16,339 `affixSplit` calls over a 20,000-pair fuzz and
  // the whole suite, fired **zero** times,
  // while `splitTokensAt`'s live atomic refusal fired 1,068. A guard that cannot
  // fail cannot be ablated, and an un-ablatable guard is decoration that reads
  // like protection.
  const shortToks = shortIsDel ? del : ins;

  const kept: Token[] = [...shortToks];
  return atStart ? [['eq', kept], [longOp, tail]] : [[longOp, head], ['eq', kept]];
}

/**
 * Re-label the chosen partition as WordDiffs. Each token keeps the kind it
 * arrived with and is never given one here: a bare value (`9.13`) is not
 * recognisable out of context, so re-deriving would stamp a kind the tokenizer
 * deliberately withheld (§5.3, ADR-0029). `atomicKindOf` appears in this module
 * only as a refuse-to-split guard, never as a source of truth.
 */
function toWordDiffs(runs: readonly Run[]): WordDiff[] {
  const out: WordDiff[] = [];
  for (const [op, toks] of runs) {
    for (const tok of toks) {
      if (tok.text === '') continue;
      out.push(tok.atomic === undefined ? { op, text: tok.text } : { op, text: tok.text, atomic: tok.atomic });
    }
  }
  return out;
}

/**
 * Refine every replace block in a decoded per-token word diff.
 *
 * Runs before {@link slideEditBoundaries} so the deletions and insertions it
 * frees are single-sided by the time anchoring sees them — anchoring is also
 * what splits a punctuation tail off a freed boundary word (ADR-0028), so this
 * pass does not need to do that itself.
 */
export function refineReplaceBlocks(diffs: WordDiff[], config: CompareConfig): WordDiff[] {
  if (diffs.length < 2) return diffs;

  const before = { orig: sideText(diffs, 'del'), rev: sideText(diffs, 'ins') };

  // One pass is not enough, and the reason is the defect ADR-0040 exists for.
  // When the re-diff IS accepted it can leave a NEW replace block inside its own
  // output: for example `[-good faith-]{+[-good faith-][+and honest+]+}` refines to `[-` · `good ` · `[-faith-]{+faith-][+…+}`,
  // recovering `good` and leaving `faith` struck beside `faith-][+…`. The
  // residual block is exactly the shape the affix branch answers, and a single
  // pass never looks at it. So the pass runs to a fixed point.
  //
  // Each iteration applies the same acceptance rules to a strictly smaller
  // block, so this terminates on its own; the cap is a guard against a future
  // rule that oscillates, not an expected limit.
  let out = diffs;
  for (let pass = 0; pass < MAX_REFINE_PASSES; pass++) {
    const next = refineOnce(out, config);
    if (sameDiffs(next, out)) break;
    out = next;
  }

  // Never-suppress guard (ADR-0008): the pass may only re-partition the same two
  // texts. If that ever failed, keep the un-refined diff.
  if (sideText(out, 'del') !== before.orig || sideText(out, 'ins') !== before.rev) return diffs;
  return out;
}

/** How many times {@link refineOnce} may re-enter its own output. */
const MAX_REFINE_PASSES = 4;

/** Two diff streams are the same rendering when every op and text matches. */
function sameDiffs(a: readonly WordDiff[], b: readonly WordDiff[]): boolean {
  return (
    a.length === b.length && a.every((d, i) => d.op === b[i]!.op && d.text === b[i]!.text)
  );
}

/** One refinement pass over every replace block in the stream. */
function refineOnce(diffs: WordDiff[], config: CompareConfig): WordDiff[] {

  // Maximal same-op runs; a `del` run touching an `ins` run is a replace block.
  const segs: WordDiff[][] = [];
  for (const d of diffs) {
    const last = segs[segs.length - 1];
    if (last && last[0]!.op === d.op) last.push(d);
    else segs.push([d]);
  }

  const out: WordDiff[] = [];
  for (let i = 0; i < segs.length; i++) {
    const a = segs[i]!;
    const b = segs[i + 1];
    const isBlock =
      b !== undefined &&
      ((a[0]!.op === 'del' && b[0]!.op === 'ins') || (a[0]!.op === 'ins' && b[0]!.op === 'del'));
    if (isBlock) {
      const del = splitTails(a[0]!.op === 'del' ? a : b, config);
      const ins = splitTails(a[0]!.op === 'del' ? b : a, config);
      // First choice: the honest fine-grained diff of the two sides. Fallback:
      // pull out only the run the deletion ends on and the insertion begins
      // with — the placement `diff_main` cannot reach on its own, because both
      // alignments are equally minimal and it keeps whichever its scan met
      // first, which is the one where the shared word sits inside both marks.
      const candidate = tokenDiff(del, ins);
      // Third choice: a shared character AFFIX (ADR-0040). Both branches above
      // reason over token identity, so a word struck and re-inserted verbatim
      // because new material was appended after it reads to them as two
      // unrelated tokens — accepted or declined, either way the reader sees it
      // twice.
      const chosen = acceptable(candidate)
        ? candidate
        : (overlapSplit(del, ins) ?? affixSplit(del, ins));
      if (acceptable(chosen)) {
        out.push(...toWordDiffs(chosen));
        i++;
        continue;
      }
    }
    out.push(...a);
  }

  return out;
}
