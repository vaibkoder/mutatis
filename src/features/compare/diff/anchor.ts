/**
 * Level-2 word diff · edit-boundary anchoring (ADR-0028).
 * Spec: spec/features/08-compare.md §5.3.
 *
 * `diff_main` returns *a* minimal edit script, not the most readable one. Where
 * several placements of the same edit are equally minimal, diff-match-patch
 * picks whichever its scan reaches first — and because `diff/word.ts` diffs in
 * SENTINEL space (every token mapped to an arbitrary code point), DMP's own
 * boundary-shifting pass (`diff_cleanupSemanticLossless`) is scoring characters
 * that carry no punctuation, whitespace or word information. No boundary
 * decision is ever taken on real text.
 *
 * The classic legal-drafting casualty is an appended proviso:
 *
 *   … ending on or before {+Completion, provided that … paid before +}Completion.
 *
 * which is minimal, and unreadable: the untouched sentence is cut off mid-flow
 * and then finished by a word that looks like part of the new text. The same
 * edit anchored one phrase later reads correctly:
 *
 *   … ending on or before Completion{+, provided that … paid before Completion+}.
 *
 * This pass slides each single-sided edit through identical neighbouring tokens
 * and keeps the placement that reads best. It is **length-preserving**: it never
 * changes which text is inserted or deleted, only where the mark sits, so it
 * cannot drop a change (ADR-0008). `slideEditBoundaries`
 * verifies that invariant on every call and returns its input unchanged if it
 * would ever be broken.
 *
 * Pure: no Office JS, no Date.now(), no Math.random().
 */
import type { CompareConfig } from '@shared/config/index.js';

import type { WordDiff } from '../model.js';

import { atomicKindOf } from './atomic-tokens.js';
import { sideText } from './side-text.js';

/** Sentence- and clause-level punctuation a token may carry as a tail. */
const TRAILING_PUNCT = /[.,;:!?]+$/;

interface Seg {
  op: WordDiff['op'];
  toks: WordDiff[];
}

/** Group per-token WordDiffs into maximal same-op segments. */
function group(diffs: WordDiff[]): Seg[] {
  const segs: Seg[] = [];
  for (const d of diffs) {
    const last = segs[segs.length - 1];
    if (last && last.op === d.op) last.toks.push({ ...d });
    else segs.push({ op: d.op, toks: [{ ...d }] });
  }
  return segs;
}

/**
 * Flatten back to per-token WordDiffs, stamping each token with the op of the
 * SEGMENT it ended up in. A token that slid across a boundary still carries the
 * op it was born with, and that op is what the author and the never-suppress
 * guard read — so the segment has to win, or the boundary moves while the mark
 * stays put.
 */
function flatten(segs: Seg[]): WordDiff[] {
  return segs.flatMap((s) => s.toks.map((t) => (t.op === s.op ? t : { ...t, op: s.op })));
}

function textOf(toks: readonly WordDiff[]): string {
  return toks.map((t) => t.text).join('');
}

/** The first word-ish run in `text` ('' when it opens on punctuation or space). */
function firstWord(text: string): string {
  return /[\p{L}\p{N}]+/u.exec(text)?.[0] ?? '';
}

/**
 * How well does a marked span read where it currently sits? Purely structural —
 * no vocabulary lists, so nothing here is jurisdiction- or language-specific.
 */
function placementScore(prev: string, span: string, next: string): number {
  let score = 0;

  // The tell-tale of a mis-anchored edit: the span OPENS on the same word that
  // reappears immediately after it closes, so the reader cannot tell which
  // occurrence belongs to the untouched sentence.
  const opening = firstWord(span);
  if (opening !== '' && opening === firstWord(next)) score -= 6;

  // A span that opens on punctuation, or straight after it, joins cleanly.
  if (/^\s*[.,;:!?]/.test(span) || /[.,;:!?]\s*$/.test(prev)) score += 3;
  // …and one that closes on punctuation is self-contained.
  if (TRAILING_PUNCT.test(span.trimEnd())) score += 2;

  // Weak tie-breaks: prefer whole-word edges either side.
  if (/\s$/.test(prev) || /^\s/.test(span)) score += 1;
  if (/\s$/.test(span) || /^\s/.test(next)) score += 1;

  return score;
}

/**
 * Split a trailing-punctuation tail off `tok` so a shift can line up with the
 * word it carries — `"Completion."` → `["Completion", "."]`. Returns null when
 * the token must stay whole: no tail, nothing left in front of the tail, or the
 * token is ATOMIC (an amount, date, period or clause reference must never be
 * cut — §5.3, and P5 of the readability principles).
 */
function splitTail(tok: WordDiff, config: CompareConfig): [WordDiff, WordDiff] | null {
  // The token's OWN kind, stamped by the tokenizer — a bare reference value is
  // not recognisable from its text alone, so re-deriving it here would miss it.
  if (tok.atomic !== undefined || atomicKindOf(tok.text, config) !== undefined) return null;
  const tail = TRAILING_PUNCT.exec(tok.text)?.[0];
  if (tail === undefined) return null;
  const head = tok.text.slice(0, tok.text.length - tail.length);
  if (head === '') return null;
  return [{ op: tok.op, text: head }, { op: tok.op, text: tail }];
}

/**
 * Can the edit shift one token to the right? The text leaving the front of the
 * span must be the same text arriving at its back, otherwise the eq segment
 * that absorbs it would be claiming text the other side does not have.
 * `next` is split at a punctuation tail when that is what makes the two line up.
 */
function shiftable(span: Seg, next: Seg, config: CompareConfig): boolean {
  const head = span.toks[0];
  const target = next.toks[0];
  if (!head || !target) return false;
  if (head.text === target.text) return true;

  // The tokenizer glues sentence- and clause-level punctuation onto the word it
  // follows, so `Completion,` and `Completion.` never line up even though the
  // WORD does — which is exactly what pins an appended proviso to the earlier
  // occurrence. Split the tail off whichever side carries one, then re-test.
  const headSplit = splitTail(head, config);
  const targetSplit = splitTail(target, config);
  if ((headSplit?.[0].text ?? head.text) !== (targetSplit?.[0].text ?? target.text)) return false;

  if (headSplit) span.toks.splice(0, 1, headSplit[0], headSplit[1]);
  if (targetSplit) next.toks.splice(0, 1, targetSplit[0], targetSplit[1]);
  return span.toks[0]!.text === next.toks[0]!.text;
}

/**
 * Slide every single-sided edit to its best-reading placement.
 *
 * Runs on the decoded per-token WordDiffs (real text, atomic kinds attached),
 * before coalescing. Only `eq → (all-ins | all-del) → eq` triples are eligible:
 * a two-sided replace has no free boundary to slide, and leaving those alone
 * keeps the atomic ins/del pairing the classifier reads (§5.3) untouched.
 */
export function slideEditBoundaries(diffs: WordDiff[], config: CompareConfig): WordDiff[] {
  if (diffs.length < 3) return diffs;

  const before = { orig: sideText(diffs, 'del'), rev: sideText(diffs, 'ins') };
  const segs = group(diffs);

  for (let i = 1; i < segs.length - 1; i++) {
    const span = segs[i]!;
    const prev = segs[i - 1]!;
    const next = segs[i + 1]!;
    if (span.op === 'eq' || prev.op !== 'eq' || next.op !== 'eq') continue;

    let best = placementScore(textOf(prev.toks), textOf(span.toks), textOf(next.toks));
    let bestState: [WordDiff[], WordDiff[], WordDiff[]] =
      [[...prev.toks], [...span.toks], [...next.toks]];

    while (shiftable(span, next, config)) {
      prev.toks.push(span.toks.shift()!);
      span.toks.push(next.toks.shift()!);
      const score = placementScore(textOf(prev.toks), textOf(span.toks), textOf(next.toks));
      // `>=` prefers the later placement on a tie: appended material reads
      // better anchored after the text it follows than before it.
      if (score >= best) {
        best = score;
        bestState = [[...prev.toks], [...span.toks], [...next.toks]];
      }
    }

    [prev.toks, span.toks, next.toks] = bestState;
  }

  const out = flatten(segs).filter((d) => d.text !== '');

  // Never-suppress guard (ADR-0008): the pass may only move marks, never change
  // what either side says. If that ever fails, keep the un-slid diff.
  if (sideText(out, 'del') !== before.orig || sideText(out, 'ins') !== before.rev) return diffs;
  return out;
}

