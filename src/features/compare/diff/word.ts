/**
 * Level-2 word diff inside a matched-but-changed (replace) paragraph pair.
 * Spec: spec/features/08-compare.md §5.3.
 *
 * Uses diff-match-patch in "word mode": tokenize both strings (atomic tokens
 * kept whole), map each unique token to a Unicode sentinel char, diff the
 * sentinel strings, then map back to WordDiff[]. Atomic tokens carry their
 * AtomicKind so the classifier can flag amount/date/obligation changes. Pure.
 */
import DiffMatchPatch from 'diff-match-patch';

import type { CompareConfig } from '@shared/config/index.js';

import type { WordDiff, AtomicKind } from '../model.js';

import { tokenizeTagged } from './atomic-tokens.js';
import type { Token } from './atomic-tokens.js';
import { slideEditBoundaries } from './anchor.js';
import { refineReplaceBlocks } from './refine.js';
import { cleanupPreservingRelocations } from './relocate.js';
import type { DmpDiff } from './relocate.js';
import { detectTransposition } from './transpose.js';
import { coalesce } from './side-text.js';

// diff-match-patch operation constants (static on the class).
const DIFF_DELETE = -1;
const DIFF_INSERT = 1;
const DIFF_EQUAL = 0;

/**
 * Compute the word-level diff between two paragraph texts.
 *
 * @returns ordered WordDiff[]; consecutive non-atomic equal tokens are coalesced.
 */
export function wordDiff(
  originalText: string,
  revisedText: string,
  config: CompareConfig,
): WordDiff[] {
  const origTokens = tokenizeTagged(originalText, config);
  const revTokens = tokenizeTagged(revisedText, config);

  // Map unique tokens → sentinel chars (mirrors DMP's linesToChars technique).
  // Keyed on text AND kind: a reference VALUE and an identical-looking ordinary
  // word are different things, and the kind cannot be re-derived downstream.
  const tokenToChar = new Map<string, string>();
  const charToToken: Token[] = [{ text: '' }];

  const encode = (tokens: readonly Token[]): string => {
    let s = '';
    for (const tok of tokens) {
      const key = `${tok.atomic ?? ''}\u0000${tok.text}`;
      let ch = tokenToChar.get(key);
      if (ch === undefined) {
        // Index 0 is reserved (DMP convention); start sentinels at code point 1.
        ch = String.fromCharCode(charToToken.length);
        tokenToChar.set(key, ch);
        charToToken.push(tok);
      }
      s += ch;
    }
    return s;
  };

  const origSentinels = encode(origTokens);
  const revSentinels = encode(revTokens);

  const dmp = new DiffMatchPatch();
  const diffs = dmp.diff_main(origSentinels, revSentinels, false) as DmpDiff[];
  // NB: cleanupSemantic runs in SENTINEL space, so the boundary-shifting half
  // of it (diff_cleanupSemanticLossless) is scoring code points that carry no
  // punctuation or whitespace information. Anchoring is redone on real tokens
  // by slideEditBoundaries below (ADR-0028).
  //
  // It is also where a slot relocation dies: the equality between the two ends
  // of the move is one token, the edits either side are two, so the rule
  // "eliminate an equality no longer than its neighbouring edits" fuses the
  // move into a block that strikes and re-inserts the untouched middle. Nothing
  // downstream can undo that — `refineReplaceBlocks` refuses to strike a word
  // in one mark and re-insert it in another, and `detectTransposition` needs a
  // shape the fusion has erased — so the cleanup is applied everywhere EXCEPT
  // across a verbatim relocation (ADR-0034).
  const cleaned = cleanupPreservingRelocations(
    diffs,
    (ch) => charToToken[ch.charCodeAt(0)] ?? { text: '' },
  );

  // Expand each DMP entry back to per-token WordDiffs, attaching atomic kinds.
  const expanded: WordDiff[] = [];
  for (const [op, sentinels] of cleaned) {
    const kind = dmpOp(op);
    for (const ch of sentinels) {
      const token = charToToken[ch.charCodeAt(0)] ?? { text: '' };
      const entry: WordDiff = { op: kind, text: token.text };
      if (token.atomic !== undefined) entry.atomic = token.atomic;
      expanded.push(entry);
    }
  }

  // Undo cleanupSemantic's over-fusion where a finer rendering earns its marks
  // (ADR-0030), then anchor each single-sided edit where it reads best before
  // coalescing (ADR-0028). Both are length-preserving: they move and re-partition
  // marks, never change what is marked.
  const refined = refineReplaceBlocks(expanded, config);
  const coalesced = coalesce(slideEditBoundaries(refined, config));
  // Finally, re-express an adjacent transposition as one relocation instead of
  // two mutual replacements (ADR-0031). Also length-preserving/self-verifying.
  // It emits the relocated span and the connective separately so each keeps its
  // atomic kind; the second coalesce joins them back into one mark.
  return coalesce(detectTransposition(coalesced, config));
}

/** Map a DMP op constant to a WordDiff op. */
function dmpOp(op: number): WordDiff['op'] {
  if (op === DIFF_INSERT) return 'ins';
  if (op === DIFF_DELETE) return 'del';
  if (op === DIFF_EQUAL) return 'eq';
  // Defensive — DMP only emits the three above.
  return 'eq';
}

/** Re-export for callers needing kind detection on a bare token. */
export type { AtomicKind };
