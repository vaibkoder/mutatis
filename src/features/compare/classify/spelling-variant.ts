/**
 * Feature 08 — Compare · en-US ↔ en-GB spelling-variant detection (Task B2).
 *
 * A spelling change (`organize → organise`) is a REAL text delta — it can never
 * be filed as formatting (§7.1). This detector recognises a del→ins pair that
 * matches the configured bidirectional variant map so the SG/UK rules can force
 * it to `substantiveText` and tag it with a `spellingVariant` sub-flag.
 *
 * Detection is case-insensitive, but the ORIGINAL casing is preserved in the
 * returned pair (§7.1, key constraint 6).
 *
 * Spec: spec/features/08-compare.md §7.1; CONVENTIONS.md §en-GB Spelling.
 */

import type { CompareConfig } from '@shared/config/index.js';
import type { DiffOp, WordDiff } from '../model.js';

/** Trim to a single word token for map lookup (strip surrounding punctuation/space). */
function cleanWord(text: string): string {
  return text.trim().replace(/^[^A-Za-z]+|[^A-Za-z]+$/g, '');
}

/**
 * Detect a single en-US→en-GB spelling variant in a replace / inline-text op.
 *
 * For each non-equal (del) WordDiff matching a key in `config.spellingVariants`,
 * look for a non-equal (ins) WordDiff whose cleaned text equals the mapped
 * en-GB value (case-insensitive). Returns the first such pair with original
 * casing preserved, else null.
 */
export function detectSpellingVariant(
  op: DiffOp,
  config: CompareConfig,
): { usSpelling: string; gbSpelling: string } | null {
  const wordDiffs = op.wordDiffs ?? [];
  const dels: WordDiff[] = wordDiffs.filter((wd) => wd.op === 'del');
  const ins: WordDiff[] = wordDiffs.filter((wd) => wd.op === 'ins');
  if (dels.length === 0 || ins.length === 0) return null;

  for (const del of dels) {
    const delWord = cleanWord(del.text);
    if (delWord.length === 0) continue;
    const expectedGb = config.spellingVariants[delWord.toLowerCase()];
    if (expectedGb === undefined) continue;
    for (const insDiff of ins) {
      const insWord = cleanWord(insDiff.text);
      if (insWord.length === 0) continue;
      if (insWord.toLowerCase() === expectedGb.toLowerCase()) {
        // Preserve the original casing as it appeared in the diff.
        return { usSpelling: delWord, gbSpelling: insWord };
      }
    }
  }
  return null;
}
