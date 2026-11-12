/**
 * The three pairing signals (ADR-0037; spec §6.3).
 *
 * In priority order, and the priority is the design:
 *
 *  1. **Content** — Jaccard over paragraph hashes. Dominant, because two
 *     versions of one agreement share most of their paragraphs verbatim and
 *     two different agreements share almost none, whatever they are called.
 *  2. **Filename** — Dice over character trigrams, after version tokens are
 *     stripped. Secondary, because filenames are how humans track versions and
 *     also how humans mislabel them.
 *  3. **Cast** — defined-term and party overlap. **A tie-breaker and nothing
 *     more.** Inside one deal bundle the SPA, the disclosure letter and the
 *     side letter share the same parties and most defined terms, so this signal
 *     barely discriminates. It is deliberately absent from the weighted sum
 *     (see `score.ts`); folding it in would let a near-useless signal move a
 *     pair across a confidence-band boundary.
 *
 * Jurisdiction (ADR-0002): the version vocabulary is SG/UK drafting vocabulary
 * — *engrossment*, *conformed*, *initialled* — and the dates stripped are
 * **day-first**. `MM/DD/YYYY` is not handled and must not be added: it is both
 * a `no-us-literals` violation and a correctness bug, since `03/04/2026` in
 * this jurisdiction is 3 April.
 *
 * Pure — no Office JS, no DOM, no network.
 */

/**
 * Jaccard similarity of two sets: |A ∩ B| / |A ∪ B|.
 *
 * Computed exactly rather than estimated. At M, N ≤ 50 with a few thousand
 * paragraphs each, every pair's intersection is trivially cheap, and an exact
 * score is one fewer thing to explain in the pairing grid. A MinHash sketch
 * estimates the same quantity and is the bounded-memory fallback if a bundle
 * ever outgrows this.
 *
 * Two empty sets score 0, not 1: two documents with no comparable content are
 * not a confident match, and returning 1 would auto-pair them.
 */
export function jaccard(a: ReadonlySet<string>, b: ReadonlySet<string>): number {
  if (a.size === 0 || b.size === 0) return 0;
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  let shared = 0;
  for (const value of small) if (large.has(value)) shared += 1;
  return shared / (a.size + b.size - shared);
}

/**
 * Version tokens stripped from a filename before comparison.
 *
 * SG/UK drafting vocabulary. Ordered longest-first where one token is a prefix
 * of another (`execution` before `exec`) so the longer match wins.
 */
const VERSION_WORDS = [
  'engrossment',
  'engrossed',
  'conformed',
  'initialled',
  'execution',
  'blackline',
  'redline',
  'signed',
  'final',
  'clean',
  'draft',
  'copy',
  'exec',
] as const;

/**
 * The browser/OS "already have one of these" marker: `SPA (2).docx`.
 *
 * Stripped while its brackets still exist, i.e. before separators are
 * flattened — afterwards there is nothing left to distinguish it from a real
 * number in the name.
 */
const DUPLICATE_MARKER = /\(\s*\d{1,3}\s*\)/g;

/** `v1`, `v 12`. */
const VERSION_NUMBER = /\bv\s?\d{1,3}\b/g;

/**
 * Characters that join words in a filename rather than belonging to one.
 *
 * Brackets are in here because they *wrap* version tokens (`(clean)`,
 * `(final)`); once the token is gone the brackets are trigram noise that would
 * stop two names for one document scoring 1. `/` is deliberately absent — see
 * the month-first note above.
 */
const SEPARATORS = /[_\-.()[\]{}]+/g;

/** Month names, for the spelled-out day-first forms. */
const MONTHS =
  'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|' +
  'aug(?:ust)?|sep(?:t)?(?:ember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?';

/**
 * Dates, **day-first only** (CONVENTIONS.md → Dates).
 *
 * Covers `2026-08-30`, `30-08-2026`, `30.08.26`, `30 Aug 2026`, `30 August 2026`.
 * There is deliberately no month-first pattern.
 */
const DATE_PATTERNS: RegExp[] = [
  /\b\d{4}[-._]\d{1,2}[-._]\d{1,2}\b/g,
  /\b\d{1,2}[-._]\d{1,2}[-._]\d{2,4}\b/g,
  new RegExp(`\\b\\d{1,2}\\s*(?:${MONTHS})\\s*\\d{2,4}\\b`, 'gi'),
];

/**
 * Reduce a filename to the part that identifies the *document* rather than the
 * *version* of it.
 *
 * `Project_Falcon_SPA_v7_exec.docx` and `SPA (clean) 2026-08-30.docx` should
 * both come down to something built from "project falcon spa" and "spa" — which
 * is the point: the version noise is exactly what makes two names for one
 * document look unrelated.
 */
export function normaliseFileName(name: string): string {
  let out = name.toLowerCase().replace(/\.[a-z0-9]{1,5}$/, '');

  // Dates first: they are the only tokens that *use* `-` and `.` internally,
  // so they have to be recognised before separators are flattened.
  for (const pattern of DATE_PATTERNS) out = out.replace(pattern, ' ');

  // Then the duplicate marker, while its brackets still exist.
  out = out.replace(DUPLICATE_MARKER, ' ');

  // Now separators become spaces. This must precede the remaining strips:
  // `_` is a word character, so `spa_v7` has no word boundary before the `v`
  // and `\bv\d{1,3}\b` could never match it.
  out = out.replace(SEPARATORS, ' ');

  out = out.replace(VERSION_NUMBER, ' ');
  for (const word of VERSION_WORDS) {
    out = out.replace(new RegExp(`\\b${word}\\b`, 'g'), ' ');
  }

  return out.replace(/\s+/g, ' ').trim();
}

/** Character trigrams of a string, padded so short names still produce some. */
function trigrams(value: string): Map<string, number> {
  const padded = `  ${value} `;
  const counts = new Map<string, number>();
  for (let i = 0; i + 3 <= padded.length; i += 1) {
    const gram = padded.slice(i, i + 3);
    counts.set(gram, (counts.get(gram) ?? 0) + 1);
  }
  return counts;
}

/**
 * Sørensen–Dice similarity over character trigrams.
 *
 * Trigrams rather than tokens, because filenames differ by punctuation and
 * abbreviation far more than prose does: `spa` and `s p a` are the same
 * document. (ADR-0026 uses Dice over *content tokens* for paragraph text,
 * where whole words are the right unit; the coefficient is the same, the unit
 * is not.)
 */
export function filenameSimilarity(a: string, b: string): number {
  const left = normaliseFileName(a);
  const right = normaliseFileName(b);
  if (left === '' || right === '') return 0;
  if (left === right) return 1;

  const ga = trigrams(left);
  const gb = trigrams(right);
  let shared = 0;
  let total = 0;
  for (const n of ga.values()) total += n;
  for (const [gram, n] of gb) {
    total += n;
    shared += Math.min(n, ga.get(gram) ?? 0);
  }
  return total === 0 ? 0 : (2 * shared) / total;
}

/**
 * Overlap of two documents' defined terms and party names.
 *
 * **Tie-breaker only.** Read the module comment before using this anywhere the
 * result can decide a pairing on its own.
 */
export function castOverlap(
  a: { definedTerms: ReadonlySet<string>; partyNames: ReadonlySet<string> },
  b: { definedTerms: ReadonlySet<string>; partyNames: ReadonlySet<string> },
): number {
  const union = (x: ReadonlySet<string>, y: ReadonlySet<string>): Set<string> =>
    new Set([...x, ...y]);
  return jaccard(union(a.definedTerms, a.partyNames), union(b.definedTerms, b.partyNames));
}
