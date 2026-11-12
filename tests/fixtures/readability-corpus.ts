/**
 * Redline readability corpus — 30 synthetic amendments, one editing act each.
 * Spec: spec/features/08-compare.md §5.3–5.5. Principles: docs/compare-readability-principles.md.
 *
 * Each synthetic case isolates ONE kind of editing act so the rendering for that
 * act can be judged on its own. `grade` is the human reviewer's verdict on the
 * rendering as at the baseline tag `redline-readability-baseline`; `ideal`
 * records the exact segment sequence that reviewer wanted, where the current
 * rendering falls short.
 *
 * The synthetic text is Singapore / England-&-Wales drafting written for this
 * suite — no client material.
 */

/**
 * How a human reviewer graded the rendering of this case. `unreviewed` marks a
 * case added from a defect found in the code rather than from review —
 * it carries a rendering worth looking at, but no verdict yet.
 */
export type ReadabilityGrade = 'good' | 'partial' | 'bad' | 'unreviewed';

/** One segment of a rendered redline: a maximal run sharing a revision state. */
export interface IdealSeg {
  t: 'eq' | 'ins' | 'del' | 'moveTo' | 'moveFrom';
  s: string;
}

/** What every case carries, whatever it is built from. */
interface ReadabilityCaseBase {
  id: string;
  title: string;
  /** The editing act being isolated. */
  act: string;
  /** Prose description of what an intuitive redline should show. */
  expect: string;
  /** Reviewer's verdict at the baseline. */
  grade: ReadabilityGrade;
  /** Principles the verdict cites (P1–P10). */
  cites: string[];
}

/**
 * A synthetic case: two strings, compared through the Text-vs-text path.
 *
 * `kind` is OPTIONAL here on purpose. The plan's acceptance for widening this
 * type is that all 30 synthetic cases score byte-identically to before, and the
 * surest way to hold that is to not touch any of them. A docx case cannot be
 * mistaken for one of these: it has no `original`/`revised`, so omitting its
 * `kind` fails to type-check rather than silently reading as text.
 */
export interface TextReadabilityCase extends ReadabilityCaseBase {
  kind?: 'text';
  original: string;
  revised: string;
  /** The reviewer's wanted rendering, where the baseline falls short. */
  ideal?: IdealSeg[];
}

/**
 * A whole-document case: two `.docx` paths, compared through the Word-vs-Word
 * path (`extractParts` → `preprocessSide` → the same shared pipeline core).
 *
 * It carries **no `ideal`**, and that is structural rather than an omission:
 * there is no hand-authored wanted rendering for a 26-paragraph contract, and
 * the reference product's output is a reference for what a reader expects, not
 * an assertion target (ADR-0038). `idealMatch` is therefore always `null` here.
 */
export interface DocxReadabilityCase extends ReadabilityCaseBase {
  kind: 'docx';
  originalPath: string;
  revisedPath: string;
}

export type ReadabilityCase = TextReadabilityCase | DocxReadabilityCase;

const eq = (s: string): IdealSeg => ({ t: 'eq', s });
const ins = (s: string): IdealSeg => ({ t: 'ins', s });
const del = (s: string): IdealSeg => ({ t: 'del', s });

export const READABILITY_CASES: ReadabilityCase[] = [
  // ── Tier 1: sub-word / single-token edits ────────────────────────────────
  {
    id: 'C01',
    title: 'Apostrophe moved (singular → plural possessive)',
    act: 'one character, meaning-critical',
    original:
      "9.1 The Vendor shall indemnify the Purchaser against all Losses suffered by the Purchaser's subsidiaries in connection with the Business.",
    revised:
      "9.1 The Vendor shall indemnify the Purchaser against all Losses suffered by the Purchasers' subsidiaries in connection with the Business.",
    expect:
      "One word marked: Purchaser's struck, Purchasers' inserted. Rest of the sentence clean.",
    grade: 'good',
    cites: ['P1', 'P5'],
  },
  {
    id: 'C02',
    title: 'Singular → plural (trailing "es")',
    act: 'one character inside a word',
    original:
      '9.2 The Vendor shall be liable for any Loss arising out of a breach of the Warranties.',
    revised:
      '9.2 The Vendor shall be liable for any Losses arising out of a breach of the Warranties.',
    expect: 'Whole word Loss/Losses swapped, not "Loss" + inserted "es".',
    grade: 'good',
    cites: ['P1'],
  },
  {
    id: 'C03',
    title: 'Punctuation only — semicolon to full stop plus new sentence start',
    act: 'punctuation',
    original:
      '9.3 The indemnity in clause 9.1 is given on a continuing basis; it survives Completion.',
    revised:
      '9.3 The indemnity in clause 9.1 is given on a continuing basis. It survives Completion.',
    expect:
      'Two tiny marks (";"→"." and "it"→"It"). Must not restate the whole sentence.',
    grade: 'partial',
    cites: ['P1', 'P3'],
    ideal: [
      eq('The indemnity in clause 9.1 is given on a continuing basis'),
      ins('.'),
      del(';'),
      eq(' '),
      ins('It'),
      del('it'),
      eq(' survives Completion.'),
    ],
  },
  {
    id: 'C04',
    title: 'Capitalisation only — becoming a defined term',
    act: 'case change (semantically large, visually tiny)',
    original:
      '9.4 The Vendor shall indemnify the Purchaser against any loss of profit suffered after Completion.',
    revised:
      '9.4 The Vendor shall indemnify the Purchaser against any Loss of profit suffered after Completion.',
    expect: 'The single word loss→Loss marked; nothing else.',
    grade: 'good',
    cites: ['P1'],
  },

  // ── Tier 2: discrete word edits ──────────────────────────────────────────
  {
    id: 'C05',
    title: 'Modal swap shall → may',
    act: 'one discrete word, obligation-bearing',
    original:
      '9.5 The Indemnifying Party shall reimburse the Indemnified Party for all reasonable legal costs incurred in defending a Third Party Claim.',
    revised:
      '9.5 The Indemnifying Party may reimburse the Indemnified Party for all reasonable legal costs incurred in defending a Third Party Claim.',
    expect: 'shall struck, may inserted, in place. Classified as obligation language.',
    grade: 'good',
    cites: ['P1', 'P10'],
  },
  {
    id: 'C06',
    title: 'Standard-of-fault word deleted',
    act: 'single-word deletion inside a list',
    original:
      '9.6 The Vendor shall not be liable under clause 9.1 except in the case of fraud, gross negligence or wilful misconduct of the Vendor.',
    revised:
      '9.6 The Vendor shall not be liable under clause 9.1 except in the case of fraud, negligence or wilful misconduct of the Vendor.',
    expect: 'Only "gross " struck. The comma and "negligence" survive untouched.',
    grade: 'good',
    cites: ['P1', 'P4'],
  },
  {
    id: 'C07',
    title: 'Several scattered discrete word swaps in one sentence',
    act: 'multiple independent one-word edits, same paragraph',
    original:
      '9.7 The Vendor shall indemnify the Purchaser against all Losses arising out of any breach of the Warranties.',
    revised:
      '9.7 The Vendor must indemnify the Purchaser against any Losses arising from each breach of the Warranties.',
    expect:
      'Four discrete word swaps (shall→must, all→any, out of→from, any→each), each in place. Not a whole-sentence rewrite.',
    grade: 'good',
    cites: ['P1'],
  },
  {
    id: 'C08',
    title: 'Word transposition — indemnify/defend/hold harmless reordered',
    act: 'reordering words within a sentence (classic diff-salad trigger)',
    original:
      '9.8 The Vendor shall indemnify, defend and hold harmless the Purchaser from and against all Losses.',
    revised:
      '9.8 The Vendor shall defend, indemnify and hold harmless the Purchaser from and against all Losses.',
    expect:
      'One move-or-swap of the smallest span. NOT a token-by-token interleave of "indemnify"/"defend".',
    grade: 'partial',
    cites: ['P7'],
    ideal: [
      eq('The Vendor shall '),
      ins('defend, '),
      eq('indemnify'),
      del(', defend'),
      eq(' and hold harmless the Purchaser from and against all Losses.'),
    ],
  },

  // ── Tier 3: token-boundary edits (merge / split / hyphenation) ───────────
  {
    id: 'C09',
    title: 'Hyphenation — two tokens merged into one',
    act: 'word merge across a space',
    original:
      '9.9 The Purchaser shall notify the Vendor of any Third Party Claim and of any time bar affecting that claim.',
    revised:
      '9.9 The Purchaser shall notify the Vendor of any Third-Party Claim and of any time-bar affecting that claim.',
    expect:
      'Two tight two-word spans re-marked. Must not spill into "Claim"/"affecting".',
    grade: 'good',
    cites: ['P1', 'P4'],
  },
  {
    id: 'C10',
    title: 'Word split — one token becomes two',
    act: 'word split',
    original:
      '9.10 No set-off, counterclaim or deduction may be made against a payment due under this clause 9.',
    revised:
      '9.10 No set off, counterclaim or deduction may be made against a payment due under this clause 9.',
    expect: 'Just "set-off"→"set off". One span.',
    grade: 'good',
    cites: ['P1'],
  },
  {
    id: 'C11',
    title: 'Redundant doublet collapsed',
    act: 'phrase shortened in place',
    original:
      '9.11 The Vendor shall indemnify and hold harmless the Purchaser against any and all Losses whatsoever.',
    revised: '9.11 The Vendor shall indemnify the Purchaser against all Losses.',
    expect:
      'Three deletions ("and hold harmless", "any and", "whatsoever") with the spine of the sentence left clean.',
    grade: 'bad',
    cites: ['P2'],
    ideal: [
      eq('The Vendor shall indemnify '),
      del('and hold harmless '),
      eq('the Purchaser against '),
      del('any and '),
      eq('all Losses'),
      del(' whatsoever'),
      eq('.'),
    ],
  },

  // ── Tier 4: phrase and clause-limb edits ────────────────────────────────
  {
    id: 'C12',
    title: 'Phrase replaced with a longer formulation',
    act: 'phrase replacement',
    original:
      '9.12 The Vendor shall indemnify the Purchaser against all Losses howsoever arising in connection with the Properties.',
    revised:
      '9.12 The Vendor shall indemnify the Purchaser against all Losses arising directly or indirectly out of or in connection with the Properties.',
    expect:
      '"howsoever arising in" struck / new formulation inserted as ONE contiguous block, with "connection with the Properties" left clean.',
    grade: 'partial',
    cites: ['P1', 'P3'],
    ideal: [
      eq('The Vendor shall indemnify the Purchaser against all Losses '),
      del('howsoever '),
      eq('arising '),
      ins('directly or indirectly out of or '),
      eq('in connection with the Properties.'),
    ],
  },
  {
    id: 'C13',
    title: 'Amount and period changed',
    act: 'numeric values',
    original:
      '9.13 The aggregate liability of the Vendor under clause 9.1 shall not exceed S$5,000,000 and no claim may be brought after the date falling 24 months after Completion.',
    revised:
      '9.13 The aggregate liability of the Vendor under clause 9.1 shall not exceed S$7,500,000 and no claim may be brought after the date falling 18 months after Completion.',
    expect:
      'Two atomic numeric swaps. The amount must NOT be split into digit fragments (S$[5→7],[0→5]00,000).',
    grade: 'good',
    cites: ['P5'],
  },
  {
    id: 'C14',
    title: 'Proviso appended to an existing clause',
    act: 'tail insertion',
    original:
      '9.14 The Vendor shall indemnify the Purchaser against all Tax Liabilities of the Company arising in respect of any period ending on or before Completion.',
    revised:
      '9.14 The Vendor shall indemnify the Purchaser against all Tax Liabilities of the Company arising in respect of any period ending on or before Completion, provided that the Vendor shall have no liability under this clause 9.14 to the extent that the Tax Liability has been paid before Completion.',
    expect:
      'The original sentence entirely clean; one contiguous inserted tail from the comma onwards.',
    grade: 'good',
    cites: ['P3', 'P4'],
  },
  {
    id: 'C15',
    title: 'Carve-out deleted mid-sentence',
    act: 'parenthetical deletion',
    original:
      '9.15 The limitations in clause 9.13 (other than in respect of fraud or fraudulent misrepresentation) apply to every claim under this Agreement.',
    revised:
      '9.15 The limitations in clause 9.13 apply to every claim under this Agreement.',
    expect: 'One contiguous struck parenthetical, brackets included. Nothing else marked.',
    grade: 'good',
    cites: ['P4'],
  },
  {
    id: 'C16',
    title: 'Cross-reference renumbered',
    act: 'cross-reference only',
    original:
      '9.16 Clause 9.13 does not limit the indemnity in clause 8.2 or the Purchaser’s rights under clause 8.4.',
    revised:
      '9.16 Clause 9.13 does not limit the indemnity in clause 8.3 or the Purchaser’s rights under clause 8.5.',
    expect:
      'Two tight cross-ref swaps, classified as cross-reference/renumbering so they can be filtered out of the substantive view.',
    grade: 'partial',
    cites: ['P6', 'P10'],
    ideal: [
      eq('Clause 9.13 does not limit the indemnity in clause '),
      ins('8.3'),
      del('8.2'),
      eq(' or the Purchaser’s rights under clause '),
      ins('8.5'),
      del('8.4'),
      eq('.'),
    ],
  },

  // ── Tier 5: paragraph-level structural edits ────────────────────────────
  {
    id: 'C17',
    title: 'Wholesale rewrite of one limb',
    act: 'near-total paragraph rewrite',
    original:
      '9.17 The Purchaser shall take all reasonable steps to mitigate any Loss which gives rise to a claim under the Warranties.',
    revised:
      '9.17 Nothing in this Agreement restricts or limits the general obligation at law of the Purchaser to mitigate any Loss or damage which it may suffer.',
    expect:
      'A CLEAN swap: whole old paragraph struck, whole new paragraph inserted. Not a token interleave (this is the ADR-0026 demotion case).',
    grade: 'good',
    cites: ['P2'],
  },
  {
    id: 'C18',
    title: 'Paragraph moved verbatim (unchanged text, new position)',
    act: 'verbatim move',
    original: [
      '9.18 The Vendor shall indemnify the Purchaser against all Environmental Liabilities.',
      '9.19 The Purchaser shall notify the Vendor in writing of any Third Party Claim within 20 Business Days.',
      '9.20 The Vendor may assume the conduct of any Third Party Claim at its own cost.',
      '9.21 This clause 9 survives termination of this Agreement.',
    ].join('\n'),
    revised: [
      '9.18 The Vendor shall indemnify the Purchaser against all Environmental Liabilities.',
      '9.19 This clause 9 survives termination of this Agreement.',
      '9.20 The Purchaser shall notify the Vendor in writing of any Third Party Claim within 20 Business Days.',
      '9.21 The Vendor may assume the conduct of any Third Party Claim at its own cost.',
    ].join('\n'),
    expect:
      'The survival clause shown as MOVED (move-from at the old slot, move-to at the new), NOT as delete+insert, and NOT as three rewritten paragraphs.',
    grade: 'good',
    cites: ['P7'],
  },
  {
    id: 'C19',
    title: 'Paragraph moved verbatim AND renumbered',
    act: 'move + number change only',
    original: [
      '9.22 The Vendor shall indemnify the Purchaser against all Losses relating to the Excluded Assets.',
      '9.23 Each indemnity in this clause 9 is given severally.',
      '9.24 Payments under this clause 9 shall be made free of set-off.',
    ].join('\n'),
    revised: [
      '9.22 Payments under this clause 9 shall be made free of set-off.',
      '9.23 The Vendor shall indemnify the Purchaser against all Losses relating to the Excluded Assets.',
      '9.24 Each indemnity in this clause 9 is given severally.',
    ].join('\n'),
    expect:
      'One moved paragraph. The renumbering of the other two must be collapsed/filterable, not shown as three substantive edits.',
    grade: 'good',
    cites: ['P7', 'P8'],
  },
  {
    id: 'C20',
    title: 'New limb inserted into a lettered list (renumbering cascade)',
    act: 'insertion causing downstream renumbering',
    original: [
      '9.25 The Vendor shall indemnify the Purchaser against:',
      '(a) all Tax Liabilities of the Company;',
      '(b) all Environmental Liabilities of the Company; and',
      '(c) all Losses arising from the Excluded Contracts.',
    ].join('\n'),
    revised: [
      '9.25 The Vendor shall indemnify the Purchaser against:',
      '(a) all Tax Liabilities of the Company;',
      '(b) all Employee Liabilities of the Company;',
      '(c) all Environmental Liabilities of the Company; and',
      '(d) all Losses arising from the Excluded Contracts.',
    ].join('\n'),
    expect:
      'ONE inserted limb. (b)/(c) relabelling must not read as two rewritten limbs; only "; and" moving from (b) to (c) is a real edit.',
    grade: 'good',
    cites: ['P8'],
  },
  {
    id: 'C21',
    title: 'Two paragraphs merged into one',
    act: 'paragraph merge',
    original: [
      '9.26 The Purchaser shall notify the Vendor of any Third Party Claim.',
      'The notice shall specify the nature of the claim in reasonable detail.',
    ].join('\n'),
    revised:
      '9.26 The Purchaser shall notify the Vendor of any Third Party Claim, and the notice shall specify the nature of the claim in reasonable detail.',
    expect:
      'A small join: full stop → comma + "and", second sentence otherwise clean. Not two deletes plus one long insert.',
    grade: 'partial',
    cites: ['P2'],
    // ── This ideal is UNREACHABLE, and the reason is not a shortcoming ──────
    // Verified 2026-09-04 . The ideal's closing `eq(' the notice
    // …')` requires `The` (original, sentence-initial) and `the` (revised, mid
    // sentence) to compare EQUAL. They do not, and they must not: the del side
    // of a redline has to reconstruct the original exactly (ADR-0008), and with
    // that equality it reconstructs `…Claim. the notice…` against an original
    // of `…Claim.¶The notice…`. Any rule that closed the gap — case-insensitive
    // equality at a merge seam — would break reconstruction.
    //
    // What the tool actually emits is optimal given that constraint: TWO marks
    // (`[-.-]` and `[-The-]{+, and the+}`), zero redundancy, and — checked in
    // the authored OOXML, not inferred — two paragraphs with one paragraph-mark
    // deletion, so the merge is properly rejectable (ADR-0032). No literal
    // newline reaches the struck text.
    //
    // So `idealMatch: false` here records a disagreement about the IDEAL, not a
    // defect in the rendering. Left as the reviewer wrote it rather than
    // rewritten to match the code — changing it is the reviewer's call. Do not
    // spend another session closing this gap; it cannot be closed.
    ideal: [
      eq('The Purchaser shall notify the Vendor of any Third Party Claim'),
      del('.'),
      ins(', and'),
      eq(' the notice shall specify the nature of the claim in reasonable detail.'),
    ],
  },
  {
    id: 'C22',
    title: 'Defined term renamed throughout',
    act: 'repeated identical substitution across paragraphs',
    original: [
      '9.27 The Indemnifying Party shall indemnify the Indemnified Party against all Losses.',
      '9.28 The Indemnified Party shall notify the Indemnifying Party of any claim.',
      '9.29 The Indemnifying Party may require the Indemnified Party to assign the claim.',
    ].join('\n'),
    revised: [
      '9.27 The Covenantor shall indemnify the Beneficiary against all Losses.',
      '9.28 The Beneficiary shall notify the Covenantor of any claim.',
      '9.29 The Covenantor may require the Beneficiary to assign the claim.',
    ].join('\n'),
    expect:
      'Six tight defined-term swaps, in place, classified as defined-term changes. Paragraph spines untouched.',
    grade: 'good',
    cites: ['P1', 'P10'],
  },

  // ── Tier 6: cross-reference granularity (found in code, awaiting review) ──
  {
    id: 'C23',
    title: 'Capitalised cross-reference renumbered',
    act: 'cross-reference only, reference token capitalised',
    original:
      '9.30 The limitations in Clause 9.13 apply to every claim under this Agreement.',
    revised:
      '9.30 The limitations in Clause 9.14 apply to every claim under this Agreement.',
    expect:
      'Only the reference value marked: the word "Clause" is identical on both sides and must stay outside the mark.',
    grade: 'unreviewed',
    cites: ['P1', 'P5'],
    ideal: [
      eq('The limitations in Clause '),
      ins('9.14'),
      del('9.13'),
      eq(' apply to every claim under this Agreement.'),
    ],
  },
  {
    id: 'C24',
    title: 'Compound cross-reference, one component changed',
    act: 'cross-reference only, "paragraph N of Schedule M" form',
    original:
      '9.31 The Vendor’s liability under paragraph 2 of Schedule 4 is not limited by this clause 9.',
    revised:
      '9.31 The Vendor’s liability under paragraph 3 of Schedule 4 is not limited by this clause 9.',
    expect:
      'One digit marked. "of Schedule 4" is unchanged and must not be dragged into the mark.',
    grade: 'unreviewed',
    cites: ['P1', 'P5'],
    ideal: [
      eq('The Vendor’s liability under paragraph '),
      ins('3'),
      del('2'),
      eq(' of Schedule 4 is not limited by this clause 9.'),
    ],
  },

  // ── Tier 7: further editing acts (found by coverage review, awaiting verdict) ──
  {
    id: 'C25',
    title: 'Two sentences transposed within one paragraph',
    act: 'verbatim reorder at sentence scale',
    original:
      '9.32 The Purchaser shall mitigate any Loss. The Vendor shall have no liability for any Loss which the Purchaser could reasonably have avoided.',
    revised:
      '9.32 The Vendor shall have no liability for any Loss which the Purchaser could reasonably have avoided. The Purchaser shall mitigate any Loss.',
    expect:
      'A reorder is one act. The shorter sentence should move; the longer one must not be struck and retyped around it. Either order of the two marks reads acceptably.',
    grade: 'unreviewed',
    cites: ['P7', 'P2'],
  },
  {
    id: 'C26',
    title: 'Quoted defined term renamed',
    act: 'defined term inside quotation marks',
    original:
      '9.33 The Vendor shall indemnify the Purchaser against all Losses arising from any “Permitted Encumbrance” affecting the Properties.',
    revised:
      '9.33 The Vendor shall indemnify the Purchaser against all Losses arising from any “Permitted Security” affecting the Properties.',
    expect:
      'One mark on the quoted term. The term must not be split mid-phrase, and the surrounding sentence stays clean.',
    grade: 'unreviewed',
    cites: ['P5', 'P1'],
  },
  {
    id: 'C27',
    title: 'Cap changed in words and figures',
    act: 'paired numeric and verbal amount (standard SG drafting)',
    original:
      '9.34 The aggregate liability of the Vendor shall not exceed S$5,000,000 (five million Singapore dollars).',
    revised:
      '9.34 The aggregate liability of the Vendor shall not exceed S$7,500,000 (seven million five hundred thousand Singapore dollars).',
    expect:
      'The figure changes as one atomic mark and the words as a second; neither may fragment, and the brackets stay outside the words mark.',
    grade: 'unreviewed',
    cites: ['P5', 'P4'],
  },
  {
    id: 'C28',
    title: 'Limb deleted from a lettered list (reverse cascade)',
    act: 'deletion causing downstream renumbering',
    original: [
      '9.35 The Vendor shall indemnify the Purchaser against:',
      '(a) all Tax Liabilities of the Company;',
      '(b) all Employee Liabilities of the Company;',
      '(c) all Environmental Liabilities of the Company; and',
      '(d) all Losses arising from the Excluded Contracts.',
    ].join('\n'),
    revised: [
      '9.35 The Vendor shall indemnify the Purchaser against:',
      '(a) all Tax Liabilities of the Company;',
      '(b) all Environmental Liabilities of the Company; and',
      '(c) all Losses arising from the Excluded Contracts.',
    ].join('\n'),
    expect:
      'ONE struck limb. The (c)→(b), (d)→(c) relabelling is a consequence and must not read as two rewritten limbs.',
    grade: 'unreviewed',
    cites: ['P8'],
  },
  {
    id: 'C29',
    title: 'Cross-reference shift and a substantive change in one sentence',
    act: 'mixed act — tests that triage does not bury the real edit',
    original:
      '9.36 Nothing in clause 9.13 limits the Vendor’s liability for fraud under clause 8.2.',
    revised:
      '9.36 Nothing in clause 9.14 limits the Vendor’s liability for fraud or wilful misconduct under clause 8.3.',
    expect:
      'Two cross-reference marks that a reviewer can fold away, and one substantive insertion that survives the fold.',
    grade: 'unreviewed',
    cites: ['P10', 'P1'],
    ideal: [
      eq('Nothing in clause '),
      ins('9.14'),
      del('9.13'),
      eq(' limits the Vendor’s liability for fraud'),
      ins(' or wilful misconduct'),
      eq(' under clause '),
      ins('8.3'),
      del('8.2'),
      eq('.'),
    ],
  },
  {
    id: 'C30',
    title: 'Whitespace and quote-style normalisation only',
    act: 'no semantic change at all',
    original:
      '9.37 The Vendor shall indemnify the  Purchaser against all "Losses" suffered after Completion.',
    revised:
      '9.37 The Vendor shall indemnify the Purchaser against all “Losses” suffered after Completion.',
    expect:
      'Currently INVISIBLE: §5.2 normalisation folds whitespace runs and curly quotes before the paragraphs are matched, so the pair hashes equal and no op is created at all. That is the designed behaviour and matches what commercial comparison tools do — nobody wants a redline that marks every double space — but it is a judgement call a reviewer should make consciously, and it means the §7 rule forcing signature-block changes to be unfilterable never fires for a spacing-only edit inside one.',
    grade: 'unreviewed',
    cites: ['P10'],
  },
  // ── Tier 6: the two word-diff defects (added 2026-09-11) ─────
  // Found by driving the engine over longer documents, not by review.
  // All three score `redundancy` ZERO at the time they were written and score
  // `charRedundancy` 1.00 — they exist because the token yardstick is derived
  // from the production tokenizer and cannot see a defect inside a token
  // (ADR-0038). Added BEFORE either fix, so the fix has something to prove.
  {
    id: 'C31',
    title: 'Two possessives in one clause, a word changed between them',
    act: 'one word, inside an apostrophe-spanned span',
    original:
      "9.31 The Company's board shall not, without the Purchaser's prior written consent, declare a dividend.",
    revised:
      "9.31 The Company's board may not, without the Purchaser's prior written consent, declare a dividend.",
    expect:
      "Two words marked: `shall` struck, `may` inserted. Everything else is identical on both sides. Today the `definedTerm` matcher in `diff/atomic-tokens.ts` includes `'[^']+'`, which spans everything between any two apostrophes — and in English legal drafting that means everything between any two possessives. The span `'s board shall not, without the Purchaser'` is ATOMIC, so P5 forbids splitting it, so a one-word change strikes and re-inserts the whole thing.",
    grade: 'unreviewed',
    cites: ['P1', 'P2', 'P5'],
  },
  {
    id: 'C32',
    title: 'A word gaining a plural suffix',
    act: 'one character at the end of a word',
    original:
      '9.32 The Financial Statement shall be prepared in accordance with GAAP.',
    revised:
      '9.32 The Financial Statements shall be prepared in accordance with GAAP.',
    expect:
      "The reader should see the word gain its `s`. Today it reads `{+Statements+}[-Statement-]` — the whole word restated on both sides to say one letter was added. `refine.ts`'s acceptance rule compares TOKEN identity (`keyOf`), so `Statement` and `Statements` are simply two unrelated tokens: the re-diff recovers no shared token, `overlapSplit` has nothing to pull out, and the block is left coarse. A word gaining a suffix is among the commonest edits in drafting.",
    grade: 'unreviewed',
    cites: ['P1', 'P2'],
  },
  {
    id: 'C33',
    title: 'One word changed inside a long apostrophe-spanned clause',
    act: 'one word, inside a 158-character bogus atomic span',
    // Written for this suite, at the LENGTH the defect reaches in a long
    // contract: two possessives far apart with one word between them.
    original:
      "9.33 Within the Notice Period, the Landlord and the Landlord's agent shall be given full notice of the claims and reports made to the Insurer, the brokers of, and any surveys commissioned by, the Tenant and the Tenant's insurer.",
    revised:
      "9.33 Within the Notice Period, the Landlord and the Landlord's agent shall be given reasonable notice of the claims and reports made to the Insurer, the brokers of, and any surveys commissioned by, the Tenant and the Tenant's insurer.",
    expect:
      'One word marked: `full` struck, `reasonable` inserted. This is C31 at the scale it reaches in a long contract — the bogus atomic span runs 158 characters from the first possessive to the second, so changing one word strikes that whole span and re-inserts it six characters longer, nearly all of it verbatim. The worst instance restates hundreds of identical characters to say that a clause was appended.',
    grade: 'unreviewed',
    cites: ['P1', 'P2', 'P5'],
  },
  {
    id: 'C34',
    title: 'A definition deleted directly above an edited definition',
    act: 'paragraph deletion beside a two-word edit, in one unanchored gap',
    // A reported defect, at its SHAPE and in new words (ADR-0039).
    // Positional gap pairing paired the deleted definition with the edited one,
    // ADR-0026 rightly demoted that mismatched pair, and the edited definition
    // was struck whole and re-typed beside a deletion.
    original: [
      'In this Agreement the following words have the meanings set out below.',
      'Approved Supplier means a supplier approved in writing by the Buyer under clause 4.',
      'Working Day means any day on which clearing banks in Singapore are open for general business.',
      'Clause headings do not affect the interpretation of this Agreement.',
    ].join('\n'),
    revised: [
      'In this Agreement the following words have the meanings set out below.',
      'Working Day means any day on which clearing banks in Singapore and London are open for general business.',
      'Clause headings do not affect the interpretation of this Agreement.',
    ].join('\n'),
    expect:
      'The Approved Supplier definition struck whole; the Working Day definition unchanged except `and London ` inserted. Not the Working Day definition struck and re-typed.',
    grade: 'unreviewed',
    cites: ['P1', 'P2'],
  },
];
