/**
 * ADR-0039 — inside a gap, Level 1 pairs paragraphs by similarity, not by position.
 * Spec: spec/features/08-compare.md §5.1.
 *
 * Driven through `runComparePipeline`, deliberately, and not through
 * `structuralDiff`: the pairing is an OPTION the pipeline passes, and a test
 * that builds the option itself would stay green with the pipeline's own seam
 * deleted (the 08b lesson — `runPipeline`'s numbering seam went unguarded for
 * several reviews for exactly that reason).
 *
 * Two kinds of test live here, and the difference matters:
 *
 *  - **Behaviour** — the reported shape and its siblings. These fail on the
 *    positional pairing that preceded ADR-0039.
 *  - **Guards** — the rendering that must NOT change (rule 1's lock, a merge, a
 *    table unit). These passed before ADR-0039 and must keep passing; each one
 *    fails if the guard it names is removed (recorded in the plan's Evidence).
 *
 * Prose is written for this suite. It carries the SHAPE of a reported defect
 * (a definition deleted directly above an edited definition), never its words.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runComparePipeline } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';
import type { DiffOp, WordDiff } from '@features/compare/model.js';

const CONFIG = getCompareConfig(SG_PROFILE);

const OPENING = 'In this Agreement the following words have the meanings set out below.';
const CLOSING = 'Clause headings do not affect the interpretation of this Agreement.';
const APPROVED =
  'Approved Supplier means a supplier approved in writing by the Buyer under clause 4.';
const WORKING_DAY =
  'Working Day means any day on which clearing banks in Singapore are open for general business.';
const WORKING_DAY_EDITED =
  'Working Day means any day on which clearing banks in Singapore and London are open for general business.';
const RECORDS =
  'Records means all books, ledgers and correspondence kept by the Seller in relation to the Business.';

/** Op types in order, with each paragraph-scale op's first few words, for a readable failure. */
async function run(original: string[], revised: string[]): Promise<DiffOp[]> {
  const result = await runComparePipeline({
    originalParas: extractTextRecords(original.join('\n')),
    revisedParas: extractTextRecords(revised.join('\n')),
    config: CONFIG,
    author: 'Test',
    date: '2026-10-01T00:00:00Z',
  });
  expectDocumentOrder(result.ops);
  return result.ops;
}

const types = (ops: DiffOp[]): string[] => ops.map((o) => o.type);

/**
 * Every op list in this file must walk BOTH documents in order — a redline is
 * authored by walking the ops (spec §5.5), so an op that visits a paragraph out
 * of order rebuilds its side in the wrong order on accept-all / reject-all.
 * Checked on every run, not just the test that first found it (08f task 9).
 * Each half of a `move` is its own op with one range, at its own side's place.
 */
function expectDocumentOrder(ops: DiffOp[]): void {
  let lastO = -1;
  let lastR = -1;
  for (const op of ops) {
    if (op.origRange !== null) {
      expect(op.origRange.start, `original side out of order at ${op.id}`).toBeGreaterThan(lastO);
      lastO = op.origRange.end;
    }
    if (op.revRange !== null) {
      expect(op.revRange.start, `revised side out of order at ${op.id}`).toBeGreaterThan(lastR);
      lastR = op.revRange.end;
    }
  }
}

/** The rendered word diff of a replace op, in the suite's `{+ins+}` / `[-del-]` notation. */
function spell(diffs: WordDiff[] | undefined): string {
  return (diffs ?? [])
    .map((d) => (d.op === 'eq' ? d.text : d.op === 'ins' ? `{+${d.text}+}` : `[-${d.text}-]`))
    .join('');
}

describe('ADR-0039 behaviour — a gap is paired by similarity', () => {
  it('a definition deleted directly ABOVE an edited one: the edit renders inline (the reported shape)', async () => {
    const ops = await run(
      [OPENING, APPROVED, WORKING_DAY, CLOSING],
      [OPENING, WORKING_DAY_EDITED, CLOSING],
    );
    expect(types(ops)).toEqual(['equal', 'delete', 'replace', 'equal']);
    expect(spell(ops[2]!.wordDiffs)).toBe(
      'Working Day means any day on which clearing banks in Singapore {+and London +}are open for general business.',
    );
  });

  it('a definition inserted directly ABOVE an edited one: the edit renders inline', async () => {
    const ops = await run(
      [OPENING, WORKING_DAY, CLOSING],
      [OPENING, APPROVED, WORKING_DAY_EDITED, CLOSING],
    );
    expect(types(ops)).toEqual(['equal', 'insert', 'replace', 'equal']);
  });

  it('a deletion above AND an insertion below the edit: the edit still renders inline', async () => {
    const ops = await run(
      [OPENING, APPROVED, WORKING_DAY, CLOSING],
      [OPENING, WORKING_DAY_EDITED, RECORDS, CLOSING],
    );
    expect(types(ops)).toEqual(['equal', 'delete', 'replace', 'insert', 'equal']);
  });

  it('every pair is shifted by one inserted paragraph: each edit still finds its counterpart', async () => {
    // The redline-readability-parity plan's issue 1: one inserted paragraph at
    // the head of an unanchored run moved every later pairing by one. The three
    // clauses share no boilerplate, so no shifted positional pair passes the
    // gate and none is locked.
    const clauses = [
      'The Seller shall deliver the Goods to the Delivery Address on the Delivery Date.',
      'Any invoice not paid when due bears interest at four per cent a year above base rate.',
      'The Supplier may not subcontract any part of the Services without first consulting the Customer.',
    ];
    const edited = [
      'The Seller shall deliver the Goods to the Delivery Address on or before the Delivery Date.',
      'Any invoice not paid when due bears interest at six per cent a year above base rate.',
      'The Supplier may not subcontract or delegate any part of the Services without first consulting the Customer.',
    ];
    const ops = await run(
      [OPENING, ...clauses, CLOSING],
      [OPENING, 'A Schedule of Rates is set out in Annex B.', ...edited, CLOSING],
    );
    expect(types(ops)).toEqual(['equal', 'insert', 'replace', 'replace', 'replace', 'equal']);
  });

  it('a paragraph with no content tokens is never a candidate, so it cannot outbid a real counterpart', async () => {
    // Two contentless separators score Dice 1 against each other (both token
    // bags empty). They cross the Working Day pair, so an unguarded alignment
    // takes the separators (1.0 beats ~0.95) and leaves the definition churned.
    const ops = await run(
      [OPENING, '* * *', WORKING_DAY, CLOSING],
      [OPENING, WORKING_DAY_EDITED, '- - -', CLOSING],
    );
    const replaces = ops.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(spell(replaces[0]!.wordDiffs)).toContain('{+and London +}');
  });
});

describe('ADR-0039 guards — renderings that must not change', () => {
  it('rule 1 (lock): a positional pair that already renders inline keeps its partner, even beside a better one', async () => {
    // Positional pairs WORKING_DAY with `closer` (Dice ≥ 0.5, so it renders
    // inline today). `closest` scores higher, and an unlocked alignment would
    // trade the inline pair for it. The requirement — nothing that
    // reads well today may change — is that this never happens.
    const closer =
      'Working Day means any day on which clearing banks in Singapore are open for business.';
    const closest =
      'Working Day means any day on which clearing banks in Singapore are open for general banking business.';
    const ops = await run([OPENING, WORKING_DAY, CLOSING], [OPENING, closer, closest, CLOSING]);
    expect(types(ops)).toEqual(['equal', 'replace', 'insert', 'equal']);
    expect(ops[1]!.revRange).toEqual({ start: 1, end: 1 });
  });

  it('ACCEPTED LIMIT (ADR-0039): a wrong positional pair that already passes the gate is kept', async () => {
    // Sibling clauses sharing their boilerplate (`… passes to the Buyer on …`)
    // score Dice 0.67 against each other, so the shifted positional pair
    // Title ↔ Risk(edited) renders inline today and rule 1 locks it. The pass
    // still re-pairs the first clause; the other two stay as they were. Relaxing
    // the lock would fix this case and is what broke C21 on the prototype —
    // a follow-up decision, not a defect of this pass. If the lock
    // is ever relaxed, this becomes ['equal', 'insert', 'replace', 'replace',
    // 'replace', 'equal'].
    const ops = await run(
      [
        OPENING,
        'The Seller shall deliver the Goods to the Delivery Address on the Delivery Date.',
        'Risk in the Goods passes to the Buyer on completion of delivery.',
        'Title to the Goods passes to the Buyer on payment of the Price in full.',
        CLOSING,
      ],
      [
        OPENING,
        'The Goods are described in Schedule 1.',
        'The Seller shall deliver the Goods to the Delivery Address on or before the Delivery Date.',
        'Risk in the Goods passes to the Buyer on completion of unloading.',
        'Title to the Goods passes to the Buyer on receipt of the Price in full.',
        CLOSING,
      ],
    );
    expect(types(ops)).toEqual(['equal', 'insert', 'replace', 'delete', 'replace', 'insert', 'equal']);
    // The locked pair: original Title (3) against revised Risk (3).
    expect(ops[4]!.origRange).toEqual({ start: 3, end: 3 });
    expect(ops[4]!.revRange).toEqual({ start: 3, end: 3 });
  });

  it('a merge stays ONE range-spanning replace (ADR-0032; the C21 shape the unlocked alignment broke)', async () => {
    const ops = await run(
      [
        OPENING,
        '9.26 The Purchaser shall notify the Vendor of any Third Party Claim.',
        'The notice shall specify the nature of the claim in reasonable detail.',
        CLOSING,
      ],
      [
        OPENING,
        '9.26 The Purchaser shall notify the Vendor of any Third Party Claim, and the notice shall specify the nature of the claim in reasonable detail.',
        CLOSING,
      ],
    );
    expect(types(ops)).toEqual(['equal', 'replace', 'equal']);
    expect(ops[1]!.origRange).toEqual({ start: 1, end: 2 });
  });

  // ── A merge or split whose HEAD was also edited (spec-validator, 08f task 7) ──
  // The head's positional pair fails ADR-0026's gate ON ITS OWN, and renders
  // inline only because ADR-0032 joins it with the orphan after it and scores
  // the joined text. A lock that asked the gate about the pair alone let the
  // alignment take the merge apart: the head struck whole and its edited
  // wording re-inserted inside the tail's paragraph — the repeated-text failure
  // ADR-0032 exists to remove. The lock now asks the real passes.
  const HEAD = 'The Purchaser shall notify the Vendor promptly.';
  const TAIL =
    'The notice shall specify the nature of the claim in reasonable detail and attach copies of all relevant correspondence received.';
  const JOINED =
    'The Purchaser shall within five Business Days notify the Vendor in writing, and the notice shall specify the nature of the claim in reasonable detail and attach copies of all relevant correspondence received.';

  it('a merge whose head was edited stays ONE range-spanning replace', async () => {
    const ops = await run([OPENING, HEAD, TAIL, CLOSING], [OPENING, JOINED, CLOSING]);
    expect(types(ops)).toEqual(['equal', 'replace', 'equal']);
    expect(ops[1]!.origRange).toEqual({ start: 1, end: 2 });
  });

  it('a split whose head was edited stays ONE range-spanning replace (the mirror)', async () => {
    const ops = await run([OPENING, JOINED, CLOSING], [OPENING, HEAD, TAIL, CLOSING]);
    expect(types(ops)).toEqual(['equal', 'replace', 'equal']);
    expect(ops[1]!.revRange).toEqual({ start: 1, end: 2 });
  });

  it('a paragraph moved verbatim keeps its move when an edited copy appears elsewhere (Level 3)', async () => {
    // Positional pairing leaves WORKING_DAY as surplus and Level 3 pairs it, by
    // hash, with its verbatim copy at the end. Re-pairing it with the edited
    // copy at the top would turn the move into a plain insertion — a different
    // rendering from one that was not churn (08f task 7). A record whose text
    // appears verbatim on the other side is HELD: never re-paired.
    const price = 'Price means the price stated in the Order.';
    const ops = await run(
      [OPENING, APPROVED, WORKING_DAY, RECORDS, price, CLOSING],
      [OPENING, WORKING_DAY_EDITED, RECORDS, price, WORKING_DAY, CLOSING],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
  });

  it('a held move does not stop its neighbours being re-paired: the move survives AND the edit reads inline', async () => {
    // A table shape: a cell value moves while the cell
    // beside it is edited. Skipping the whole stretch for the move's sake (the
    // first fix) kept the move and lost the edit.
    const mover = 'Records means all books and ledgers kept by the Seller.';
    const middle = 'Price means the price stated in the Order.';
    const ops = await run(
      [OPENING, mover, middle, APPROVED, WORKING_DAY, CLOSING],
      [OPENING, middle, WORKING_DAY_EDITED, mover, CLOSING],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
    const replaces = ops.filter((o) => o.type === 'replace');
    expect(replaces).toHaveLength(1);
    expect(spell(replaces[0]!.wordDiffs)).toContain('{+and London +}');
  });

  it('the FALLBACK never pairs a held record either, even with a partner that would read inline', async () => {
    // `held` moves verbatim to the end. Beside where it was, the revised side
    // has `heldExtended` — the same words with a long passage appended, so the
    // pair clears ADR-0026 only through ADR-0041's contiguity escape and is not
    // a Dice candidate. The alignment re-pairs the definition; positional
    // fallback would then pair `held` with `heldExtended` and take the move away.
    const held = 'Records means all books and ledgers kept by the Seller.';
    const heldExtended =
      'Records means all books and ledgers kept by the Seller, together with every electronic copy, back-up, archive and index of them held by or for the Seller or any member of its Group at any time before Completion.';
    const mid1 = 'Price means the price stated in the Order.';
    const mid2 = 'Order means a purchase order issued by the Buyer.';
    const ops = await run(
      [OPENING, APPROVED, WORKING_DAY, held, mid1, mid2, CLOSING],
      [OPENING, WORKING_DAY_EDITED, heldExtended, mid1, mid2, held, CLOSING],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
    expect(spell(ops.find((o) => o.type === 'replace')?.wordDiffs)).toContain('{+and London +}');
  });

  it('the hold applies to the REVISED side too: a revised record headed for a move is never a candidate', async () => {
    // The mirror of the verbatim-move test. WORKING_DAY is new text in the
    // revised document but appears verbatim lower down the original, so it is
    // held; aligning it with the edited copy at the top would take the move
    // away (08f task 11 — every earlier hold test held an ORIGINAL record).
    const price = 'Price means the price stated in the Order.';
    const ops = await run(
      [OPENING, WORKING_DAY_EDITED, RECORDS, price, WORKING_DAY, CLOSING],
      [OPENING, APPROVED, WORKING_DAY, RECORDS, price, CLOSING],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
  });

  it('the FALLBACK never pairs a held REVISED record either (the mirror)', async () => {
    const held = 'Records means all books and ledgers kept by the Seller.';
    const heldExtended =
      'Records means all books and ledgers kept by the Seller, together with every electronic copy, back-up, archive and index of them held by or for the Seller or any member of its Group at any time before Completion.';
    const mid1 = 'Price means the price stated in the Order.';
    const mid2 = 'Order means a purchase order issued by the Buyer.';
    const ops = await run(
      [OPENING, held, mid1, mid2, APPROVED, WORKING_DAY, heldExtended, CLOSING],
      [OPENING, mid1, mid2, WORKING_DAY_EDITED, held, CLOSING],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
  });

  describe('a held record is never absorbed by a merge it did not join before (08f task 11)', () => {
    // Re-pairing can put a held record directly after a replace it never
    // followed under positional pairing. Stage 2.25 then took it as a merge or
    // split orphan, its verbatim twin lost its partner, and the move became a
    // plain insertion — in a gap with no churn at all.
    const schedules = 'Each Schedule to this Agreement is numbered in the order in which it is first cited.';
    const held = 'Price means the price stated in the Order.';
    const milestone = 'Milestone means a date by which a Deliverable must be accepted.';
    const recordsEdited = RECORDS.replace('kept', 'held');
    const unstop = (t: string): string => t.replace(/\.$/, '');

    it('beside a re-aligned host (merge)', async () => {
      const ops = await run(
        [OPENING, milestone, RECORDS, held, CLOSING, schedules],
        [OPENING, `${unstop(recordsEdited)} ${held}`, CLOSING, schedules, held],
      );
      expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
    });

    it('beside a LOCKED host (merge)', async () => {
      const noticePeriod =
        'Notice Period means the period of thirty days beginning on the day after a notice is served under clause 12.';
      const ops = await run(
        [OPENING, noticePeriod, held, WORKING_DAY, CLOSING, schedules],
        [OPENING, `${noticePeriod} ${held}`, WORKING_DAY_EDITED, CLOSING, schedules, held],
      );
      expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
      expect(ops.filter((o) => o.type === 'replace').map((o) => o.origRange)).toContainEqual({ start: 1, end: 1 });
    });

    it('beside a re-aligned host (split, the mirror)', async () => {
      const ops = await run(
        [OPENING, `${unstop(RECORDS)} ${held}`, CLOSING, schedules, held],
        [OPENING, milestone, recordsEdited, held, CLOSING, schedules],
      );
      expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
    });

    it('a held record positional pairing DID absorb is still absorbed (nothing that merged before stops merging)', async () => {
      // HEAD and TAIL join; TAIL also appears verbatim later in the revised
      // document, so it is held — but the merge took it before ADR-0039 and
      // must still take it.
      const ops = await run([OPENING, HEAD, TAIL, CLOSING], [OPENING, JOINED, CLOSING, TAIL]);
      expect(ops[1]!.type).toBe('replace');
      expect(ops[1]!.origRange).toEqual({ start: 1, end: 2 });
    });

    it('… and the same for a SPLIT whose added half is held (the mirror)', async () => {
      // TAIL appears verbatim earlier in the original, so the split's added
      // half is held on the revised side; positional pairing split it before.
      const price = 'Price means the price stated in the Order.';
      const order = 'Order means a purchase order issued by the Buyer.';
      const ops = await run(
        [OPENING, TAIL, price, order, JOINED, CLOSING],
        [OPENING, price, order, HEAD, TAIL, CLOSING],
      );
      expect(ops.filter((o) => o.type === 'replace').map((o) => o.revRange)).toContainEqual({ start: 3, end: 4 });
    });
  });

  it('a held record BEFORE a fallback pair is emitted in its place, not after the pair (08f task 9)', async () => {
    // H moves verbatim to the end and is held; A and B are an unrelated
    // positional fallback pair; W is re-aligned with its edit. Emitting the
    // fallback's pairs first and its held records after put H behind A on the
    // original side, so reject-all rebuilt `OPEN, A, H, W`. The document-order
    // check in `run` is the assertion; the moves are checked as well.
    const held = 'Records means all books and ledgers kept by the Seller.';
    const a = 'Deliverables means the documents to be produced under clause 6.';
    const b = 'Milestone means a date by which a Deliverable must be accepted.';
    const tail = 'Annexes are listed in the order in which they are cited.';
    const ops = await run(
      [OPENING, held, a, WORKING_DAY, CLOSING, tail],
      [OPENING, b, WORKING_DAY_EDITED, CLOSING, tail, held],
    );
    expect(ops.filter((o) => o.type === 'move')).toHaveLength(2);
    expect(spell(ops.find((o) => o.type === 'replace' && o.origRange?.start === 3)?.wordDiffs)).toContain(
      '{+and London +}',
    );
  });

  it('a table unit keeps its stretch positional', async () => {
    // A Markdown table is one alignment unit (ADR-0027) with its own redline
    // author, and its normText is its cell text, so it can out-score a
    // paragraph. A stretch holding one is left exactly as positional pairing
    // leaves it.
    const table = ['| Item | Price |', '| --- | --- |', '| Goods | S$1,000 |'].join('\n');
    const tableEdited = ['| Item | Price |', '| --- | --- |', '| Goods | S$1,200 |'].join('\n');
    const ops = await run(
      [OPENING, APPROVED, table, CLOSING],
      [OPENING, tableEdited, CLOSING],
    );
    // Positional: APPROVED is paired with the table (and demoted); the old table
    // is deleted. The table unit's own redline is 13b's (gated) concern.
    expect(types(ops)).toEqual(['equal', 'delete', 'insert', 'delete', 'equal']);
  });

  it('a gap with nothing confident in it is byte-identical to positional pairing', async () => {
    const ops = await run(
      [OPENING, APPROVED, RECORDS, CLOSING],
      [OPENING, WORKING_DAY, 'Price means the price stated in the Order.', CLOSING],
    );
    expect(types(ops)).toEqual(['equal', 'delete', 'insert', 'delete', 'insert', 'equal']);
  });
});

describe('ADR-0039 invariants over generated documents', () => {
  it('600 generated pairs: both sides walked in document order, every paragraph present once', async () => {
    // The shapes that broke the first two versions — held moves, repeated
    // boilerplate, edited neighbours, merges — mixed at random. Deterministic.
    const pool = [
      APPROVED,
      WORKING_DAY,
      RECORDS,
      'Price means the price stated in the Order.',
      'Order means a purchase order issued by the Buyer.',
      'Not used.',
      'Deliverables means the documents to be produced under clause 6.',
      'Milestone means a date by which a Deliverable must be accepted.',
      'The Seller shall deliver the Goods to the Delivery Address on the Delivery Date.',
      'Any invoice not paid when due bears interest at four per cent a year above base rate.',
    ];
    const edit = (t: string): string => t.replace(/ the /, ' the relevant ').replace(/\.$/, ' in full.');
    let seed = 20261001;
    const rnd = (): number => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
    const pick = <T,>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)]!;

    for (let n = 0; n < 600; n++) {
      const original = [OPENING, ...Array.from({ length: 3 + Math.floor(rnd() * 8) }, () => pick(pool)), CLOSING];
      const revised: string[] = [];
      const moved: string[] = [];
      for (const p of original) {
        const roll = rnd();
        if (p === OPENING || p === CLOSING) revised.push(p);
        else if (roll < 0.15) continue; // deleted
        else if (roll < 0.25) moved.push(p); // moved verbatim to the end
        else if (roll < 0.55) revised.push(edit(p)); // edited
        else revised.push(p);
        if (rnd() < 0.4) revised.push(pick(pool)); // inserted
      }
      revised.splice(revised.length - 1, 0, ...moved);

      const ops = await run(original, revised); // asserts document order
      const seen = { o: new Set<number>(), r: new Set<number>() };
      for (const op of ops) {
        for (const [range, set] of [[op.origRange, seen.o], [op.revRange, seen.r]] as const) {
          if (range === null) continue;
          for (let i = range.start; i <= range.end; i++) {
            expect(set.has(i), `record ${i} claimed twice in case ${n}`).toBe(false);
            set.add(i);
          }
        }
      }
      expect(seen.o.size, `original records lost in case ${n}`).toBe(original.length);
      expect(seen.r.size, `revised records lost in case ${n}`).toBe(revised.length);
    }
  });
});

