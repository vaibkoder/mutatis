/**
 * The reconstruction yardstick: accept-all of the
 * authored redline is the revised document and reject-all the original,
 * CHARACTER FOR CHARACTER and RUN FORMATTING FOR RUN FORMATTING, paragraph
 * count included (`tests/fixtures/reconstruct.ts`).
 *
 * Run over hand-built Word pairs that carry what a firm's documents carry
 * (direct formatting, typographic quotes and dashes, non-breaking spaces, tabs
 * after a manual number, line breaks), the ADR-0032 merge/split shapes, a Word
 * table (ADR-0046), and the synthetic readability corpus.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline, runComparePipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';
import type { CompareParts } from '@features/compare/io/index.js';
import type { ParaRecord } from '@features/compare/model.js';

import {
  reconstruct, isClean, textLost, summarise, paragraphViews, compareViews, type ParaView, type ReconstructionReport,
} from '../../fixtures/reconstruct.js';
import { READABILITY_CASES } from '../../fixtures/readability-corpus.js';
import { inlineParts, run, para, footnoteRef, hyperlink } from '../../fixtures/inline-objects-docx.js';
import { FEE_ROWS, table, tableBundle, para as plainPara } from '../../fixtures/word-table-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);
const DATE = '2026-10-26T00:00:00Z';

const body = (records: readonly ParaRecord[]): string => `<w:body>${records.map((r) => r.rawXml).join('')}</w:body>`;

async function compareParts(o: CompareParts, r: CompareParts): Promise<ReconstructionReport> {
  const result = await runPipeline({ originalParts: o, revisedParts: r, config: CONFIG, author: 'Mutatis', date: DATE });
  return reconstruct(result.authoredBodyXml, body(result.originalParas), body(result.revisedParas));
}

async function compareText(o: string, r: string): Promise<ReconstructionReport> {
  const originalParas = extractTextRecords(o);
  const revisedParas = extractTextRecords(r);
  const result = await runComparePipeline({ originalParas, revisedParas, config: CONFIG, author: 'Mutatis', date: DATE });
  return reconstruct(result.authoredBodyXml, body(originalParas), body(revisedParas));
}

/** Both views clean, with the reason in the message when not. */
function expectClean(rep: ReconstructionReport): void {
  expect(isClean(rep.accept), `accept-all: ${summarise(rep.accept)}`).toBe(true);
  expect(isClean(rep.reject), `reject-all: ${summarise(rep.reject)}`).toBe(true);
}

/**
 * Typography as Word writes it — the character itself, never a character
 * reference (the shared parser does not decode a numeric character reference, and Word never
 * writes one).
 */
const ch = (code: number): string => String.fromCharCode(code);
const RSQUO = ch(0x2019);
const LDQUO = ch(0x201c);
const RDQUO = ch(0x201d);
const NDASH = ch(0x2013);
const MDASH = ch(0x2014);
const NBSP = ch(0xa0);

/** A run with direct formatting. */
const fmt = (rPr: string, text: string): string =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tab = '<w:r><w:tab/></w:r>';
const br = '<w:r><w:br/></w:r>';
const B = '<w:b/>';
const I = '<w:i/>';

/** A paragraph before the pair, so the pair is neither first nor last of its part. */
const OPEN = para(run('This Agreement is made on the date set out above.'));

// ---------------------------------------------------------------------------
// The yardstick itself — what it sees, and what it does not excuse
// ---------------------------------------------------------------------------

describe('the yardstick sees what a reader sees', () => {
  const views = (inner: string) => paragraphViews(`<w:body><w:p>${inner}</w:p></w:body>`);

  it('reads raw characters: a curly apostrophe, a tab, a break and an NBSP are themselves', () => {
    const [v] = views(run(`Seller${RSQUO}s`) + tab + run(`S$${NBSP}1`) + br + run('end'));
    expect(v!.text).toBe(`Seller${RSQUO}s\tS$${NBSP}1\nend`);
  });

  it('a curly apostrophe written straight is a difference, not a match', () => {
    const rep = compareViews(views(run("Seller's")), views(run(`Seller${RSQUO}s`)));
    expect(rep.textDiffers).toEqual([0]);
    expect(isClean(rep)).toBe(false);
  });

  it('bold lost on one word is a formatting difference; rsids and rPrChange are not', () => {
    const lost = compareViews(views(run('The Seller')), views(fmt(B, 'The Seller')));
    expect(lost.formattingDiffers).toEqual([0]);
    const same = compareViews(
      views('<w:r w:rsidR="00AB"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:rsid="1"/><w:b/></w:rPr><w:t>X</w:t></w:r>'),
      views('<w:r><w:rPr><w:b/><w:rFonts w:hAnsi="Arial" w:ascii="Arial"/><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr><w:t>X</w:t></w:r>'),
    );
    expect(isClean(same)).toBe(true);
  });

  it('a blank paragraph is not counted; a lost paragraph is', () => {
    const rep = compareViews(paragraphViews('<w:body><w:p/><w:p>' + run('One') + '</w:p></w:body>'),
      paragraphViews('<w:body><w:p>' + run('One') + '</w:p><w:p>' + run('Two') + '</w:p></w:body>'));
    expect(rep.actualParagraphs).toBe(1);
    expect(rep.expectedParagraphs).toBe(2);
    expect(isClean(rep)).toBe(false);
  });

  it('reject-all excuses an unchanged stretch spelled differently ONLY where accept-all proved it, never a struck one', () => {
    // The two documents spell the unchanged "the  Buyer" differently; the redline
    // authors it once, from the revised side — right for accept-all.
    const authored = `<w:body><w:p>${run('the Buyer shall pay')}<w:ins w:id="1" w:author="a" w:date="d">${run(' promptly')}</w:ins></w:p></w:body>`;
    const ok = reconstruct(authored, `<w:body><w:p>${run('the  Buyer shall pay')}</w:p></w:body>`,
      `<w:body><w:p>${run('the Buyer shall pay promptly')}</w:p></w:body>`);
    expect(ok.reject.unrevisedOnly).toEqual([0]);
    expect(isClean(ok.reject)).toBe(true);
    // A struck character spelled wrongly is a loss, folded or not.
    const struck = `<w:body><w:p>${run('the Buyer shall pay')}<w:del w:id="1" w:author="a" w:date="d"><w:r><w:delText xml:space="preserve"> Seller's</w:delText></w:r></w:del></w:p></w:body>`;
    const bad = reconstruct(struck, `<w:body><w:p>${run(`the Buyer shall pay Seller${RSQUO}s`)}</w:p></w:body>`,
      `<w:body><w:p>${run('the Buyer shall pay')}</w:p></w:body>`);
    expect(textLost(bad.reject)).toEqual([0]);
    // And an unrevised character accept-all got wrong too is not excused.
    const both = `<w:body><w:p>${run("the Buyer's sum")}<w:ins w:id="1" w:author="a" w:date="d">${run(' due')}</w:ins></w:p></w:body>`;
    const wrong = reconstruct(both, `<w:body><w:p>${run(`the Buyer${RSQUO}s sum`)}</w:p></w:body>`,
      `<w:body><w:p>${run(`the Buyer${RSQUO}s sum due`)}</w:p></w:body>`);
    expect(textLost(wrong.reject)).toEqual([0]);
  });

  it('only a §5.2 fold-set character is excused: a non-breaking hyphen lost in unchanged words is a loss', () => {
    const del = (t: string): string => `<w:del w:id="2" w:author="a" w:date="d"><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
    const ins = (t: string): string => `<w:ins w:id="1" w:author="a" w:date="d"><w:r><w:t>${t}</w:t></w:r></w:ins>`;
    const rep = reconstruct(`<w:body><w:p>${run('cooperate ')}${ins('today.')}${del('now.')}</w:p></w:body>`,
      `<w:body><w:p>${run('co')}<w:r><w:noBreakHyphen/></w:r>${run('operate now.')}</w:p></w:body>`,
      `<w:body><w:p>${run('cooperate today.')}</w:p></w:body>`);
    expect(isClean(rep.accept)).toBe(true);
    expect(textLost(rep.reject)).toEqual([0]);
  });

  it('nor is a non-fold-set character the view shows: a non-breaking hyphen the original lacks', () => {
    const del = (t: string): string => `<w:del w:id="2" w:author="a" w:date="d"><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
    const ins = (t: string): string => `<w:ins w:id="1" w:author="a" w:date="d"><w:r><w:t>${t}</w:t></w:r></w:ins>`;
    const rep = reconstruct(`<w:body><w:p>${run('co')}<w:r><w:noBreakHyphen/></w:r>${run('operate ')}${ins('today.')}${del('now.')}</w:p></w:body>`,
      `<w:body><w:p>${run('cooperate now.')}</w:p></w:body>`,
      `<w:body><w:p>${run('co')}<w:r><w:noBreakHyphen/></w:r>${run('operate today.')}</w:p></w:body>`);
    expect(isClean(rep.accept)).toBe(true);
    expect(textLost(rep.reject)).toEqual([0]);
  });

  it('a space missing beside a struck word is a loss, though its other neighbour is unchanged', () => {
    const rep = reconstruct(
      `<w:body><w:p>${run('Pay ')}<w:del w:id="1" w:author="a" w:date="d"><w:r><w:delText xml:space="preserve">Seller </w:delText></w:r></w:del>${run('now')}</w:p></w:body>`,
      `<w:body><w:p>${run('Pay  Seller now')}</w:p></w:body>`,
      `<w:body><w:p>${run('Pay now')}</w:p></w:body>`);
    expect(isClean(rep.accept)).toBe(true);
    expect(textLost(rep.reject)).toEqual([0]);
  });

  it('the excused characters are counted, across paragraphs — a substitution counts both sides', () => {
    const two = (a: string, b: string): string => `<w:body><w:p>${run(a)}</w:p><w:p>${run(b)}</w:p></w:body>`;
    const rep = reconstruct(two("the Buyer's  Price", 'paid - in full'),
      two(`the Buyer${RSQUO}s Price`, `paid ${NDASH} in full`), two("the Buyer's  Price", 'paid - in full'));
    expect(rep.reject.unrevisedOnly).toEqual([0, 1]);
    expect(rep.reject.excusedCharacters).toBe(5); // apostrophe ×2 and a space; dash ×2
  });

  it('a revised character is never excused, even one a caller says accept-all proved', () => {
    const view = (text: string, formats: string[], revised: boolean[] = []): ParaView => ({
      text, formats, origins: [...text].map((_, k) => ({ id: `1:${k}`, revised: revised[k] ?? false })),
    });
    const proven = { text: new Set(['1:0', '1:1', '1:2']), format: new Set(['1:0', '1:1', '1:2']) };
    // Formatting: the first character differs; unrevised it is excused, revised it is not.
    const doc = view('ab', ['<w:b>', '<w:b>']);
    expect(compareViews([view('ab', ['', '<w:b>'])], [doc], proven).formattingOnly).toEqual([0]);
    expect(compareViews([view('ab', ['', '<w:b>'], [true])], [doc], proven).formattingDiffers).toEqual([0]);
    // Text: an extra space; unrevised it is excused, revised it is not.
    const spaced = view('a b', ['', '', '']);
    expect(compareViews([spaced], [view('a  b', ['', '', '', ''])], proven).unrevisedOnly).toEqual([0]);
    expect(compareViews([view('a b', ['', '', ''], [false, true])], [view('a  b', ['', '', '', ''])], proven).unrevisedOnly).toEqual([]);
  });

  it('a formatting-only change between the documents is excused on reject-all only where accept-all proved it', () => {
    const one = (inner: string): string => `<w:body><w:p>${inner}</w:p></w:body>`;
    // Proved: accept-all reads as the revised paragraph exactly.
    const proved = reconstruct(one(run('Pay now')), one(fmt(B, 'Pay') + run(' now')), one(run('Pay now')));
    expect(proved.reject.formattingOnly).toEqual([0]);
    expect(proved.reject.formattingDiffers).toEqual([]);
    // Not proved: accept-all is wrong, so nothing it shows is proof.
    const unproved = reconstruct(one(run('Pay now')), one(fmt(B, 'Pay') + run(' now')), one(run('Pay now!')));
    expect(unproved.reject.formattingDiffers).toEqual([0]);
    // A struck word that lost its formatting is a loss.
    const struck = reconstruct(
      one(`${run('Pay ')}<w:del w:id="1" w:author="a" w:date="d"><w:r><w:delText xml:space="preserve">Buyer </w:delText></w:r></w:del>${run('now')}`),
      one(run('Pay ') + fmt(B, 'Buyer ') + run('now')), one(run('Pay now')));
    expect(struck.reject.formattingDiffers).toEqual([0]);
  });

  it('a long paragraph whose unchanged words differ in typography is judged in linear-ish time', () => {
    // The author runs this on every edited paragraph; a table LCS took 1.5 s
    // and a 121M-cell table on this shape (spec-validator).
    const words = 'The Purchaser shall pay the Price on Completion and not otherwise. '.repeat(165); // ~11,000 characters
    const curly = `${RSQUO}${words}${RSQUO}`;
    const straight = `'${words}'`;
    const started = performance.now();
    const rep = reconstruct(`<w:body><w:p>${run(straight)}</w:p></w:body>`,
      `<w:body><w:p>${run(curly)}</w:p></w:body>`, `<w:body><w:p>${run(straight)}</w:p></w:body>`);
    expect(performance.now() - started).toBeLessThan(250);
    expect(rep.reject.unrevisedOnly).toEqual([0]);
    expect(textLost(rep.reject)).toEqual([]);
  });

  it('a space lost between two struck words is a loss, though the views differ only in whitespace', () => {
    // The original's double space sits inside the deleted words, so no unchanged
    // stretch can explain its absence.
    const authored = `<w:body><w:p>${run('Pay ')}<w:del w:id="1" w:author="a" w:date="d"><w:r><w:delText xml:space="preserve">the Seller</w:delText></w:r></w:del>${run(' now')}</w:p></w:body>`;
    const rep = reconstruct(authored, `<w:body><w:p>${run('Pay the  Seller now')}</w:p></w:body>`,
      `<w:body><w:p>${run('Pay now')}</w:p></w:body>`);
    expect(rep.reject.whitespaceOnly).toEqual([0]);
    expect(textLost(rep.reject)).toEqual([0]);
  });
});

// ---------------------------------------------------------------------------
// Hand-built Word pairs (invented text)
// ---------------------------------------------------------------------------

interface WordCase {
  id: string;
  what: string;
  original: string[];
  revised: string[];
}

const WORD_CASES: WordCase[] = [
  {
    id: 'R1', what: 'a bold defined term kept in an edited clause',
    original: [OPEN, para(fmt(B, 'The Seller') + run(' shall pay the Price.'))],
    revised: [OPEN, para(fmt(B, 'The Seller') + run(' shall pay the Purchase Price.'))],
  },
  {
    id: 'R2', what: 'struck words keep the ORIGINAL run formatting, inserted words the revised',
    original: [OPEN, para(run('Notice is given ') + fmt(I, 'by post') + run(' to the Buyer.'))],
    revised: [OPEN, para(run('Notice is given ') + fmt(B, 'by email') + run(' to the Buyer.'))],
  },
  {
    id: 'R3', what: 'curly apostrophe, curly quotes and an en dash in the unchanged words',
    original: [OPEN, para(run(`The Seller${RSQUO}s ${LDQUO}Warranties${RDQUO} apply 1${NDASH}3 May.`))],
    revised: [OPEN, para(run(`The Seller${RSQUO}s ${LDQUO}Warranties${RDQUO} apply 1${NDASH}3 June.`))],
  },
  {
    id: 'R4', what: 'a non-breaking space in an amount',
    original: [OPEN, para(run(`The fee is S$${NBSP}1,000 payable now.`))],
    revised: [OPEN, para(run(`The fee is S$${NBSP}1,000 payable on demand.`))],
  },
  {
    id: 'R5', what: 'a tab after a manual number, the clause edited',
    original: [OPEN, para(run('1.1') + tab + run('The Buyer shall pay.'))],
    revised: [OPEN, para(run('1.1') + tab + run('The Buyer shall pay promptly.'))],
  },
  {
    id: 'R6', what: 'a manual number in a bold run',
    original: [OPEN, para(fmt(B, '2.3 ') + run('The Lender may assign.'))],
    revised: [OPEN, para(fmt(B, '2.3 ') + run('The Lender may assign or transfer.'))],
  },
  {
    id: 'R7', what: 'a line break inside an edited paragraph',
    original: [OPEN, para(run('Attention: Company Secretary') + br + run('Address: 1 Raffles Place'))],
    revised: [OPEN, para(run('Attention: General Counsel') + br + run('Address: 1 Raffles Place'))],
  },
  {
    id: 'R8', what: 'a trailing space the paragraph ends with',
    original: [OPEN, para(run('The Borrower shall repay the Loan. '))],
    revised: [OPEN, para(run('The Borrower shall repay the Loan in full. '))],
  },
  {
    id: 'R9', what: 'a run split inside an edit: bold then plain in one word group',
    original: [OPEN, para(run('pay the ') + fmt(B, 'Purchase') + run(' Price on Completion.'))],
    revised: [OPEN, para(run('pay the ') + fmt(B, 'Purchase') + run(' Consideration on Completion.'))],
  },
  {
    id: 'R10', what: 'an em dash in the unchanged words (one character, four in normText)',
    original: [OPEN, para(run(`The Buyer${MDASH}and only the Buyer${MDASH}may elect.`))],
    revised: [OPEN, para(run(`The Buyer${MDASH}and only the Buyer${MDASH}may decide.`))],
  },
  {
    id: 'R11', what: 'a footnote reference beside an italic word (the object-aware path)',
    original: [OPEN, para(fmt(I, 'Notice') + footnoteRef() + run(' is given by post.'))],
    revised: [OPEN, para(fmt(I, 'Notice') + footnoteRef() + run(' is given by email.'))],
  },
  {
    id: 'R12', what: 'a hyperlink beside an edit (the object-aware path)',
    original: [OPEN, para(fmt(B, 'See ') + hyperlink('the portal') + run(' for terms.'))],
    revised: [OPEN, para(fmt(B, 'See ') + hyperlink('the portal') + run(' for the terms.'))],
  },
  {
    id: 'R13', what: 'two sides spelling an unchanged stretch with different whitespace (the known hazard)',
    original: [OPEN, para(run('The  Buyer shall pay.'))],
    revised: [OPEN, para(run('The Buyer shall pay promptly.'))],
  },
  {
    id: 'R15', what: 'a leading tab, the first word replaced',
    original: [OPEN, para(tab + run('Old wording of the clause.'))],
    revised: [OPEN, para(tab + run('New wording of the clause.'))],
  },
  {
    id: 'R16', what: 'a revised clause cut short, ending with a space the original had mid-sentence',
    original: [OPEN, para(run('The Buyer shall pay the Price. The Seller shall deliver the Shares.'))],
    revised: [OPEN, para(run('The Buyer shall pay the Price. '))],
  },
  {
    id: 'R17', what: 'a non-breaking hyphen in the unchanged words, in its own run and inside one',
    original: [OPEN, para(run('co') + '<w:r><w:noBreakHyphen/></w:r>' + run('operate fully and re') + '<w:r><w:t>-</w:t><w:noBreakHyphen/><w:t>sign now.</w:t></w:r>')],
    revised: [OPEN, para(run('co') + '<w:r><w:noBreakHyphen/></w:r>' + run('operate fully and re') + '<w:r><w:t>-</w:t><w:noBreakHyphen/><w:t>sign today.</w:t></w:r>')],
  },
  {
    id: 'R18', what: 'a footnote added inside a link beside a bold word (a property op, its words rebuilt from source)',
    original: [OPEN, para(fmt(B, 'Seller') + run(' see ') + '<w:hyperlink r:id="rId9">' + run('the portal') + '</w:hyperlink>' + run(' now.'))],
    revised: [OPEN, para(fmt(B, 'Seller') + run(' see ') + '<w:hyperlink r:id="rId9">' + run('the portal') + footnoteRef() + '</w:hyperlink>' + run(' now.'))],
  },
  {
    id: 'R19', what: 'a character outside the Basic Multilingual Plane beside a bold word (two UTF-16 code units)',
    original: [OPEN, para(run(`The ${String.fromCodePoint(0x20bb7)} family `) + fmt(B, 'Seller') + run(' shall deliver the Shares.'))],
    revised: [OPEN, para(run(`The ${String.fromCodePoint(0x20bb7)} family `) + fmt(B, 'Seller') + run(' shall transfer the Shares.'))],
  },
  {
    id: 'R14', what: 'a whole bold word deleted and an italic one inserted',
    original: [OPEN, para(run('The ') + fmt(B, 'Vendor') + run(' shall deliver the Shares.'))],
    revised: [OPEN, para(run('The ') + fmt(I, 'Seller') + run(' shall deliver the Shares.'))],
  },
];

describe('a hand-built Word pair reconstructs both ways', () => {
  for (const c of WORD_CASES) {
    it(`${c.id}: ${c.what}`, async () => {
      expectClean(await compareParts(inlineParts(...c.original), inlineParts(...c.revised)));
    });
  }
});

describe('a run rebuilt from source carries no input document\'s formatting revision', () => {
  it('nor through a run that holds only a non-breaking hyphen', async () => {
    const hyphen = '<w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Counterparty" w:date="2025-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:noBreakHyphen/></w:r>';
    const { rep, xml } = await authoredClause(
      run('The parties shall co') + hyphen + run('operate and deliver the Shares.'),
      run('The parties shall co') + hyphen + run('operate and transfer the Shares.'),
    );
    expectClean(rep);
    expect(xml).toContain('<w:noBreakHyphen/>');
    expect(xml).not.toContain('w:rPrChange');
  });

  it('nor through such a run at the end of the paragraph', async () => {
    const hyphen = '<w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Counterparty" w:date="2025-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:noBreakHyphen/></w:r>';
    const { rep, xml } = await authoredClause(
      run('The parties shall deliver the Shares to co') + hyphen,
      run('The parties shall transfer the Shares to co') + hyphen,
    );
    expectClean(rep);
    expect(xml).toContain('<w:noBreakHyphen/>');
    expect(xml).not.toContain('w:rPrChange');
  });

  it('a w:rPrChange in an edited clause is not copied: ids stay unique, the bold stays', async () => {
    const changed = (t: string): string =>
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Counterparty" w:date="2025-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
    const { rep, xml } = await authoredClause(
      run('The ') + changed('Seller') + run(' shall deliver the Shares.'),
      run('The ') + changed('Seller') + run(' shall transfer the Shares.'),
    );
    expectClean(rep);
    expect(xml).toContain('<w:b/>');
    expect(xml).not.toContain('w:rPrChange');
  });
});

// ---------------------------------------------------------------------------
// a manual number in an edited clause (ADR-0028 as amended)
// ---------------------------------------------------------------------------

/** The paragraph after OPEN, as authored. */
async function authoredClause(o: string, r: string): Promise<{ rep: ReconstructionReport; xml: string }> {
  const result = await runPipeline({
    originalParts: inlineParts(OPEN, para(o)), revisedParts: inlineParts(OPEN, para(r)),
    config: CONFIG, author: 'Mutatis', date: DATE,
  });
  const rep = reconstruct(result.authoredBodyXml, body(result.originalParas), body(result.revisedParas));
  const paras = result.authoredBodyXml.match(/<w:p[\s>][\s\S]*?<\/w:p>/g) ?? [];
  return { rep, xml: paras[1] ?? '' };
}
const struck = (xml: string): string => [...xml.matchAll(/<w:delText[^>]*>([^<]*)<\/w:delText>/g)].map((m) => m[1]).join('');
const inserted = (xml: string): string =>
  [...xml.matchAll(/<w:ins [^>]*>([\s\S]*?)<\/w:ins>/g)].map((m) => m[1]!.replace(/<[^>]+>/g, '')).join('');

describe('a manual number in an edited clause survives both views', () => {
  it('(a) became (b): the new number inserted, the old struck', async () => {
    const { rep, xml } = await authoredClause(
      run('(a) The Seller shall deliver the Shares.'),
      run('(b) The Seller shall transfer the Shares.'),
    );
    expectClean(rep);
    expect(struck(xml)).toContain('(a)');
    expect(inserted(xml)).toContain('(b)');
  });

  it('1.1 became 1.2, the number in its own run before a tab', async () => {
    expectClean((await authoredClause(
      run('1.1') + tab + fmt(B, 'Price') + run(' is payable on Completion.'),
      run('1.2') + tab + fmt(B, 'Price') + run(' is payable within 30 days of Completion.'),
    )).rep);
  });

  it('a number removed, and a number added', async () => {
    expectClean((await authoredClause(run('(c) The Buyer shall pay the Price.'), run('The Buyer shall pay the Deposit.'))).rep);
    expectClean((await authoredClause(run('The Buyer shall pay the Price.'), run('(c) The Buyer shall pay the Deposit.'))).rep);
  });

  it('an unchanged number is written once, unmarked', async () => {
    const { rep, xml } = await authoredClause(
      run('(a) The Seller shall deliver the Shares.'),
      run('(a) The Seller shall transfer the Shares.'),
    );
    expectClean(rep);
    expect(struck(xml) + inserted(xml)).not.toContain('(a)');
    expect(xml.match(/\(a\)/g)).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// ADR-0032 merge / split — the paragraph count
// ---------------------------------------------------------------------------

const HEAD = 'The Purchaser shall notify the Vendor of any Third Party Claim.';
const TAIL = 'The notice shall specify the nature of the claim in reasonable detail.';
const JOINED = 'The Purchaser shall notify the Vendor of any Third Party Claim, and the notice shall specify the nature of the claim in reasonable detail.';
const CLOSE = 'Closing words of this deed.';

describe('a merge or split reconstructs to the right paragraph count (ADR-0032)', () => {
  const mergeParts = (texts: string[]): CompareParts => extractParts(tableBundle([OPEN, ...texts.map((t) => plainPara(t)), plainPara(CLOSE)].join('')));

  it('a merge: accept-all is ONE paragraph, reject-all two', async () => {
    expectClean(await compareParts(mergeParts([HEAD, TAIL]), mergeParts([JOINED])));
  });

  it('a split: accept-all is two paragraphs, reject-all one', async () => {
    expectClean(await compareParts(mergeParts([JOINED]), mergeParts([HEAD, TAIL])));
  });

  it('a split of a clause with a bold word keeps the bold both ways (rebuilt from source)', async () => {
    const parts = (...ps: string[]): CompareParts => inlineParts(OPEN, ...ps);
    expectClean(await compareParts(
      parts(para(fmt(B, 'Notices') + run(' shall be in writing. They may be sent by email.'))),
      parts(para(fmt(B, 'Notices') + run(' shall be in writing.')), para(run('They may be sent by email.'))),
    ));
  });

  it('a merge holding a footnote is rebuilt from text, and its changed number is still marked (ADR-0028 amendment)', async () => {
    const result = await runPipeline({
      originalParts: inlineParts(OPEN,
        para(run('(a) The Purchaser shall notify the Vendor of any Third Party Claim.') + footnoteRef()),
        para(run('The notice shall specify the nature of the claim in reasonable detail.'))),
      revisedParts: inlineParts(OPEN,
        para(run('(b) The Purchaser shall notify the Vendor of any Third Party Claim, and the notice shall specify the nature of the claim in reasonable detail.') + footnoteRef())),
      config: CONFIG, author: 'Mutatis', date: DATE,
    });
    const rep = reconstruct(result.authoredBodyXml, body(result.originalParas), body(result.revisedParas));
    expect(textLost(rep.reject)).toEqual([]);
    expect(textLost(rep.accept)).toEqual([]);
    expect(rep.reject.expectedParagraphs).toBe(rep.reject.actualParagraphs);
    expect(result.authoredBodyXml).toMatch(/<w:delText[^>]*>\(a\) <\/w:delText>/);
  });

  it('a merge whose join became a line break, the words bold (a common Word export shape)', async () => {
    const parts = (...ps: string[]): CompareParts => inlineParts(OPEN, ...ps);
    expectClean(await compareParts(
      parts(para(fmt(B, 'Notices') + run(' shall be in writing.')), para(run('They may be sent by email.'))),
      parts(para(fmt(B, 'Notices') + run(' shall be in writing.') + br + run('They may be sent by email.'))),
    ));
  });
});

// ---------------------------------------------------------------------------
// A Word table (ADR-0046) — must stay green
// ---------------------------------------------------------------------------

describe('a Word table reconstructs both ways (ADR-0046)', () => {
  it('a changed cell in a fee table', async () => {
    const rows = FEE_ROWS.map((r) => [...r]);
    rows[1]![1] = 'S$15,000';
    const tbl = (r: readonly (readonly string[])[]) => extractParts(tableBundle(plainPara('Fees.') + table(r) + plainPara('End.')));
    expectClean(await compareParts(tbl(FEE_ROWS), tbl(rows)));
  });
});

// ---------------------------------------------------------------------------
// The synthetic readability corpus (Text-vs-text)
// ---------------------------------------------------------------------------

/**
 * RECORDS ITS LIMIT (spec 08 §6.1, not this plan's): a clause whose words did
 * not change but whose MANUAL number did — every clause below an inserted or
 * deleted one — is an `equal` op, emitted from the revised side, so reject-all
 * reads the revised numbers there. Exactly those paragraphs, and nothing else.
 */
const RENUMBERED_UNCHANGED = new Map<string, number>([['C18', 2], ['C19', 2], ['C20', 2], ['C28', 2]]);

describe('the readability corpus reconstructs both ways', () => {
  for (const c of READABILITY_CASES) {
    if (c.kind === 'docx') continue;
    const renumbered = RENUMBERED_UNCHANGED.get(c.id);
    if (renumbered !== undefined) {
      it(`${c.id}: RECORDS ITS LIMIT — only the ${renumbered} renumbered unchanged clauses differ on reject-all`, async () => {
        const rep = await compareText(c.original, c.revised);
        expect(isClean(rep.accept), summarise(rep.accept)).toBe(true);
        expect(textLost(rep.reject)).toHaveLength(renumbered);
        expect(rep.reject.formattingDiffers).toEqual([]);
      });
      continue;
    }
    it(`${c.id}: ${c.title}`, async () => {
      expectClean(await compareText(c.original, c.revised));
    });
  }
});
