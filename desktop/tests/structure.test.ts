import { describe, it, expect } from 'vitest';
import { makeDocxBytes } from './helpers/make-docx.js';
import { readDocx } from '../src/renderer/src/structure/read-docx.js';
import { analyseStructure } from '../src/renderer/src/structure/analyse.js';
import { compose, numberShape, tidyLabel, typedNumber } from '../src/renderer/src/structure/labels.js';
import { buildOutline, outlineToJson, outlineToMarkdown } from '../src/renderer/src/structure/outline.js';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';

/**
 * The contract-structure analyser. These pin the rules a
 * regression would most likely break, on documents small enough to read.
 */

const AGREEMENT = [
  { text: 'SERVICES AGREEMENT', styleId: 'Heading1' },
  'Dated 1 March 2026',
  'PARTIES',
  '(1) ACME LIMITED, a company incorporated in England; and',
  '(2) BETA LLP, a limited liability partnership.',
  'BACKGROUND',
  '(A) The Supplier provides services.',
  '(B) The Customer wishes to buy them.',
  'AGREED TERMS',
  '1. Definitions',
  '1.1 In this Agreement:',
  '“Business Day” means a day other than a Saturday or Sunday;',
  '“Services” means the services in Schedule 1, which include:',
  '(a) installation; and',
  '(b) support.',
  '2. Payment',
  '2.1 The Customer shall pay the Charges.',
  '(a) within 30 days; and',
  '(b) in pounds sterling.',
  '2.2 Late payment bears interest.',
  'IN WITNESS WHEREOF the parties have signed this Agreement.',
  'Signed by ACME LIMITED',
  'SCHEDULE 1 – THE SERVICES',
  '1. Installation',
  '(a) The Supplier shall install the equipment.',
  '2. Support',
];

function analyse(paras: Parameters<typeof makeDocxBytes>[0]) {
  const read = readDocx(makeDocxBytes(paras));
  return { read, structure: analyseStructure(read.paragraphs) };
}

describe('labels', () => {
  it('composes citations the way a lawyer writes them', () => {
    expect(compose(['1.', '1.1', '(a)', '(i)'])).toBe('1.1(a)(i)');
    expect(compose(['6.', '1'])).toBe('6(1)');
    expect(compose(['4.2', '14.10'])).toBe('14.10');
    expect(tidyLabel('a)')).toBe('(a)');
    expect(tidyLabel('2.8.')).toBe('2.8');
  });

  it('reads typed numbers strictly', () => {
    expect(typedNumber('4.2\tThe Buyer')?.token).toBe('4.2');
    expect(typedNumber('4.A The Academy Trust must:')?.token).toBe('4.A');
    expect(typedNumber('[10.15\tExport Control')?.token).toBe('10.15');
    expect(typedNumber('5 days after delivery')).toBeNull();
    expect(typedNumber('1976 Act')).toBeNull();
  });

  it('lets Word settle letter-or-numeral', () => {
    expect(numberShape('(i)').cls).toBe('lowerLetter');
    expect(numberShape('(ii)').cls).toBe('lowerRoman');
    expect(numberShape('(i)', 'lowerRoman')).toMatchObject({ cls: 'lowerRoman', certain: true });
  });
});

describe('analyseStructure', () => {
  const { structure } = analyse(AGREEMENT);
  const byStart = (i: number) => structure.nodes.find((n) => n.start === i);
  const cite = (i: number) => structure.paragraphs[i]!.fullLabel;

  it('finds the front matter, parties and recitals', () => {
    expect(byStart(0)?.kind).toBe('front_matter');
    expect(byStart(2)?.kind).toBe('parties');
    expect(cite(3)).toBe('Party 1');
    expect(cite(4)).toBe('Party 2');
    expect(byStart(6)).toMatchObject({ kind: 'recital', fullLabel: 'Recital A' });
    expect(byStart(7)).toMatchObject({ kind: 'recital', fullLabel: 'Recital B' });
  });

  it('nests clauses, definitions and their limbs', () => {
    expect(byStart(9)).toMatchObject({ kind: 'clause', fullLabel: '1', heading: 'Definitions' });
    expect(byStart(10)).toMatchObject({ kind: 'definitions', fullLabel: '1.1' });
    expect(byStart(11)).toMatchObject({ kind: 'definition', term: 'Business Day' });
    expect(byStart(12)).toMatchObject({ kind: 'definition', term: 'Services' });
    expect(cite(13)).toBe('1.1(a)');
    expect(byStart(13)?.parentId).toBe(byStart(12)?.id);
    expect(cite(17)).toBe('2.1(a)');
    expect(cite(19)).toBe('2.2');
  });

  it('closes the body at the execution block and names the schedule', () => {
    expect(byStart(20)?.kind).toBe('execution');
    expect(byStart(22)).toMatchObject({ kind: 'schedule', fullLabel: 'SCHEDULE 1', heading: 'THE SERVICES' });
    expect(cite(23)).toBe('SCHEDULE 1 para 1');
    expect(cite(24)).toBe('SCHEDULE 1 para 1(a)');
    expect(cite(25)).toBe('SCHEDULE 1 para 2');
  });

  it('gives every non-empty paragraph an owner', () => {
    structure.paragraphs.forEach((p, i) => {
      if (AGREEMENT[i]) expect(p.nodeId, `paragraph ${i}`).not.toBeNull();
    });
  });

  it('keeps a numbered contents list before the parties out of the clause tree', () => {
    const { structure: s } = analyse([
      'CONTENTS',
      '1. Definitions',
      '2. Payment',
      'THIS AGREEMENT is made on 1 March 2026',
      'BETWEEN:',
      '(1) ACME LIMITED; and',
      '(2) BETA LLP.',
      'IT IS AGREED as follows:',
      '1. Definitions',
      '1.1 Words have their meanings.',
    ]);
    expect(s.paragraphs[1]).toMatchObject({ label: '1.', fullLabel: null });
    expect(s.nodes.find((n) => n.start === 1)).toBeUndefined();
    expect(s.paragraphs[5]!.fullLabel).toBe('Party 1');
    expect(s.paragraphs[9]!.fullLabel).toBe('1.1');
  });

  it('reads bold-free, unquoted definitions only when a defining verb follows', () => {
    const { structure: s } = analyse(['1. Interpretation', '“Claim” includes any demand;', 'Claims are rare.']);
    expect(s.nodes.find((n) => n.start === 1)).toMatchObject({ kind: 'definition', term: 'Claim' });
    expect(s.nodes.find((n) => n.start === 2)).toBeUndefined();
  });
});

describe('outline', () => {
  const outline = buildOutline(makeDocxBytes(AGREEMENT));

  it('is plain data that survives structured cloning', () => {
    expect(structuredClone(outline)).toEqual(outline);
  });

  it('copies as an indented Markdown outline with citations', () => {
    const md = outlineToMarkdown(outline, 'Services.docx');
    expect(md).toContain('# Contract structure — Services.docx');
    expect(md).toContain('- **1** — Definitions');
    expect(md).toContain('  - **1.1** — In this Agreement');
    expect(md).toContain('      - **1.1(a)** — installation; and');
    expect(md).toContain('- **SCHEDULE 1** — THE SERVICES');
    expect(md).toContain('  - **SCHEDULE 1 para 1(a)**');
    expect(md).toMatch(/“Business Day”/);
  });

  it('saves as JSON with a node per entry', () => {
    const json = JSON.parse(outlineToJson(outline, 'Services.docx')) as { nodes: { citation: string | null }[] };
    expect(json.nodes.map((n) => n.citation)).toContain('SCHEDULE 1 para 1(a)');
  });
});

describe('structure in a comparison', () => {
  it('rides on the outcome and gives the assistant schedule-qualified citations', async () => {
    const revised = AGREEMENT.map((p) =>
      p === '(a) The Supplier shall install the equipment.' ? '(a) The Supplier shall install and test the equipment.' : p,
    );
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(AGREEMENT),
      revisedBytes: makeDocxBytes(revised),
      date: '2026-03-01T00:00:00.000Z',
    });
    expect(outcome.structure.original?.entries.length).toBeGreaterThan(10);
    expect(outcome.structure.revised?.entries.length).toBeGreaterThan(10);
    const change = outcome.context.changes.find((c) => (c.revised ?? '').includes('install and test'));
    expect(change?.clause).toBe('SCHEDULE 1 para 1(a)');
    expect(outcome.context.blindSpots.join(' ')).toMatch(/structure analysis/);
  });

  it('never stops a comparison when the analyser cannot read a side', async () => {
    const outcome = await runComparison({
      originalBytes: makeDocxBytes(['One paragraph.']),
      revisedBytes: makeDocxBytes(['One paragraph, edited.']),
      date: '2026-03-01T00:00:00.000Z',
    });
    expect(outcome.summary.total).toBeGreaterThan(0);
    expect(outcome.structure).toBeDefined();
  });
});
