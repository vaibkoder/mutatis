/**
 * ADR-0046 — a seeded random sweep over Word tables.
 *
 * WT1–WT12 pin the shapes someone thought of. This file generates the ones
 * nobody did: a random table (multi-paragraph cells, a nested table, a
 * footnote reference in a cell, merged cells) and one to three random edits
 * (a word changed, a row in or out, a column in or out, cells merged, a
 * paragraph added or removed in a cell, a nested cell edited, a footnote added
 * or removed), then the whole Word-vs-Word pipeline. Every case must:
 *
 *  - reconstruct both ways (WT10): accept-all spells the revised body and
 *    reject-all the original, cell for cell, word for word, object for object;
 *  - mint every revision id once;
 *  - report at least one change (ADR-0008 — never suppress);
 *  - revise the table IN PLACE: exactly one top-level table in the redline. A
 *    second one means the self-check fell back to whole-table-out-and-in, which
 *    is safe but is the outcome this plan exists to avoid on a valid shape.
 *
 * The first run at 2,000 cases found two defects 200 had not: a blank cell
 * paragraph written `<w:p/>` took its revision outside it, and a column added
 * beside a row inserted and deleted paired rows by position. Replay one case
 * with `SEED=<n> npx vitest run <this file>`.
 *
 * Synthetic only — every word below is invented.
 */
import { describe, it, expect } from 'vitest';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import { runPipeline } from '@features/compare/pipeline.js';
import { extractParts } from '@features/compare/io/parts.js';
import { flattenTrackedChanges } from '@features/compare/preprocess/flatten.js';
import { signature } from '@features/compare/output/author-word-table.js';
import { para, table, tableBundle, type CellSpec } from '../../fixtures/word-table-docx.js';

const CONFIG = getCompareConfig(SG_PROFILE);
const CASES = 200;

const WORDS = [
  'fee', 'tranche', 'margin', 'lender', 'agent', 'quarterly', 'annually', 'S$5,000', '0.25%',
  'security', 'trustee', 'commitment', 'utilisation', 'repayment', 'interest', 'period', 'notice',
  'facility', 'borrower', 'guarantor', 'default', 'event', 'date', 'amount', 'schedule',
];

/** mulberry32 — a small seeded PRNG, so a failing case can be replayed by its seed. */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

interface SweepCell {
  paras: string[];
  span: number;
  vMerge?: 'restart' | 'continue';
  footnote: boolean;
  nested: string[][] | null;
}
type SweepTable = SweepCell[][];

const FOOTNOTE = '<w:p><w:r><w:t xml:space="preserve">see note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>';

function render(t: SweepTable): string {
  const width = Math.max(...t.map((row) => row.reduce((n, c) => n + c.span, 0)));
  const cells = t.map((row) =>
    row.map((c): CellSpec => {
      let xml = c.paras.map((p) => para(p)).join('');
      if (c.footnote) xml += FOOTNOTE;
      // A cell must end with a paragraph, so a nested table is followed by one.
      if (c.nested !== null) xml += table(c.nested) + para('');
      if (xml === '') xml = para('');
      return { xml, gridSpan: c.span, ...(c.vMerge === undefined ? {} : { vMerge: c.vMerge }) };
    }),
  );
  return para('The schedule follows.') + table(cells, { gridCols: width }) + para('Nothing further.');
}

function clone(t: SweepTable): SweepTable {
  return t.map((row) => row.map((c) => ({ ...c, paras: [...c.paras], nested: c.nested?.map((r) => [...r]) ?? null })));
}

function generate(next: () => number): SweepTable {
  const pick = <T,>(xs: readonly T[]): T => xs[Math.floor(next() * xs.length)]!;
  const phrase = (): string => Array.from({ length: 1 + Math.floor(next() * 4) }, () => pick(WORDS)).join(' ');
  const rows = 2 + Math.floor(next() * 5);
  const cols = 2 + Math.floor(next() * 3);
  let nestedPlaced = false;
  const t: SweepTable = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, (): SweepCell => {
      const nested = !nestedPlaced && next() < 0.08;
      if (nested) nestedPlaced = true;
      return {
        paras: Array.from({ length: next() < 0.2 ? 2 : 1 }, phrase),
        span: 1,
        footnote: next() < 0.06,
        nested: nested ? [[phrase(), phrase()], [phrase(), phrase()]] : null,
      };
    }),
  );
  // A column merged DOWN across two rows — the shape of a shared value.
  if (next() < 0.25) {
    const r = Math.floor(next() * (rows - 1));
    const c = Math.floor(next() * cols);
    t[r]![c]!.vMerge = 'restart';
    t[r + 1]![c]!.vMerge = 'continue';
    t[r + 1]![c]!.paras = [''];
  }
  // A heading or total row merged ACROSS the whole table — common in schedules,
  // and the shape the first real table to take a new column had.
  if (next() < 0.3) t.push([{ paras: [phrase()], span: cols, footnote: false, nested: null }]);
  return t;
}

type Edit = (t: SweepTable, next: () => number) => boolean;

const anyCell = (t: SweepTable, next: () => number): SweepCell => {
  const row = t[Math.floor(next() * t.length)]!;
  return row[Math.floor(next() * row.length)]!;
};
/** Every row one cell per column, or one cell merged across the table — the rows a column edit can run through. */
const columnable = (t: SweepTable): boolean => {
  const width = Math.max(...t.map((row) => row.reduce((n, c) => n + c.span, 0)));
  return t.every((row) => row.every((c) => c.span === 1) ? row.length === width : row.length === 1 && row[0]!.span === width);
};
const spanning = (row: SweepCell[]): boolean => row.length === 1 && row[0]!.span > 1;
const freshRow = (t: SweepTable, next: () => number): SweepCell[] => {
  const width = Math.max(...t.map((row) => row.reduce((n, c) => n + c.span, 0)));
  return Array.from({ length: width }, () => ({ paras: [`${WORDS[Math.floor(next() * WORDS.length)]} new`], span: 1, footnote: false, nested: null }));
};

const EDITS: Record<string, Edit> = {
  editWord: (t, next) => {
    const cell = anyCell(t, next);
    const i = Math.floor(next() * cell.paras.length);
    const words = (cell.paras[i] ?? '').split(' ');
    words[Math.floor(next() * words.length)] = `revised${Math.floor(next() * 100)}`;
    cell.paras[i] = words.join(' ');
    return true;
  },
  insertRow: (t, next) => {
    t.splice(Math.floor(next() * (t.length + 1)), 0, freshRow(t, next));
    return true;
  },
  deleteRow: (t, next) => {
    if (t.length <= 2) return false;
    t.splice(Math.floor(next() * t.length), 1);
    return true;
  },
  // A column through a row merged across the table widens or narrows that merge,
  // as Word does.
  addColumn: (t, next) => {
    if (!columnable(t)) return false;
    const width = Math.max(...t.map((row) => row.length));
    const at = Math.floor(next() * (width + 1));
    t.forEach((row, i) => {
      if (spanning(row)) row[0]!.span++;
      else row.splice(at, 0, { paras: [`added ${i}`], span: 1, footnote: false, nested: null });
    });
    return true;
  },
  removeColumn: (t, next) => {
    const width = Math.max(...t.map((row) => row.length));
    if (!columnable(t) || width <= 2) return false;
    const at = Math.floor(next() * width);
    t.forEach((row) => {
      if (spanning(row)) row[0]!.span--;
      else row.splice(at, 1);
    });
    return true;
  },
  mergeCells: (t, next) => {
    const row = t[Math.floor(next() * t.length)]!;
    if (row.length < 2) return false;
    const at = Math.floor(next() * (row.length - 1));
    const [a, b] = [row[at]!, row[at + 1]!];
    if (a.nested !== null || b.nested !== null || a.vMerge !== undefined || b.vMerge !== undefined) return false;
    row.splice(at, 2, { paras: [...a.paras, ...b.paras], span: a.span + b.span, footnote: a.footnote || b.footnote, nested: null });
    return true;
  },
  addPara: (t, next) => {
    const cell = anyCell(t, next);
    cell.paras.splice(Math.floor(next() * (cell.paras.length + 1)), 0, 'an added paragraph');
    return true;
  },
  deletePara: (t, next) => {
    const cell = anyCell(t, next);
    if (cell.paras.length < 2) return false;
    cell.paras.splice(Math.floor(next() * cell.paras.length), 1);
    return true;
  },
  editNested: (t, next) => {
    const cell = t.flat().find((c) => c.nested !== null);
    if (cell === undefined) return false;
    const row = cell.nested![Math.floor(next() * cell.nested!.length)]!;
    row[Math.floor(next() * row.length)] = `nested${Math.floor(next() * 100)}`;
    return true;
  },
  toggleFootnote: (t, next) => {
    const cell = anyCell(t, next);
    cell.footnote = !cell.footnote;
    return true;
  },
};
const EDIT_NAMES = Object.keys(EDITS);

/** Top-level `<w:tbl>` elements in some body XML (nested ones not counted). */
function topLevelTables(xml: string): number {
  let depth = 0;
  let count = 0;
  for (const m of xml.matchAll(/<(\/?)w:tbl>/g)) {
    if (m[1] === '/') depth--;
    else if (depth++ === 0) count++;
  }
  return count;
}

describe('ADR-0046 — seeded random table sweep', () => {
  it(`${CASES} random table edits reconstruct both ways, report, and revise in place`, async () => {
    const failures: string[] = [];
    const used = new Map<string, number>();
    for (let seed = Number(process.env.SEED ?? 1); seed <= Number(process.env.SEED ?? CASES); seed++) {
      const next = rng(seed);
      const original = generate(next);
      const revised = clone(original);
      const applied: string[] = [];
      const want = 1 + Math.floor(next() * 3);
      for (let tries = 0; applied.length < want && tries < 12; tries++) {
        const name = EDIT_NAMES[Math.floor(next() * EDIT_NAMES.length)]!;
        if (EDITS[name]!(revised, next)) applied.push(name);
      }
      const origBody = render(original);
      const revBody = render(revised);
      if (origBody === revBody) continue;
      for (const name of applied) used.set(name, (used.get(name) ?? 0) + 1);

      const result = await runPipeline({
        originalParts: extractParts(tableBundle(origBody)),
        revisedParts: extractParts(tableBundle(revBody)),
        config: CONFIG,
        author: 'Mutatis',
        date: '2026-10-18T00:00:00Z',
      });
      const xml = result.authoredBodyXml;
      const tag = `seed ${seed} [${applied.join(', ')}]`;
      if (signature(flattenTrackedChanges(xml, 'acceptAll')) !== signature(revBody)) failures.push(`${tag}: accept-all ≠ revised`);
      if (signature(flattenTrackedChanges(xml, 'rejectAll')) !== signature(origBody)) failures.push(`${tag}: reject-all ≠ original`);
      const ids = [...xml.matchAll(/<w:(?:ins|del|cellIns|cellDel)\s[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
      if (new Set(ids).size !== ids.length) failures.push(`${tag}: duplicate revision id`);
      if (result.summary.total === 0) failures.push(`${tag}: no change reported`);
      if (topLevelTables(xml) !== 1) failures.push(`${tag}: ${topLevelTables(xml)} top-level tables — not revised in place`);
    }
    expect(failures).toEqual([]);
    // The sweep is only as good as its coverage: every edit kind must have run.
    for (const name of EDIT_NAMES) expect(used.get(name) ?? 0, `edit ${name} never ran`).toBeGreaterThan(5);
  }, 120_000);
});
