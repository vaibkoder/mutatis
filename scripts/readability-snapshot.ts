/**
 * Snapshot the whole readability corpus as it renders right now.
 *
 *   npx vite-node --config vitest.config.ts scripts/readability-snapshot.ts <label>
 *
 * Writes `.redline-poc/runs/<NN>-<label>.json` (gitignored) carrying the case
 * text, the reviewer's grade and wanted rendering, the rendering produced by the
 * current working tree, and the metrics. `scripts/readability-gallery.mjs` turns
 * any set of these into a side-by-side page for a human to judge.
 *
 * Development tool only — not part of the shipped bundle.
 */
import { mkdirSync, writeFileSync, readdirSync, existsSync } from 'node:fs';

import { READABILITY_CASES } from '../tests/fixtures/readability-corpus.js';
import { renderCase, measure, formatRendered, formatSegs } from '../tests/fixtures/readability-metrics.js';

const label = (process.argv[2] ?? 'run').replace(/[^A-Za-z0-9._-]/g, '-');
const dir = '.redline-poc/runs';
mkdirSync(dir, { recursive: true });

const seq = existsSync(dir)
  ? readdirSync(dir).filter((f) => f.endsWith('.json')).length
  : 0;
const file = `${dir}/${String(seq).padStart(2, '0')}-${label}.json`;

const cases = [];
for (const c of READABILITY_CASES) {
  const rendered = await renderCase(c);
  cases.push({
    id: c.id,
    title: c.title,
    act: c.act,
    grade: c.grade,
    cites: c.cites,
    expect: c.expect,
    original: c.original,
    revised: c.revised,
    wanted: c.ideal ? formatSegs(c.ideal) : null,
    lines: formatRendered(rendered),
    ops: rendered.ops,
    metrics: measure(c, rendered),
  });
}

const totals = {
  redundancy: Math.round(cases.reduce((a, c) => a + c.metrics.redundancy, 0) * 100) / 100,
  marks: cases.reduce((a, c) => a + c.metrics.marks, 0),
};

writeFileSync(
  file,
  JSON.stringify({ label, at: new Date().toISOString(), totals, cases }, null, 1),
);
console.log(`${file}   redundancy=${totals.redundancy}  marks=${totals.marks}`);
