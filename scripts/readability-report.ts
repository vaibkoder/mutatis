/**
 * Redline readability report — prints every corpus case as the reader sees it,
 * with the objective metrics beside it.  Run:
 *
 *   npx vite-node --config vitest.config.ts scripts/readability-report.ts [C11 C21 ...]
 *   npx vite-node --config vitest.config.ts scripts/readability-report.ts --json
 *
 * Spec: spec/features/08-compare.md §5.3. Principles: docs/compare-readability-principles.md.
 * Development tool only — not part of the shipped bundle.
 */
import { READABILITY_CASES } from '../tests/fixtures/readability-corpus.js';
import {
  renderCase,
  measure,
  companions,
  formatRendered,
  formatSegs,
} from '../tests/fixtures/readability-metrics.js';

const args = process.argv.slice(2);
const asJson = args.includes('--json');
const only = args.filter((a) => /^C\d\d$/.test(a));
const cases = only.length > 0 ? READABILITY_CASES.filter((c) => only.includes(c.id)) : READABILITY_CASES;

const rows: Array<Record<string, unknown>> = [];

for (const c of cases) {
  const rendered = await renderCase(c);
  const m = measure(c, rendered);
  rows.push({ id: c.id, grade: c.grade, ...m, lines: formatRendered(rendered), ops: rendered.ops });
  if (asJson) continue;

  console.log(`\n${'═'.repeat(78)}`);
  console.log(`${c.id}  ${c.title}   [reviewer: ${c.grade}]`);
  console.log(`      act: ${c.act}`);
  console.log(`${'─'.repeat(78)}`);
  console.log('AS THE READER SEES IT:');
  for (const line of formatRendered(rendered)) console.log('  ' + line);
  if (c.ideal) {
    console.log(`WANTED:${m.idealMatch === true ? '   ✓ matched' : ''}`);
    console.log('  ' + formatSegs(c.ideal));
  }
  console.log(`${'─'.repeat(78)}`);
  console.log(
    `  marks=${m.marks}  redundancy=${m.redundancy}  charRedundancy=${m.charRedundancy}` +
      `  atomicIntact=${m.atomicIntact}  numbered=${m.numbered}` +
      (m.idealMatch === null ? '' : `  idealMatch=${m.idealMatch}`),
  );
  console.log(`  ops:  ${rendered.ops.join(', ') || 'none'}`);
}

if (asJson) {
  console.log(JSON.stringify({ synthetic: rows }, null, 1));
} else {
  const worst = rows.filter((r) => (r.redundancy as number) > 0.3);
  console.log(`\n${'═'.repeat(78)}`);
  console.log(`${rows.length} case(s).  redundancy > 0.30 in: ${worst.map((r) => `${r.id as string}(${r.redundancy as number})`).join(' ') || 'none'}`);
  const sum = rows.reduce((a, r) => a + (r.redundancy as number), 0);
  console.log(`total redundancy = ${sum.toFixed(2)}   total marks = ${rows.reduce((a, r) => a + (r.marks as number), 0)}`);
  const withIdeal = rows.filter((r) => r.idealMatch !== null);
  const met = withIdeal.filter((r) => r.idealMatch === true);
  console.log(
    `reviewer's wanted rendering met in ${met.length}/${withIdeal.length}: ` +
      `${met.map((r) => r.id as string).join(' ') || 'none'}` +
      `${met.length < withIdeal.length ? `  · still unmet: ${withIdeal.filter((r) => r.idealMatch !== true).map((r) => r.id as string).join(' ')}` : ''}`,
  );
}
