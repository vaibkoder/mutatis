/**
 * Feature 08 — Compare §5.5 · the two pipeline wirings of the mark-order pass.
 * ADR-0043.
 *
 * `insertBeforeDelete` is applied at exactly two places: Stage 6.5 of
 * `runComparePipeline` (which serves Word-vs-Word, Text-vs-text **and** the
 * batch driver, which calls it directly) and Stage 6.5 of `runPipelinePdf`.
 *
 * These guards drive those two functions **end to end**. They deliberately do
 * NOT call `insertBeforeDelete` with a context built here: that is exactly how
 * `runPipeline`'s own numbering seam stayed unguarded through several reviews
 * in the numbering work — everyone who wrote a test called the sub-function directly.
 *
 * Each guard was run with its one wiring line deleted and shown to fail; both
 * transcripts are in the plan's Evidence.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve, relative, sep } from 'node:path';

import { getCompareConfig, SG_PROFILE } from '@shared/config/index.js';
import type { CompareConfig } from '@shared/config/index.js';
import type { DiffOp, WordDiff } from '@features/compare/model.js';
import { runComparePipeline, runPipelinePdf } from '@features/compare/pipeline.js';
import { extractTextRecords } from '@features/compare/io/text/extract-text-records.js';

const CONFIG: CompareConfig = getCompareConfig(SG_PROFILE);
const AUTHOR = 'Mutatis Compare';
const DATE = '2026-09-17T00:00:00.000Z';

const ORIGINAL = 'The Vendor grants the Purchaser an exclusive licence over the Business IP.';
const REVISED = 'The Vendor grants the Purchaser a non-exclusive licence over the Business IP.';

/** The op carrying a word diff, of which these fixtures produce exactly one. */
function theReplace(ops: readonly DiffOp[]): WordDiff[] {
  const withDiffs = ops.filter((op) => op.wordDiffs !== undefined);
  expect(withDiffs, 'fixture no longer produces exactly one word-diffed op').toHaveLength(1);
  return withDiffs[0]!.wordDiffs!;
}

/** Ops as a string, so a failure reads as the order it is about. */
const order = (diffs: readonly WordDiff[]): string => diffs.map((d) => d.op).join(',');

describe('Stage 6.5 wiring — runComparePipeline (Word, Text, batch)', () => {
  it('leaves every insertion before the deletion it replaces', async () => {
    const result = await runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    });

    const diffs = theReplace(result.ops);
    // The shape this fixture produces, pinned so the assertion below cannot go
    // vacuous if the diff changes: one replace block, nothing else.
    expect(order(diffs)).toBe('eq,ins,del,eq');
    for (let i = 0; i < diffs.length - 1; i += 1) {
      expect(
        diffs[i]!.op === 'del' && diffs[i + 1]!.op === 'ins',
        `a deletion still precedes its insertion at ${String(i)}`,
      ).toBe(false);
    }
  });

  it('authors <w:ins> before its adjacent <w:del> in the redline body', async () => {
    const result = await runComparePipeline({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    });

    const xml = result.authoredBodyXml;
    const ins = xml.indexOf('<w:ins ');
    const del = xml.indexOf('<w:del ');
    expect(ins, 'no insertion authored').toBeGreaterThanOrEqual(0);
    expect(del, 'no deletion authored').toBeGreaterThanOrEqual(0);
    expect(ins).toBeLessThan(del);
  });
});

describe('Stage 6.5 wiring — runPipelinePdf', () => {
  it('leaves every insertion before the deletion it replaces', async () => {
    // The PDF path authors no redline, but these `wordDiffs` drive the
    // side-by-side panel, the candidate-differences report and the AI export —
    // which must read in the same order as the Word path's.
    const result = await runPipelinePdf({
      originalParas: extractTextRecords(ORIGINAL),
      revisedParas: extractTextRecords(REVISED),
      config: CONFIG,
      author: AUTHOR,
      date: DATE,
    });

    const diffs = theReplace(result.ops);
    expect(order(diffs)).toBe('eq,ins,del,eq');
    for (let i = 0; i < diffs.length - 1; i += 1) {
      expect(
        diffs[i]!.op === 'del' && diffs[i + 1]!.op === 'ins',
        `a deletion still precedes its insertion at ${String(i)}`,
      ).toBe(false);
    }
  });
});

describe('Stage 6.5 runs EXACTLY ONCE per op', () => {
  // `insertBeforeDelete` is not idempotent on two changed regions with no
  // equality between them: `[-A-]{+B+}[-C-]{+D+}` → `{+B+}[-A-]{+D+}[-C-]` →
  // `{+BD+}[-AC-]`. Both sides still reconstruct, so **the never-suppress guard
  // cannot see it** — and neither can anything else. Measured: add a second
  // `applyMarkOrder(ops);` beside the first and the full suite stays green.
  //
  // That is why this needs a SOURCE guard and not a behavioural one. A
  // behavioural fixture would need the chained shape, and it does not occur —
  // 0 of 3,500+ word-diff streams across the suite, because `diff_cleanupMerge`
  // leaves an equality between changed regions. The invariant is unfalsifiable
  // from the outside, so a count of the call sites is the only instrument. The
  // repo already owns this shape: `batch.acceptance.test.ts` → 'the bundle
  // driver runs a pipeline from exactly one place'.
  //
  // ── Why this sweeps all of src/ and resolves the local BINDING ────────────
  // This guard has now been defeated three times, each by a mechanism the
  // previous version did not model. All three are worth keeping, because the
  // pattern is the point: a guard is only as wide as the thing it looks for.
  //
  //   1. It matched the byte string `applyMarkOrder(ops);`, so
  //      `applyMarkOrder( ops );` — same call, two spaces — was invisible.
  //   2. Its second test was titled "from exactly one place in src/" and read
  //      ONLY `pipeline.ts`, so a call added in `batch/driver.ts` passed.
  //   3. Widened to `src/features/compare` and to a call SHAPE, it still missed
  //      BOTH of these:
  //        (a) `import { insertBeforeDelete as reorderMarks }` — an alias never
  //            writes the exported name, so no shape matcher keyed on that name
  //            can see the call;
  //        (b) a call site in `src/features/ai-export/...`, which already
  //            imports from `@features/compare`. The barrel makes the symbol
  //            reachable from ANY module under `src/`, so "one feature" was
  //            never the true scope.
  //
  // So: sweep `src/`; resolve what the symbol is called IN EACH FILE (named
  // import, aliased import, or namespace import) and count references to THAT
  // binding; and count a bare reference, not a call, because
  // `const f = insertBeforeDelete;` is a call site one line later. Comments and
  // the import statement itself are stripped first so prose about the pass does
  // not register as a use.
  //
  // The transitive hole — a NEW barrel re-exporting it under another name, which
  // no amount of alias resolution here would follow — is closed by forbidding
  // the mechanism instead of chasing it: the second assertion pins which files
  // may re-export the symbol at all.
  const SRC = resolve(__dirname, '../../../src');
  const SYMBOL = 'insertBeforeDelete';
  const DEFINITION = 'features/compare/diff/mark-order.ts';
  const BARREL = 'features/compare/diff/index.ts';

  /** Every `.ts`/`.tsx` under `src/`, with its path relative to `src/`. */
  function allSources(): Array<{ rel: string; code: string }> {
    const out: Array<{ rel: string; code: string }> = [];
    const walk = (dir: string): void => {
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = resolve(dir, entry.name);
        if (entry.isDirectory()) walk(full);
        else if (entry.name.endsWith('.ts') || entry.name.endsWith('.tsx')) {
          out.push({ rel: relative(SRC, full).split(sep).join('/'), code: readFileSync(full, 'utf8') });
        }
      }
    };
    walk(SRC);
    return out;
  }

  /** Drop block and line comments, so prose about the pass is not a use of it. */
  function stripComments(code: string): string {
    return code.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/.*$/gm, '$1');
  }

  /**
   * What `insertBeforeDelete` is called in this file, however it was imported.
   * Returns the local names to look for — `[]` when the file never imports it.
   */
  function localBindings(code: string): string[] {
    const names: string[] = [];
    const named = /import\s+(?:type\s+)?\{([^}]*)\}\s*from\s*['"][^'"]*['"]/g;
    for (const m of code.matchAll(named)) {
      for (const spec of m[1].split(',')) {
        const alias = new RegExp(`^\\s*${SYMBOL}(?:\\s+as\\s+([A-Za-z_$][\\w$]*))?\\s*$`).exec(spec);
        if (alias) names.push(alias[1] ?? SYMBOL);
      }
    }
    const ns = /import\s+\*\s+as\s+([A-Za-z_$][\w$]*)\s+from\s*['"]([^'"]*)['"]/g;
    for (const m of code.matchAll(ns)) {
      if (/diff(\/index)?(\.js)?$|mark-order/.test(m[2])) names.push(`${m[1]}\\.${SYMBOL}`);
    }
    return names;
  }

  /** References to a local binding, with imports and comments already removed. */
  function referencesTo(code: string, binding: string): number {
    const body = stripComments(code).replace(/import[\s\S]*?from\s*['"][^'"]*['"];?/g, ' ');
    return (body.match(new RegExp(`(?<![\\w$.])${binding}(?![\\w$])`, 'g')) ?? []).length;
  }

  /** Call sites of `name(`, tolerant of whitespace inside and around the parens. */
  function callsTo(code: string, name: string): number {
    return (code.match(new RegExp(`\\b${name}\\s*\\(`, 'g')) ?? []).length;
  }

  it('calls insertBeforeDelete from exactly one place in the whole of src/', () => {
    // Not one file, and not one FEATURE — `diff/index.ts` exports it, so any
    // module under `src/` is one import away. Two earlier versions of this test
    // claimed a scope wider than the one they read.
    const sites = allSources()
      .filter((f) => f.rel !== DEFINITION && f.rel !== BARREL)
      .map((f) => ({
        rel: f.rel,
        n: localBindings(f.code).reduce((t, b) => t + referencesTo(f.code, b), 0),
      }))
      .filter((f) => f.n > 0);

    expect(sites.map((f) => `${f.rel}:${String(f.n)}`)).toEqual([
      'features/compare/pipeline.ts:1',
    ]);
  });

  it('is re-exported from exactly one barrel, so no alias can be smuggled in', () => {
    // The one hole alias resolution above cannot close: a NEW barrel doing
    // `export { insertBeforeDelete as reorderMarks }` would let a consumer
    // import a name this guard has never heard of. Forbid the mechanism.
    const exporters = allSources()
      .filter((f) => f.rel !== DEFINITION)
      .filter((f) => new RegExp(`export[\\s\\S]{0,200}?\\b${SYMBOL}\\b`).test(stripComments(f.code)))
      .map((f) => f.rel);

    expect(exporters).toEqual([BARREL]);
  });

  it('applies the pass once per pipeline, and from nowhere else', () => {
    const code = readFileSync(resolve(SRC, 'features/compare/pipeline.ts'), 'utf8');

    // Exactly two calls in the file: one per pipeline entry point.
    expect(callsTo(code, 'applyMarkOrder')).toBe(3); // 2 calls + the definition

    // …and ONE in each function body, so two calls in one pipeline — the
    // corruption this guard exists for — cannot satisfy the total. Since
    // ADR-0046 the Word/Text core's stages live in `analyseRecords`, which
    // `runComparePipeline` calls once; the Word table stage calls it again for
    // each cell's OWN records, whose ops are not the body's — still once per op.
    const analyse = code.slice(
      code.indexOf('export async function analyseRecords'),
      code.indexOf('export async function runComparePipeline'),
    );
    const core = code.slice(
      code.indexOf('export async function runComparePipeline'),
      code.indexOf('async function applyWordDiffs'),
    );
    const pdf = code.slice(code.indexOf('export async function runPipelinePdf'));
    expect(callsTo(analyse, 'applyMarkOrder')).toBe(1);
    expect(callsTo(core, 'applyMarkOrder')).toBe(0);
    expect(callsTo(core, 'analyseRecords')).toBe(1);
    expect(callsTo(pdf, 'applyMarkOrder')).toBe(1);

    // The helper itself must stay a plain per-op loop: a second application
    // hidden INSIDE it would satisfy every count above.
    const body = code.slice(
      code.indexOf('function applyMarkOrder'),
      code.indexOf('Replace every {@link PARA_SEAM}'),
    );
    expect(callsTo(body, 'insertBeforeDelete')).toBe(1);
  });
});
