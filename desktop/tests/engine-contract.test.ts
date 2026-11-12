import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { runComparison } from '../src/renderer/src/compare/run-comparison.js';
import { makeDocxBytes, SAMPLE_ORIGINAL, SAMPLE_REVISED } from './helpers/make-docx.js';
import { engineAliases, engineDedupe } from '../scripts/engine-aliases.mjs';

/**
 * THE ENGINE CONTRACT.
 *
 * The desktop app wraps the compare engine at the repository root, which the
 * Word add-in also builds from and which changes on its own schedule. This file
 * is what keeps an engine change from silently breaking the desktop app.
 *
 * It has TWO HALVES and they mean opposite things.
 *
 *   PART A — API SURFACE. The modules and functions we import still exist. A
 *   failure here is ALWAYS a real integration break: the engine moved, renamed
 *   or deleted something this app calls.
 *
 *   PART B — OUTPUT SNAPSHOT. What the engine produces for a fixed input. This
 *   is EXPECTED TO FAIL when the engine legitimately improves — that is the
 *   point. It fails loudly so a redline change gets reviewed deliberately rather
 *   than discovered in the UI three weeks later. A failure here is a question,
 *   not a verdict.
 */

/**
 * Every engine module this app imports, enumerated by hand.
 *
 * Enumerated rather than globbed because a glob cannot fail for a module that
 * VANISHED — only for one that broke. The list is short on purpose: this app
 * reaches for the pipeline, the parts reader, the pairing maths, the render
 * model and the config, and nothing else. Anything added here is a widening of
 * the engine coupling and should be a deliberate decision.
 *
 * Note what is NOT here. `@features/numbering/index.js` and the batch
 * `components/` directory are Preact UI modules — importing one drags a second
 * framework into a React bundle, which is why the project rules forbid reaching into
 * `src/features/**` for a driver, a ui-types or a .tsx. An early draft of this
 * test imported it and blew up on its embedded markup, which is the tripwire
 * working exactly as designed.
 */
const REQUIRED_MODULES = [
  '@features/ai-export/serialise/change-set.js',
  '@features/ai-export/types.js',
  '@features/compare/batch/limits.js',
  '@features/compare/batch/pairing/model.js',
  '@features/compare/batch/pairing/propose.js',
  '@features/compare/batch/pairing/score.js',
  '@features/compare/batch/pairing/sketch.js',
  '@features/compare/io/parts.js',
  '@features/compare/model.js',
  '@features/compare/pipeline.js',
  '@shared/config/index.js',
  '@shared/numbering/numbering-scheme.js',
  '@shared/ooxml/bundle.js',
  '@shared/render-core/input/index.js',
  '@shared/render-core/types.js',
];

describe('engine contract — Part A: the API surface we depend on', () => {
  it('declares the same engine location in tsconfig.json as in the aliases', () => {
    // TypeScript cannot read engine-aliases.mjs, so tsconfig.json keeps its own
    // copy of the three source aliases. They must point at the same folders.
    const tsconfig = readFileSync(resolve(__dirname, '../tsconfig.json'), 'utf8');
    for (const alias of ['@shared', '@features', '@host-web'] as const) {
      const m = new RegExp(`"${alias}/\\*":\\s*\\["([^"]+)/\\*"\\]`).exec(tsconfig);
      expect(m, `${alias} in tsconfig.json`).not.toBeNull();
      expect(resolve(__dirname, '..', m![1]!)).toBe(engineAliases[alias]);
    }
  });

  it('has every aliased source tree', () => {
    for (const [alias, target] of Object.entries(engineAliases)) {
      expect(existsSync(target), `${alias} → ${target}`).toBe(true);
    }
  });

  it.each(REQUIRED_MODULES)('can import %s', async (specifier) => {
    const module = await import(/* @vite-ignore */ specifier);
    expect(module).toBeTypeOf('object');
  });

  it('still exports the specific functions this app calls', async () => {
    // Named individually rather than snapshotting every export: a snapshot
    // would fail on any ADDITION upstream, which is noise. These are the ones
    // whose disappearance actually breaks us.
    const pipeline = await import('@features/compare/pipeline.js');
    expect(Object.keys(pipeline).length).toBeGreaterThan(0);

    const parts = await import('@features/compare/io/parts.js');
    expect(parts).toHaveProperty('extractParts');

    const changeSet = await import('@features/ai-export/serialise/change-set.js');
    expect(changeSet).toHaveProperty('serialiseChangeSet');

    const pairing = await import('@features/compare/batch/pairing/propose.js');
    expect(Object.keys(pairing).length).toBeGreaterThan(0);

    const renderInput = await import('@shared/render-core/input/index.js');
    expect(Object.keys(renderInput).length).toBeGreaterThan(0);

    const bundle = await import('@shared/ooxml/bundle.js');
    expect(Object.keys(bundle).length).toBeGreaterThan(0);
  });

  it('installs every deduplicated engine dependency in the desktop app itself', () => {
    // The engine lives at the repository root, next to the Word add-in's own
    // node_modules. A bare import inside engine source would resolve from THAT
    // install unless it is listed in `engineDedupe` — and dedupe only works if
    // the package is installed here. If one were missing, the engine would
    // silently run against the add-in's copy, possibly a different version.
    for (const pkg of engineDedupe) {
      expect(existsSync(resolve(__dirname, '../node_modules', pkg, 'package.json')), pkg).toBe(true);
    }
  });
});

describe('engine contract — Part B: output snapshot (EXPECTED to change)', () => {
  /**
   * A failure in this block is not necessarily a bug. If the engine was just
   * improved, read the diff, decide it is the intended change, update the
   * expectation, and bump again. What must never happen is the redline changing
   * without anyone looking at it.
   */
  const input = {
    originalBytes: makeDocxBytes(SAMPLE_ORIGINAL),
    revisedBytes: makeDocxBytes(SAMPLE_REVISED),
    // Fixed so the result is reproducible run to run.
    date: '2026-01-01T00:00:00.000Z',
    author: 'Test',
  };

  it('is deterministic — same input, same output', async () => {
    // Without this the snapshot below would be noise rather than evidence.
    const first = await runComparison(input);
    const second = await runComparison(input);
    expect(summarise(first)).toEqual(summarise(second));
  });

  it('produces a stable change set for a fixed input', async () => {
    const outcome = await runComparison(input);
    expect(summarise(outcome)).toMatchSnapshot();
  });
});

function summarise(outcome: Awaited<ReturnType<typeof runComparison>>) {
  return {
    // The engine's own count, not ours — this is the number the summary report
    // and the deviation matrix are both derived from.
    changesTotal: outcome.summary.total,
    ops: outcome.ops.length,
    blocks: outcome.model.blocks.length,
  };
}
