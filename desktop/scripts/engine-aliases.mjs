import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const desktopRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * WHERE THE ENGINE COMES FROM.
 *
 * The desktop app has no engine of its own. It reads the comparison engine
 * straight from the repository root (`../src`, `../public/fonts`) — the same
 * source the Word add-in and web app build from — so there is exactly one copy
 * of it.
 */
export const engineRoot = resolve(desktopRoot, '..');

/**
 * Path aliases into the engine, read-only.
 *
 * SINGLE SOURCE OF TRUTH. These must be identical in the bundler, the test
 * runner and tsconfig.json, and they fail INDEPENDENTLY — an alias can resolve
 * in `tsc` and in `vite build` while the same import throws "Cannot find
 * package" under vitest. tsconfig.json declares its own copy (TypeScript cannot
 * read this file); `tests/engine-contract.test.ts` checks the two agree.
 *
 * Deliberately ABSENT: `@taskpane` and `@host-office`. That is the Preact and
 * Office JS blast radius. Leaving them unaliased is a tripwire — an import that
 * reaches for one fails loudly instead of quietly dragging Preact into a React
 * bundle.
 */
export const engineAliases = {
  '@shared': resolve(engineRoot, 'src/shared'),
  '@features': resolve(engineRoot, 'src/features'),
  '@host-web': resolve(engineRoot, 'src/host-web'),
  '@fonts': resolve(engineRoot, 'public/fonts'),
};

/**
 * Packages that must resolve to the desktop app's copy, not the engine's.
 *
 * A bare `import 'diff-match-patch'` INSIDE engine source resolves by walking up
 * from that file, which would reach the repository root's node_modules — a
 * different install, possibly a different version. `resolve.dedupe` forces every
 * listed package to resolve from the desktop's own node_modules, and guarantees
 * a single copy, which matters for anything holding module state.
 */
export const engineDedupe = [
  'pizzip',
  'fast-xml-parser',
  'diff-match-patch',
  'fast-diff',
  'fuse.js',
  'chrono-node',
  'words-to-numbers',
  // The PDF/layout path.
  'harfbuzzjs',
  'fontkit',
  '@pdfme/pdf-lib',
  // PDF input (reading an uploaded PDF). The engine's loader does a bare
  // `import('pdfjs-dist')`.
  'pdfjs-dist',
  // Preact. Not ours by choice: the pairing sketch imported for the pairing
  // grid reaches the defined-terms barrel, which re-exports its driver, which
  // value-imports the findings store built on Preact signals. So the add-in's
  // UI state layer is in this module graph transitively. Listing it here makes
  // that explicit. The real fix belongs in the engine: an analysis feature's
  // barrel should not re-export its driver.
  '@preact/signals',
  'preact',
];
