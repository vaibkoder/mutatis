/**
 * Module-graph guard for the pairing sketch (ADR-0037)
 *
 * `sketch.ts` is host-free and runs at bundle scale (Compare-at-scale,
 * standalone/web only), so nothing it imports may drag in the add-in's UI
 * state layer. It reused `analyseDefinedTerms`/`analysePartyRole` via each
 * feature's barrel (`index.ts`) — but a feature barrel also re-exports that
 * feature's `clearThenEmit`, which VALUE-imports `@shared/findings`, whose
 * `store.ts` makes a module-level `signal()` call from `@preact/signals` that
 * cannot be tree-shaken. Importing a single named export from a barrel does
 * not protect against this: the barrel module itself still has a static
 * `export { clearThenEmit } from './driver.js'` edge, which a bundler's
 * module graph resolves regardless of which export a given consumer uses.
 *
 * Neither of the two existing safety nets sees this class of leak:
 *  - the four custom invariant lint rules (`eslint-plugins/invariants.cjs`)
 *    check import *sources* against a fixed disallow-list, not transitive
 *    reachability;
 *  - the "office-free web grep" (see `tests/unit/web/web-compare-e2e.test.tsx`)
 *    greps *built* chunks for Office-JS
 *    calls — it has nothing to say about a pure-analysis module pulling in
 *    Preact signals, and it only ever covers the chunks the web entry point
 *    actually loads, not an arbitrary internal module's own graph.
 *
 * So this test resolves the graph itself: starting from `sketch.ts`, it walks
 * every `import`/re-export statement transitively (resolving the
 * `@shared/*`/`@features/*`/`@host-web/*` aliases from `tsconfig.json`, the
 * same source of truth `vite.config.ts` and `vitest.config.ts` mirror), and
 * asserts neither `@shared/findings` nor `@preact/signals` is reachable.
 *
 * `import type`/`export type` edges (and named specifiers marked `type`) are
 * excluded from the walk: TypeScript (and esbuild, which Vite uses) elides
 * these from the emitted JS entirely, so they create no runtime module edge.
 * Several files on this graph (e.g. `party-role/pr1.ts`) do have a type-only
 * import of `Locator` from a feature barrel — that is inert by construction,
 * which is exactly why only VALUE edges are walked here.
 */
import { describe, it, expect, beforeAll } from 'vitest';
import * as ts from 'typescript';
import { readFileSync, existsSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../../'); // tests/unit/compare -> repo root
const ENTRY = join(ROOT, 'src/features/compare/batch/pairing/sketch.ts');

interface AliasRule {
  /** e.g. "@shared/" (the "*" stripped, trailing "/" kept). */
  specifierPrefix: string;
  /** e.g. "src/shared/" (relative to ROOT). */
  targetPrefix: string;
}

/** Read `paths` straight from tsconfig.json — the alias source of truth. */
function loadAliasRules(): AliasRule[] {
  const tsconfig = JSON.parse(readFileSync(join(ROOT, 'tsconfig.json'), 'utf8')) as {
    compilerOptions: { paths: Record<string, string[]> };
  };
  const rules: AliasRule[] = [];
  for (const [key, targets] of Object.entries(tsconfig.compilerOptions.paths)) {
    const target = targets[0];
    if (!key.endsWith('/*') || !target || !target.endsWith('/*')) continue;
    rules.push({
      specifierPrefix: key.slice(0, -1), // "@shared/*" -> "@shared/"
      targetPrefix: target.slice(0, -1), // "src/shared/*" -> "src/shared/"
    });
  }
  return rules;
}

/** Candidate file paths a bare (extension-stripped-to-`.js`) resolved path could be. */
function fileCandidates(rawPath: string): string[] {
  const candidates: string[] = [];
  if (rawPath.endsWith('.js')) {
    const withoutExt = rawPath.slice(0, -'.js'.length);
    candidates.push(`${withoutExt}.ts`, `${withoutExt}.tsx`);
  }
  candidates.push(rawPath, `${rawPath}.ts`, `${rawPath}.tsx`);
  candidates.push(join(rawPath, 'index.ts'), join(rawPath, 'index.tsx'));
  return candidates;
}

function resolveToFile(rawPath: string): string | null {
  for (const candidate of fileCandidates(rawPath)) {
    if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
  }
  return null;
}

/**
 * One VALUE-level module edge out of a source file: `{specifier, isTypeOnly}`
 * has already been resolved to `isTypeOnly === false` by the time it is
 * returned — see `extractValueEdges`.
 */
function extractValueEdges(filePath: string): string[] {
  const text = readFileSync(filePath, 'utf8');
  const sourceFile = ts.createSourceFile(
    filePath,
    text,
    ts.ScriptTarget.Latest,
    /* setParentNodes */ true,
    filePath.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );

  const edges: string[] = [];

  sourceFile.forEachChild(node => {
    // import ... from '...'
    if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
      const specifier = node.moduleSpecifier.text;
      const clause = node.importClause;
      if (!clause) {
        // Side-effect import: `import './x.js'` — always a value edge.
        edges.push(specifier);
        return;
      }
      if (clause.isTypeOnly) return; // `import type {...} from '...'` — fully elided.

      let hasValueBinding = Boolean(clause.name); // default import binding
      const named = clause.namedBindings;
      if (named && ts.isNamespaceImport(named)) {
        hasValueBinding = true; // `import * as ns from '...'`
      } else if (named && ts.isNamedImports(named)) {
        for (const el of named.elements) {
          if (!el.isTypeOnly) hasValueBinding = true;
        }
      }
      if (hasValueBinding) edges.push(specifier);
      return;
    }
    // export {...} from '...' (barrel re-export)
    if (ts.isExportDeclaration(node) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) {
      const specifier = node.moduleSpecifier.text;
      if (node.isTypeOnly) return; // `export type {...} from '...'` — fully elided.

      if (node.exportClause && ts.isNamedExports(node.exportClause)) {
        let hasValueBinding = false;
        for (const el of node.exportClause.elements) {
          if (!el.isTypeOnly) hasValueBinding = true;
        }
        if (hasValueBinding) edges.push(specifier);
      } else {
        // `export * from '...'` (no named clause) — conservatively a value edge.
        edges.push(specifier);
      }
    }
  });

  return edges;
}

interface GraphResult {
  /** Absolute paths of every file reached via a VALUE edge, including the entry file. */
  visitedFiles: Set<string>;
  /** Bare (non-relative, non-aliased) specifiers reached via a VALUE edge — npm packages. */
  externalSpecifiers: Set<string>;
}

/** BFS the VALUE-edge module graph from `entryFile`, resolving local aliases. */
function walkModuleGraph(entryFile: string, aliasRules: AliasRule[]): GraphResult {
  const visitedFiles = new Set<string>();
  const externalSpecifiers = new Set<string>();
  const queue: string[] = [entryFile];

  while (queue.length > 0) {
    const current = queue.pop()!;
    if (visitedFiles.has(current)) continue;
    visitedFiles.add(current);

    for (const specifier of extractValueEdges(current)) {
      if (specifier.startsWith('.')) {
        const raw = join(dirname(current), specifier);
        const resolved = resolveToFile(raw);
        if (!resolved) {
          throw new Error(`Could not resolve relative import "${specifier}" from ${current}`);
        }
        queue.push(resolved);
        continue;
      }
      const alias = aliasRules.find(r => specifier.startsWith(r.specifierPrefix));
      if (alias) {
        const raw = join(ROOT, alias.targetPrefix, specifier.slice(alias.specifierPrefix.length));
        const resolved = resolveToFile(raw);
        if (!resolved) {
          throw new Error(`Could not resolve aliased import "${specifier}" from ${current}`);
        }
        queue.push(resolved);
        continue;
      }
      // Bare specifier — an npm package. Not walked into node_modules; its
      // presence here is exactly what the @preact/signals assertion checks.
      externalSpecifiers.add(specifier);
    }
  }

  return { visitedFiles, externalSpecifiers };
}

describe('pairing sketch — module graph (task 0, barrel-leak guard)', () => {
  let graph: GraphResult;

  beforeAll(() => {
    graph = walkModuleGraph(ENTRY, loadAliasRules());
  });

  it('sanity: the walk actually traverses the real graph (positive control)', () => {
    // If this were vacuous (e.g. a resolution bug that silently visits only
    // the entry file), the leak assertions below would pass for the wrong
    // reason. Pin a lower bound and name files that MUST be on this graph.
    expect(graph.visitedFiles.size).toBeGreaterThan(20);
    const relative = [...graph.visitedFiles].map(f => f.slice(ROOT.length + 1));
    expect(relative).toContain('src/features/compare/batch/pairing/sketch.ts');
    expect(relative).toContain('src/features/defined-terms/analyse.ts');
    expect(relative).toContain('src/features/defined-terms/definitions.ts');
    expect(relative).toContain('src/features/party-role/analyse.ts');
    expect(relative).toContain('src/features/party-role/roles.ts');
  });

  it('does not reach @shared/findings (the store module, specifically)', () => {
    const findingsFiles = [...graph.visitedFiles]
      .map(f => f.slice(ROOT.length + 1))
      .filter(f => f.startsWith('src/shared/findings/'));
    expect(findingsFiles).toEqual([]);
  });

  it('does not reach @preact/signals', () => {
    const signalsRefs = [...graph.externalSpecifiers].filter(
      s => s === '@preact/signals' || s.startsWith('@preact/signals/'),
    );
    expect(signalsRefs).toEqual([]);
  });
});
