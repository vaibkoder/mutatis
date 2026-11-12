/**
 * The web target's visual seam (ADR-0044).
 *
 * Two things this file guards, both of which HAVE been wrong in this tree and
 * neither of which any rendering test can see:
 *
 *  1. **The stylesheet seam.** `web.css` re-values the --dp-* tokens for the
 *     browser. That is only safe because the task pane never loads it and this
 *     file never loads the Office theme bridge. Wire the two together and the
 *     add-in silently inherits a palette built for a different host — a change
 *     no unit test asserting markup would notice.
 *
 *  2. **Fallback colours.** shell.css is scrupulously token-only, but it
 *     references nine tokens that tokens.css never defines, each with a
 *     hardcoded Fluent-era fallback baked into the `var()` — or, for two of
 *     them, with NO fallback at all, which makes the whole declaration invalid
 *     at computed-value time (a transparent background, a zero padding). Those
 *     defaults are invisible while they happen to match the palette and become
 *     foreign the moment it moves. The test asserts the web target defines
 *     every one of them, so no web surface can fall back to a colour nobody
 *     chose.
 *
 * Deliberately NOT a snapshot of the values: a snapshot would fail on every
 * intentional design change and prove nothing about the seam.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(__dirname, '../../..');
const read = (p: string) => readFileSync(resolve(root, p), 'utf8');

const webCss = read('src/styles/web.css');
const shellCss = read('src/styles/shell.css');
const tokensCss = read('src/styles/tokens.css');
const officeThemeCss = read('src/styles/office-theme.css');
const webHtml = read('src/web/index.html');
const taskpaneHtml = read('src/taskpane/index.html');
const webMain = read('src/web/main.tsx');
const taskpaneMain = read('src/taskpane/main.tsx');

/** Every `--dp-*` custom property DECLARED at the start of a line. */
function declaredTokens(css: string): Set<string> {
  return new Set(Array.from(css.matchAll(/^\s*(--dp-[a-z0-9-]+)\s*:/gm), m => m[1]!));
}

/**
 * The declarations inside ONE rule, found by brace-matching from its selector.
 *
 * Needed because "declared somewhere in web.css" is not the invariant that
 * matters and the first version of this file got that wrong: a token declared
 * only in the two DARK blocks satisfies a whole-file scan while light mode
 * still falls back to whatever literal is baked into the `var()`. The base
 * `:root` block is the one that must always resolve.
 */
function ruleBlock(css: string, selector: string): string {
  const at = css.indexOf(selector + ' {');
  if (at === -1) throw new Error(`selector not found: ${selector}`);
  let depth = 0;
  for (let i = at; i < css.length; i++) {
    if (css[i] === '{') depth++;
    else if (css[i] === '}' && --depth === 0) return css.slice(at, i);
  }
  throw new Error(`unbalanced braces after ${selector}`);
}

/** Tokens declared in web.css's BASE light :root — the always-resolving set. */
function baseTokens(): Set<string> {
  // indexOf finds ':root {' before ':root:not(' / ':root[' because the base
  // block is written first and the others never match the exact string.
  return declaredTokens(ruleBlock(webCss, ':root'));
}

/** Every `var(--dp-x, <fallback>)` where the fallback is not itself a var(). */
function tokensWithLiteralFallback(css: string): Set<string> {
  const out = new Set<string>();
  for (const m of css.matchAll(/var\(\s*(--dp-[a-z0-9-]+)\s*,\s*([^)]*)\)/g)) {
    if (!m[2]!.includes('var(')) out.add(m[1]!);
  }
  return out;
}

/** Every `--dp-*` referenced with NO fallback — an undefined one is invalid CSS. */
function tokensWithNoFallback(css: string): Set<string> {
  return new Set(Array.from(css.matchAll(/var\(\s*(--dp-[a-z0-9-]+)\s*\)/g), m => m[1]!));
}

describe('web target — stylesheet seam', () => {
  it('the task pane never loads web.css, and the web app never loads the Office bridge', () => {
    // The whole re-skin rests on this. If either side gains the other's
    // stylesheet the token values cross over and one host gets the other's look.
    expect(taskpaneHtml).not.toContain('web.css');
    expect(taskpaneMain).not.toContain('web.css');
    expect(webHtml).not.toContain('office-theme.css');
    expect(webMain).not.toContain('office-theme.css');
  });

  it('web.css is loaded by the web entry points, and loaded last', () => {
    expect(webMain).toContain('web.css');
    // Order matters: web.css re-values tokens.css's :root at equal specificity,
    // so it only wins by coming later in the cascade.
    const tokensAt = webHtml.indexOf('tokens.css');
    const shellAt = webHtml.indexOf('shell.css');
    const webAt = webHtml.indexOf('web.css');
    expect(tokensAt).toBeGreaterThan(-1);
    expect(webAt).toBeGreaterThan(tokensAt);
    expect(webAt).toBeGreaterThan(shellAt);
  });

  it('the shared stylesheets carry no web-only palette', () => {
    // The re-skin is oklch throughout. Its appearance in a shared file means it
    // leaked out of the web target and into the add-in.
    for (const [name, css] of [
      ['tokens.css', tokensCss],
      ['office-theme.css', officeThemeCss],
      ['shell.css', shellCss],
    ] as const) {
      expect(css, `${name} must not carry the web palette`).not.toContain('oklch(');
    }
  });
});

describe('web target — no surface falls back to an unchosen colour', () => {
  // The BASE :root block, deliberately — see ruleBlock().
  const webTokens = baseTokens();

  it('defines every token that shell.css or web.css references with a literal fallback', () => {
    const needed = new Set([
      ...tokensWithLiteralFallback(shellCss),
      ...tokensWithLiteralFallback(webCss),
    ]);
    // Spacing/typography fallbacks are dimensions, not colours, and tokens.css
    // already defines them; only the ones tokens.css misses matter here.
    const tokensDefined = declaredTokens(tokensCss);
    const missing = [...needed].filter(t => !tokensDefined.has(t) && !webTokens.has(t));
    expect(missing, `web.css must define: ${missing.join(', ')}`).toEqual([]);
  });

  it('defines every token referenced with no fallback at all', () => {
    // `background: var(--dp-undefined)` is invalid at computed-value time — the
    // property falls back to its INITIAL value, not to anything sensible. Two
    // of these (--dp-bg-elevated, --dp-space-md) were live in shell.css with
    // no definition anywhere, painting a transparent background and a zero
    // padding on any surface that used them.
    const needed = new Set([
      ...tokensWithNoFallback(shellCss),
      ...tokensWithNoFallback(webCss),
    ]);
    const tokensDefined = declaredTokens(tokensCss);
    const missing = [...needed].filter(t => !tokensDefined.has(t) && !webTokens.has(t));
    expect(missing, `web.css must define: ${missing.join(', ')}`).toEqual([]);
  });

  it('renames nothing — every token tokens.css declares still exists for the web', () => {
    // shell.css is ~3500 lines of consumers and ADR-0012 forbids a component
    // hardcoding a colour, so the web target may only re-VALUE tokens.
    // A token that tokens.css declares keeps its light value unless web.css
    // deliberately overrides it; the failure this catches is a *rename*.
    const renamed = [...webTokens].filter(
      t => !declaredTokens(tokensCss).has(t) && !shellCss.includes(`var(${t}`) && !webCss.includes(`var(${t}`),
    );
    expect(renamed, `web.css declares tokens nothing consumes: ${renamed.join(', ')}`).toEqual([]);
  });
});

describe('web target — theme control (ADR-0044)', () => {
  it('the two dark blocks declare exactly the same tokens', () => {
    // They are duplicated on purpose (one is the pre-paint @media block, one
    // is the explicit override), which means they can silently drift: a token
    // added to one and not the other gives a user who PICKED dark a different
    // palette from a user whose OS is dark, and no screenshot of either alone
    // would show it.
    const media = declaredTokens(ruleBlock(webCss, ':root:not([data-theme])'));
    const explicit = declaredTokens(ruleBlock(webCss, ":root[data-theme='dark']"));
    expect([...media].sort()).toEqual([...explicit].sort());
    expect(media.size).toBeGreaterThan(10);
  });

  it('every dark override has a light counterpart in the base block', () => {
    // A dark-only token resolves to nothing in light mode.
    const base = baseTokens();
    const darkOnly = [...declaredTokens(ruleBlock(webCss, ":root[data-theme='dark']"))]
      .filter(t => t !== 'color-scheme' && !base.has(t));
    expect(darkOnly, `declared for dark but not for light: ${darkOnly.join(', ')}`).toEqual([]);
  });

  it('the pre-paint dark block is scoped away from an explicit choice', () => {
    // This is the whole reason "follow system" removes the attribute rather
    // than writing the resolved value. If the scope is dropped, an explicit
    // light choice loses to a dark OS.
    expect(webCss).toMatch(
      /@media \(prefers-color-scheme: dark\) \{\s*:root:not\(\[data-theme\]\)/,
    );
  });

  it('an explicit dark choice is honoured outside any media query', () => {
    // Without this, a user on a light OS who picks dark gets nothing.
    const explicit = webCss.indexOf(":root[data-theme='dark']");
    expect(explicit).toBeGreaterThan(-1);
    // ...and it must sit at the top level, not nested inside the @media block.
    const mediaStart = webCss.indexOf('@media (prefers-color-scheme: dark)');
    const mediaEnd = webCss.indexOf('\n}', webCss.indexOf(':root:not([data-theme])'));
    expect(explicit).toBeGreaterThan(mediaEnd);
    expect(mediaStart).toBeLessThan(mediaEnd);
  });

  it('keeps the forced-colours block, and lets it win', () => {
    const forced = webCss.indexOf('@media (forced-colors: active)');
    expect(forced).toBeGreaterThan(-1);
    // Equal specificity, so it only overrides the theme blocks by coming last.
    expect(forced).toBeGreaterThan(webCss.indexOf(":root[data-theme='dark']"));
  });

  it('focus stays visible and uses the focus token', () => {
    expect(webCss).toContain('--dp-focus-ring: 2px solid var(--dp-focus)');
    expect(webCss).toMatch(/:focus-visible\s*\{[^}]*outline:\s*var\(--dp-focus-ring\)/);
  });
});
