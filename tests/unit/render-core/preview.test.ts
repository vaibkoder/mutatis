import { describe, it, expect, beforeAll } from 'vitest';
import {
  renderToLayout, renderPreview, createFontResolver, createHarfBuzzShaper, DEFAULT_THEME, PAGE_INK,
  type Shaper, type LayoutResult,
} from '@shared/render-core/index.js';
import { testFaces } from './_fixtures.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const fonts = createFontResolver(testFaces());
let shaper: Shaper;
beforeAll(async () => { shaper = await createHarfBuzzShaper(); });

function multiPageLayout(): LayoutResult {
  const paras: string[] = [];
  for (let i = 0; i < 60; i++) {
    paras.push(`<w:p><w:r><w:t xml:space="preserve">Clause ${i}: the parties agree this Agreement is governed by the laws of Singapore. </w:t></w:r></w:p>`);
  }
  return renderToLayout({ authoredBodyXml: `<w:body ${NS}>${paras.join('')}</w:body>` }, fonts, shaper);
}

/** Parse a `"12.5pt"` inline style back to a number. */
const ptOf = (s: string): number => parseFloat(s.replace('pt', ''));

/**
 * Relative luminance of a `#rrggbb` literal. Arithmetic on a constant we own —
 * deliberately NOT a colour parser and NOT a CSS resolver; see the docblock below
 * for why this file no longer has one of those.
 */
function luminance(hex: string): number {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  if (!m) throw new Error(`PAGE_INK is not a #rrggbb literal: ${hex}`);
  const [r, g, b] = [1, 2, 3].map((i) => parseInt(m[i] as string, 16) / 255) as [number, number, number];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

describe('Preview painter (Output A)', () => {
  it('renders one page frame per layout page and paints line content', () => {
    const layout = multiPageLayout();
    expect(layout.pages.length).toBeGreaterThan(1);
    const { root } = renderPreview(layout);

    const frames = root.querySelectorAll('.rc-page');
    expect(frames.length).toBe(layout.pages.length);
    // jsdom reports zero height → every page is painted.
    const lines = root.querySelectorAll('.rc-line');
    const totalLayoutLines = layout.pages.reduce((a, p) => a + p.bodyLines.length, 0);
    expect(lines.length).toBe(totalLayoutLines);
  });

  it('positions line + run boxes to match the layout result (pt-for-pt)', () => {
    const layout = multiPageLayout();
    const { root } = renderPreview(layout);
    const page0 = layout.pages[0]!;
    const g = page0.geometry;
    const firstLine = page0.bodyLines[0]!;

    const lineEl = root.querySelector('.rc-page .rc-line') as HTMLElement;
    expect(ptOf(lineEl.style.left)).toBeCloseTo(g.marginLeftPt, 3);
    expect(ptOf(lineEl.style.top)).toBeCloseTo(g.marginTopPt + firstLine.topPt, 3);

    const runEl = lineEl.querySelector('.rc-run') as HTMLElement;
    expect(ptOf(runEl.style.left)).toBeCloseTo(firstLine.runs[0]!.xPt, 3);
  });

  it('applies the theme via CSS variables and decorates changed runs (never filters)', () => {
    const inner = '<w:p><w:r><w:t xml:space="preserve">base </w:t></w:r>' +
      '<w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins>' +
      '<w:del w:id="2"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const { root } = renderPreview(layout, { theme: DEFAULT_THEME });

    expect(root.style.getPropertyValue('--rc-ins')).toBe(DEFAULT_THEME.insertColor);
    expect(root.style.getPropertyValue('--rc-del')).toBe(DEFAULT_THEME.deleteColor);

    const runs = [...root.querySelectorAll('.rc-run')] as HTMLElement[];
    const added = runs.find((r) => r.textContent === 'added')!;
    const gone = runs.find((r) => r.textContent?.startsWith('gone'))!;
    // Tracked-change runs take their decoration colour; underline/strike are now
    // continuous `.rc-deco` rules (one per merged span), not per-span text-decoration.
    expect(added.style.color).toBeTruthy();
    expect(gone.style.color).toBeTruthy();
    const decos = [...root.querySelectorAll('.rc-deco')] as HTMLElement[];
    expect(decos.length).toBeGreaterThanOrEqual(2); // one underline (added) + one strike (gone)
    // insertion highlight realised as a rect div
    expect(root.querySelectorAll('.rc-rect').length).toBeGreaterThan(0);
    // changed line carries a change bar
    expect(root.querySelectorAll('.rc-changebar').length).toBeGreaterThan(0);
  });

  /**
   * The page frame is paper: white background, paper ink. Its text colour must
   * come from the painter, never from the host, because a dark task pane sets a
   * near-white `--dp-fg` that would inherit straight onto the white page — the
   * DECORATED runs carry their own colour and stay readable, so the failure is
   * invisible to every assertion above: only the UNCHANGED body text vanishes.
   *
   * Two things this test learned the hard way, both by a reviewer constructing the
   * case the previous version forbade:
   *
   *  - It does NOT try to resolve the colour. jsdom hands back the unresolved
   *    declaration for a `var()`, and an earlier version answered that with a
   *    hand-rolled `var()` resolver — which let `color: var(--dp-fg, ...)`, the
   *    defect itself pointed straight at the host foreground, pass green. Assert
   *    the declaration text and the pinned value; resolve nothing.
   *  - It samples the UNDECORATED RUN, not the `.rc-page` frame. `paintRun` sets
   *    an inline colour only for `insUnderline`/`delStrike`, so plain body text is
   *    defended by inheritance alone — and any rule matching `.rc-run` beats
   *    inheritance while the frame's own declaration stays correct. Sampling the
   *    frame made that mutation green.
   *
   * What it still cannot see: whether the ink is LEGIBLE against the paper beyond
   * the two constants checked below, and any defect that moves every site at once.
   */
  it('paints undecorated text in paper ink, not the host theme foreground', () => {
    const inner = '<w:p><w:r><w:t xml:space="preserve">base </w:t></w:r>' +
      '<w:ins w:id="1"><w:r><w:t>added</w:t></w:r></w:ins></w:p>';
    const layout = renderToLayout({ authoredBodyXml: `<w:body ${NS}>${inner}</w:body>` }, fonts, shaper, { theme: DEFAULT_THEME });
    const { root } = renderPreview(layout, { theme: DEFAULT_THEME });

    // Host is a dark task pane: near-white foreground on the container.
    const host = document.createElement('div');
    host.style.color = 'rgb(243, 242, 241)';
    host.appendChild(root);
    document.body.appendChild(host);
    try {
      const frame = root.querySelector('.rc-page') as HTMLElement;
      // The run carrying "base" — undecorated, so `paintRun` gave it no inline
      // colour and it is the text the reader loses. Sampling `.rc-page` instead
      // misses any rule that recolours the runs themselves.
      const plain = [...root.querySelectorAll<HTMLElement>('.rc-run')]
        .find((r) => r.style.color === '');
      expect(plain, 'no undecorated run to sample').toBeDefined();
      const declared = getComputedStyle(plain!).color.replace(/\s+/g, ' ').trim();

      // (a) Declared somewhere above it, so nothing is inherited from the host:
      //     drop the `color` rule from `.rc-page` and the host's near-white wins.
      expect(declared).not.toBe('rgb(243, 242, 241)');
      // (b) And declared from the PAPER ink, with paper ink as its fallback. Name
      //     the host's variable instead (`var(--dp-fg, ...)`) and a dark pane
      //     paints near-white on white paper. The fallback half bites in the one
      //     case where the variable is never set — `setThemeVars` did not run —
      //     NOT when a host defines `--rc-ink`, which is when a fallback is by
      //     definition not consulted. This pins declaration TEXT, so reformatting
      //     it (dropping the space, adding `!important`) is a deliberate act.
      expect(declared).toBe(`var(--rc-ink, ${PAGE_INK})`);
      // (c) And the painter pins that variable to paper ink. This is what decides
      //     the rendered colour whenever anything ELSE in the cascade defines
      //     `--rc-ink` — custom properties inherit, so an ancestor that sets it
      //     would otherwise reach the page. Exact match, because a colour that
      //     merely parses to the same RGB can still be transparent or invalid.
      expect(root.style.getPropertyValue('--rc-ink')).toBe(PAGE_INK);
      // (d) And the two constants that (a)–(c) only pin into AGREEMENT are a
      //     readable pair. Without this, retyping PAGE_INK as the host's near-
      //     white moves the stylesheet, the painter and this test's own oracle
      //     together, and everything above stays green on invisible body text.
      expect(luminance(PAGE_INK)).toBeLessThan(0.2);
      expect(getComputedStyle(frame).background).toContain('var(--rc-page, #fff)');
    } finally {
      host.remove();
    }
  });
});
