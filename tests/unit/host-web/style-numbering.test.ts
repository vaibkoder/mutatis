/**
 * Shared numbering — the numbering a paragraph inherits from its style.
 * ADR-0042 · ECMA-376 §17.3.1.19.
 *
 * A firm template puts `<w:numPr>` on the STYLE and leaves paragraphs carrying
 * nothing but `<w:pStyle>`. Both readers that needed a paragraph's list level —
 * Compare's `ParaRecord` and RenderCore's `parse-body` — read only the direct
 * `<w:pPr><w:numPr>`, so on such a document every clause looked unnumbered.
 */
import { describe, it, expect } from 'vitest';

import {
  parseStyleNumberingSource,
  resolveParaNumPr,
  readNumPr,
} from '@shared/numbering/index.js';
import { parseStyles } from '@shared/render-core/input/styles.js';
import { parsePreserveOrder, tagOf, childrenOf } from '@shared/ooxml/po-xml.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function styles(inner: string): string {
  return `<w:styles ${W}>${inner}</w:styles>`;
}

/** A `<w:style>` with an optional basedOn and an optional numPr. */
function style(
  id: string,
  opts: { basedOn?: string; numId?: string; ilvl?: number; isDefault?: boolean } = {},
): string {
  const numPr =
    opts.numId === undefined
      ? ''
      : `<w:numPr><w:ilvl w:val="${opts.ilvl ?? 0}"/><w:numId w:val="${opts.numId}"/></w:numPr>`;
  const basedOn = opts.basedOn ? `<w:basedOn w:val="${opts.basedOn}"/>` : '';
  const dflt = opts.isDefault ? ' w:default="1"' : '';
  return `<w:style w:type="paragraph"${dflt} w:styleId="${id}">${basedOn}<w:pPr>${numPr}</w:pPr></w:style>`;
}

/** The `<w:pPr>` children of a one-paragraph body. */
function pPrOf(xml: string) {
  const p = parsePreserveOrder(xml).find((n) => tagOf(n) === 'w:p');
  if (!p) return [];
  for (const kid of childrenOf(p)) if (tagOf(kid) === 'w:pPr') return childrenOf(kid);
  return [];
}

describe('style-chain numbering resolution', () => {
  it('resolves a paragraph with NO direct numPr through its pStyle', () => {
    const source = parseStyleNumberingSource(styles(style('Sub2', { numId: '3', ilvl: 2 })));
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Sub2"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Sub2', pPr)).toEqual({ numId: '3', ilvl: 2 });
  });

  it('walks the basedOn ancestry to the nearest definition', () => {
    const source = parseStyleNumberingSource(
      styles(
        style('Root', { numId: '7', ilvl: 0 }) +
          style('Middle', { basedOn: 'Root' }) +
          style('Leaf', { basedOn: 'Middle' }),
      ),
    );
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Leaf"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Leaf', pPr)).toEqual({ numId: '7', ilvl: 0 });
  });

  it('prefers the NEAREST ancestor when more than one in the chain numbers', () => {
    const source = parseStyleNumberingSource(
      styles(
        style('Root', { numId: '7', ilvl: 0 }) + style('Leaf', { basedOn: 'Root', numId: '9', ilvl: 4 }),
      ),
    );
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Leaf"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Leaf', pPr)).toEqual({ numId: '9', ilvl: 4 });
  });

  it("a direct numPr overrides the style's", () => {
    const source = parseStyleNumberingSource(styles(style('Sub2', { numId: '3', ilvl: 2 })));
    const pPr = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Sub2"/>` +
        `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Sub2', pPr)).toEqual({ numId: '11', ilvl: 0 });
  });

  it('a direct ilvl with no numId re-levels the INHERITED instance (the indent button)', () => {
    const source = parseStyleNumberingSource(styles(style('Sub3', { numId: '3', ilvl: 3 })));
    const pPr = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Sub3"/>` +
        `<w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Sub3', pPr)).toEqual({ numId: '3', ilvl: 1 });
  });

  // ── numId and ilvl are two properties, not one ────────────────────────────
  // Found in review. `<w:numPr>` was treated as
  // atomic whenever the DECLARATION came from a style, and the two directions of
  // that asymmetry are the next three tests. Both were unguarded: the only
  // direct-override test above supplies BOTH children, so it cannot express
  // either case, and the whole suite stayed green with the defect in place.

  it('a STYLE that declares only <w:ilvl> re-levels the instance it inherits', () => {
    // `basedOn` a numbered base style, overriding nothing but the level, is
    // ordinary template construction. Skipping such a style discarded the level
    // it declares: `Sub` resolved to 0, and a re-levelling the reader sees as
    // `(a) (b)` → `(i) (ii)` produced NO op at all (ADR-0008 suppression).
    const source = parseStyleNumberingSource(
      styles(
        style('Body', { numId: '3', ilvl: 0 }) +
          `<w:style w:type="paragraph" w:styleId="Sub"><w:basedOn w:val="Body"/>` +
          `<w:pPr><w:numPr><w:ilvl w:val="2"/></w:numPr></w:pPr></w:style>`,
      ),
    );
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Sub"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Sub', pPr)).toEqual({ numId: '3', ilvl: 2 });
  });

  it("a direct numId with no ilvl keeps the level the STYLE declared", () => {
    // The mirror. Word writes a bare `<w:numId>` when a paragraph is attached to
    // a list without changing its level; reading that as level 0 invented a
    // `formatting` change between two documents that differ only in whether a
    // redundant `<w:ilvl>` is spelled out.
    const source = parseStyleNumberingSource(styles(style('Clause2', { numId: '3', ilvl: 1 })));
    const pPr = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Clause2"/>` +
        `<w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Clause2', pPr)).toEqual({ numId: '3', ilvl: 1 });
  });

  it('falls back to level 0 only when NOTHING in the chain declares a level', () => {
    // The `?? 0` itself, which is the remaining default and must stay one:
    // changing it to any other integer has to fail here.
    const source = parseStyleNumberingSource(styles(style('Plain')));
    const pPr = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Plain"/>` +
        `<w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Plain', pPr)).toEqual({ numId: '5', ilvl: 0 });
  });

  it('honours numId="0" as "not numbered", not as instance zero', () => {
    const source = parseStyleNumberingSource(styles(style('Sub2', { numId: '3', ilvl: 2 })));
    const pPr = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Sub2"/>` +
        `<w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Sub2', pPr)).toBeNull();
  });

  it('honours numId="0" on a STYLE too — a derived style cancelling its parent', () => {
    // The direct-`numPr` half above was guarded; the style-chain half was not,
    // and this module exists for the style chain. Word's documented way to say
    // "this style is not numbered" when the style it is based on IS numbered is
    // `<w:numId w:val="0"/>` on the derived style, and a reader that treats it
    // as instance zero renders a number the document does not have.
    const source = parseStyleNumberingSource(
      styles(style('Base', { numId: '3', ilvl: 1 }) + style('Unnumbered', { basedOn: 'Base', numId: '0' })),
    );
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Unnumbered"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Unnumbered', pPr)).toBeNull();
    // And the parent is unaffected — the cancel is the derived style's, not a
    // property of the chain.
    expect(resolveParaNumPr(source, 'Base', [])).toEqual({ numId: '3', ilvl: 1 });
  });

  it('an unstyled paragraph inherits from the DEFAULT paragraph style', () => {
    const source = parseStyleNumberingSource(
      styles(style('Normal', { numId: '5', ilvl: 1, isDefault: true })),
    );

    expect(resolveParaNumPr(source, null, [])).toEqual({ numId: '5', ilvl: 1 });
  });

  it('returns null when nothing in the chain numbers the paragraph', () => {
    const source = parseStyleNumberingSource(styles(style('Plain')));
    const pPr = pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="Plain"/></w:pPr></w:p>`);

    expect(resolveParaNumPr(source, 'Plain', pPr)).toBeNull();
  });

  it('survives a basedOn cycle instead of looping', () => {
    const source = parseStyleNumberingSource(
      styles(style('A', { basedOn: 'B' }) + style('B', { basedOn: 'A' })),
    );

    expect(resolveParaNumPr(source, 'A', [])).toBeNull();
  });

  it('an absent styles.xml resolves everything to un-numbered', () => {
    const source = parseStyleNumberingSource(null);
    expect(source.styles.size).toBe(0);
    expect(resolveParaNumPr(source, 'Anything', [])).toBeNull();
  });

  it('ignores a <w:ilvl> with no w:val instead of re-levelling the paragraph to 0', () => {
    // `Number('')` is 0 and finite, so a valueless <w:ilvl/> used to OVERRIDE the
    // style's level with 0 — while `w:val="abc"`, no more malformed, was
    // correctly ignored. Two equally broken inputs, handled opposite ways, and
    // the empty one in the dangerous direction.
    const source = parseStyleNumberingSource(styles(style('Sub3', { numId: '3', ilvl: 3 })));
    const bare = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Sub3"/><w:numPr><w:ilvl/></w:numPr></w:pPr></w:p>`,
    );
    const junk = pPrOf(
      `<w:p ${W}><w:pPr><w:pStyle w:val="Sub3"/><w:numPr><w:ilvl w:val="abc"/></w:numPr></w:pPr></w:p>`,
    );

    expect(resolveParaNumPr(source, 'Sub3', bare)).toEqual({ numId: '3', ilvl: 3 });
    expect(resolveParaNumPr(source, 'Sub3', junk)).toEqual({ numId: '3', ilvl: 3 });
  });

  it('reads w:default as an ECMA-376 boolean, so w:default="0" is NOT the default', () => {
    const xml = styles(
      '<w:style w:type="paragraph" w:default="0" w:styleId="NotDefault"><w:pPr>' +
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr></w:style>',
    );
    const source = parseStyleNumberingSource(xml);

    expect(source.defaultParaStyleId).toBeNull();
    // An unstyled paragraph must not inherit numbering from a style that
    // explicitly declares itself NOT the default.
    expect(resolveParaNumPr(source, null, [])).toBeNull();
  });

  it("render-core's own style table satisfies the interface — the claim 08c depends on", () => {
    // `resolveParaNumPr` takes a STRUCTURAL interface so RenderCore can pass the
    // `StyleTable` it already builds, instead of Compare taking a dependency on
    // RenderCore's input layer or a second parser existing for nothing. ADR-0042
    // and the 08b plan both say wiring RenderCore later is "an argument, not a
    // rewrite". That was asserted only in prose; this is the assertion.
    //
    // **Deliberately a RUNTIME matrix, not a type annotation.** `tsconfig.json`
    // excludes `tests/` and Vitest's esbuild strips types without checking them,
    // so `const x: StyleNumberingSource = table` inside this file is evaluated by
    // nothing at all — it would go on passing after the shapes diverged. Every
    // assertion below executes.
    const xml = styles(
      style('Root', { numId: '7', ilvl: 0 }) +
        style('Mid', { basedOn: 'Root' }) +
        style('Leaf', { basedOn: 'Mid', numId: '3', ilvl: 2 }) +
        style('Plain') +
        style('Norm', { numId: '5', ilvl: 1, isDefault: true }),
    );
    const table = parseStyles(xml);
    const own = parseStyleNumberingSource(xml);

    // The shape the interface actually requires, checked on the real object.
    expect(typeof table.defaultParaStyleId === 'string' || table.defaultParaStyleId === null).toBe(true);
    for (const id of ['Root', 'Mid', 'Leaf', 'Plain', 'Norm']) {
      const entry = table.styles.get(id);
      expect(entry, `render-core's table lost style ${id}`).toBeDefined();
      expect(entry!.basedOn === null || typeof entry!.basedOn === 'string').toBe(true);
      expect(Array.isArray(entry!.pPr)).toBe(true);
    }

    // ...and the same ANSWER from either input, across every resolution path:
    // inherited, walked through basedOn, nearest-wins, directly overridden,
    // unnumbered, and the default style an unstyled paragraph falls back to.
    const cases: Array<[string | null, string]> = [
      ['Leaf', `<w:p ${W}><w:pPr><w:pStyle w:val="Leaf"/></w:pPr></w:p>`],
      ['Mid', `<w:p ${W}><w:pPr><w:pStyle w:val="Mid"/></w:pPr></w:p>`],
      ['Plain', `<w:p ${W}><w:pPr><w:pStyle w:val="Plain"/></w:pPr></w:p>`],
      [null, `<w:p ${W}><w:pPr/></w:p>`],
      [
        'Leaf',
        `<w:p ${W}><w:pPr><w:pStyle w:val="Leaf"/>` +
          `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr></w:pPr></w:p>`,
      ],
    ];
    for (const [styleId, paraXml] of cases) {
      const pPr = pPrOf(paraXml);
      expect(resolveParaNumPr(table, styleId, pPr), `styleId ${styleId ?? '(none)'}`).toEqual(
        resolveParaNumPr(own, styleId, pPr),
      );
    }
    // Not vacuously equal — the matrix really does resolve numbering.
    expect(resolveParaNumPr(table, 'Leaf', pPrOf(cases[0]![1]))).toEqual({ numId: '3', ilvl: 2 });
    expect(resolveParaNumPr(table, 'Mid', pPrOf(cases[1]![1]))).toEqual({ numId: '7', ilvl: 0 });
    expect(resolveParaNumPr(table, null, pPrOf(cases[3]![1]))).toEqual({ numId: '5', ilvl: 1 });
  });

  it('readNumPr distinguishes "no numPr" from "a numPr with no numId"', () => {
    expect(readNumPr(pPrOf(`<w:p ${W}><w:pPr><w:pStyle w:val="X"/></w:pPr></w:p>`))).toBeNull();
    expect(
      readNumPr(pPrOf(`<w:p ${W}><w:pPr><w:numPr><w:ilvl w:val="2"/></w:numPr></w:pPr></w:p>`)),
    ).toEqual({ ilvl: 2 });
  });
});
