/**
 * host-web — numbering.xml IR + parser tests.
 *
 * MUST pass now: feed `numbering.xml` strings, assert the parsed IR, and assert
 * `resolveLevel` applies `<w:lvlOverride>` (startOverride + inline lvl) on top of
 * the abstractNum level. Pure parse/lookup — no rendered numbers here.
 */
import { describe, it, expect } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';
import {
  parseNumberingScheme,
  resolveLevel,
  type NumberingScheme,
} from '@host-web/numbering-scheme.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function schemeFrom(numberingXml: string): NumberingScheme {
  const b = new PizZipBundle();
  b.setPart('word/numbering.xml', numberingXml);
  return parseNumberingScheme(b);
}

// abstractNum 0: a 2-level decimal/lowerLetter scheme; num 1 → abstract 0.
const BASIC =
  `<?xml version="1.0"?><w:numbering ${W}>` +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:lvlRestart w:val="0"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

describe('parseNumberingScheme — basic IR', () => {
  it('parses abstractNum levels with format, lvlText, start, lvlRestart', () => {
    const scheme = schemeFrom(BASIC);
    const abs = scheme.abstractNums.get('0');
    expect(abs).toBeDefined();
    expect(abs!.levels.size).toBe(2);

    const l0 = abs!.levels.get(0)!;
    expect(l0).toEqual({ ilvl: 0, numFmt: 'decimal', lvlText: '%1.', start: 1, isLgl: false });

    const l1 = abs!.levels.get(1)!;
    expect(l1).toEqual({
      ilvl: 1,
      numFmt: 'lowerLetter',
      lvlText: '(%2)',
      start: 1,
      lvlRestart: 0,
      isLgl: false,
    });
  });

  it('parses num instances pointing at their abstractNum', () => {
    const scheme = schemeFrom(BASIC);
    const num = scheme.nums.get('1');
    expect(num).toBeDefined();
    expect(num!.abstractNumId).toBe('0');
    expect(num!.lvlOverrides.size).toBe(0);
  });

  it('defaults start to 1 when <w:start> is absent', () => {
    const scheme = schemeFrom(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="9">` +
        '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
        '</w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="9"/></w:num></w:numbering>',
    );
    expect(scheme.abstractNums.get('9')!.levels.get(0)!.start).toBe(1);
  });

  it('captures <w:isLgl/> when present', () => {
    const scheme = schemeFrom(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="2">` +
        '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:isLgl/></w:lvl>' +
        '</w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="2"/></w:num></w:numbering>',
    );
    expect(scheme.abstractNums.get('2')!.levels.get(0)!.isLgl).toBe(true);
  });

  it('returns an empty scheme when numbering.xml is absent', () => {
    const scheme = parseNumberingScheme(new PizZipBundle());
    expect(scheme.nums.size).toBe(0);
    expect(scheme.abstractNums.size).toBe(0);
  });
});

describe('parseNumberingScheme — lvlOverride', () => {
  // num 2 → abstract 0, overriding level 0's start to 5 and replacing level 1 inline.
  const WITH_OVERRIDE =
    `<w:numbering ${W}>` +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="0"/>' +
    '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>' +
    '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%2)"/></w:lvl></w:lvlOverride>' +
    '</w:num>' +
    '</w:numbering>';

  it('captures startOverride and inline-lvl overrides per level', () => {
    const scheme = schemeFrom(WITH_OVERRIDE);
    const num = scheme.nums.get('2')!;
    expect(num.lvlOverrides.get(0)).toEqual({ startOverride: 5 });
    const ov1 = num.lvlOverrides.get(1)!;
    expect(ov1.startOverride).toBeUndefined();
    expect(ov1.lvl).toEqual({
      ilvl: 1,
      numFmt: 'upperRoman',
      lvlText: '%2)',
      start: 1,
      isLgl: false,
    });
  });
});

describe('resolveLevel', () => {
  it('returns the abstractNum level when no override applies', () => {
    const scheme = schemeFrom(BASIC);
    expect(resolveLevel(scheme, '1', 0)).toEqual({
      ilvl: 0,
      numFmt: 'decimal',
      lvlText: '%1.',
      start: 1,
      isLgl: false,
    });
  });

  it('applies startOverride on top of the abstract level', () => {
    const scheme = schemeFrom(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">` +
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
        '</w:abstractNum>' +
        '<w:num w:numId="2"><w:abstractNumId w:val="0"/>' +
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>',
    );
    expect(resolveLevel(scheme, '2', 0)).toEqual({
      ilvl: 0,
      numFmt: 'decimal',
      lvlText: '%1.',
      start: 5, // overridden
      isLgl: false,
    });
  });

  it('returns the inline-lvl replacement when present', () => {
    const scheme = schemeFrom(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">` +
        '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
        '</w:abstractNum>' +
        '<w:num w:numId="2"><w:abstractNumId w:val="0"/>' +
        '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride>' +
        '</w:num></w:numbering>',
    );
    expect(resolveLevel(scheme, '2', 0)?.numFmt).toBe('upperLetter');
    expect(resolveLevel(scheme, '2', 0)?.lvlText).toBe('%1)');
  });

  it('returns undefined for an unknown numId or level', () => {
    const scheme = schemeFrom(BASIC);
    expect(resolveLevel(scheme, '999', 0)).toBeUndefined();
    expect(resolveLevel(scheme, '1', 7)).toBeUndefined();
  });
});

/**
 * `<w:numStyleLink>` — a list built through a Word LIST STYLE.
 *
 * Found in review. , as the THIRD
 * ADR-0008 suppression of that plan: the `<w:abstractNum>` a `<w:num>` points at
 * holds nothing but a delegation, so every level resolved to `undefined`, every
 * paragraph rendered no label, and Compare reported NOTHING for a redefinition
 * that moves every number on the page (`compare.diff.property.test.ts`,
 * *"a list built through a Word LIST STYLE"*).
 *
 * The round trip runs numbering.xml -> styles.xml -> numbering.xml in ECMA-376
 * (§17.9.21 names a numbering style; its `<w:numPr>` names a `<w:num>`), and is
 * observable from this part alone because Word writes `<w:styleLink>` on the
 * definition at the far end as the pair of it (§17.9.28).
 */
describe('resolveLevel — numStyleLink delegation', () => {
  const LEVELS =
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl>';

  /** What Word writes for a list style: a delegating definition and the real one. */
  const DELEGATED =
    `<w:numbering ${W}>` +
    '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="HouseList"/></w:abstractNum>' +
    `<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="HouseList"/>${LEVELS}</w:abstractNum>` +
    '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="9"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>';

  it('follows the delegation to the definition whose styleLink matches', () => {
    const scheme = schemeFrom(DELEGATED);
    expect(scheme.abstractNums.get('0')!.numStyleLink).toBe('HouseList');
    expect(scheme.abstractNums.get('0')!.levels.size).toBe(0);
    expect(scheme.abstractNums.get('1')!.styleLink).toBe('HouseList');

    expect(resolveLevel(scheme, '3', 0)?.numFmt).toBe('decimal');
    expect(resolveLevel(scheme, '3', 0)?.lvlText).toBe('%1.');
    expect(resolveLevel(scheme, '3', 1)?.numFmt).toBe('lowerLetter');
  });

  it('still applies a lvlOverride on top of a delegated level', () => {
    // The two mechanisms compose: the instance overrides a level it reaches
    // THROUGH the style, which is what a "restart at 5" on a house list is.
    const scheme = schemeFrom(
      DELEGATED.replace(
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>',
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/>' +
          '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>',
      ),
    );
    expect(resolveLevel(scheme, '3', 0)?.start).toBe(5);
    expect(resolveLevel(scheme, '3', 0)?.numFmt).toBe('decimal');
  });

  it('prefers a level the definition declares ITSELF over the delegation', () => {
    // The hop is ADDITIVE: nothing that resolved before it resolves differently
    // after. Word never writes both, so this precedence is unobservable in a
    // real package — it is here to pin that the hop cannot change an existing
    // answer, which is what makes it safe for the renderer's gated consumers.
    const scheme = schemeFrom(
      DELEGATED.replace(
        '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="HouseList"/></w:abstractNum>',
        '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="HouseList"/>' +
          '<w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl>' +
          '</w:abstractNum>',
      ),
    );
    expect(resolveLevel(scheme, '3', 0)?.numFmt).toBe('upperRoman');
    // …and a level it does NOT declare still comes from the style.
    expect(resolveLevel(scheme, '3', 1)?.numFmt).toBe('lowerLetter');
  });

  it('returns undefined when no definition carries the matching styleLink', () => {
    const scheme = schemeFrom(
      DELEGATED.replace('<w:styleLink w:val="HouseList"/>', '<w:styleLink w:val="OtherList"/>'),
    );
    expect(resolveLevel(scheme, '3', 0)).toBeUndefined();
  });

  it('terminates on a delegation cycle', () => {
    // Two definitions naming each other's style. Malformed, and a `seen` set is
    // cheaper than discovering it as a hang in a task pane.
    const scheme = schemeFrom(
      `<w:numbering ${W}>` +
        '<w:abstractNum w:abstractNumId="0"><w:styleLink w:val="A"/><w:numStyleLink w:val="B"/></w:abstractNum>' +
        '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="B"/><w:numStyleLink w:val="A"/></w:abstractNum>' +
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>' +
        '</w:numbering>',
    );
    expect(resolveLevel(scheme, '3', 0)).toBeUndefined();
  });
});
