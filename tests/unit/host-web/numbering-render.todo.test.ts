/**
 * host-web — renderListNumbers CONTRACT cases (authored `.skip` earlier;
 * live since the renderer implemented the counter state machine).
 *
 * Each case drives the real seams: `collectParaNumPrs` (extraction) +
 * `parseNumberingScheme` (IR) → `renderListNumbers`. Covers: multi-level counters
 * with deeper-level reset + ancestor-token substitution, level skip, startOverride,
 * lvlRestart=0 (continuous), lvlRestart≥1 (one-based ECMA-376 §17.9.10), isLgl
 * (all-decimal — a KNOWN DEFECT pinned as current behaviour, not a contract; see
 * that case), lowerLetter z→aa rollover, and unknown-numId/level omission.
 */
import { describe, it, expect } from 'vitest';
import { PizZipBundle } from '@shared/ooxml/index.js';
import { parseNumberingScheme } from '@host-web/numbering-scheme.js';
import { collectParaNumPrs } from '@host-web/ooxml-loader.js';
import { renderListNumbers, type ListNumber } from '@host-web/render-list-numbers.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** numId 1 → abstract 0 with the given per-level (numFmt, lvlText[, extra]) tuples. */
function abstract0(levels: Array<{ fmt: string; text: string; extra?: string }>): string {
  const lvls = levels
    .map(
      (l, i) =>
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${l.fmt}"/>` +
        `<w:lvlText w:val="${l.text}"/>${l.extra ?? ''}</w:lvl>`,
    )
    .join('');
  return (
    `<?xml version="1.0"?><w:numbering ${W}>` +
    `<w:abstractNum w:abstractNumId="0">${lvls}</w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'
  );
}

/** A body of list paragraphs, each at the given (numId, ilvl). */
function listBody(items: Array<{ numId: string; ilvl: number }>): string {
  const ps = items
    .map(
      it =>
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="${it.ilvl}"/><w:numId w:val="${it.numId}"/></w:numPr></w:pPr>` +
        '<w:r><w:t>item</w:t></w:r></w:p>',
    )
    .join('');
  return `<?xml version="1.0"?><w:document ${W}><w:body>${ps}</w:body></w:document>`;
}

function render(documentXml: string, numberingXml: string): Map<string, ListNumber> {
  const b = new PizZipBundle();
  b.setPart('word/document.xml', documentXml);
  b.setPart('word/numbering.xml', numberingXml);
  return renderListNumbers(parseNumberingScheme(b), collectParaNumPrs(b));
}

/** The listString sequence in document order (list paras only). */
function strings(documentXml: string, numberingXml: string): string[] {
  const m = render(documentXml, numberingXml);
  return collectParaNumPrs(bundleDoc(documentXml)).map(p => m.get(p.paraId)?.listString ?? '');
}

function bundleDoc(documentXml: string): PizZipBundle {
  const b = new PizZipBundle();
  b.setPart('word/document.xml', documentXml);
  return b;
}

describe('renderListNumbers — CONTRACT', () => {
  it('multi-level outline: deeper levels reset, ancestor tokens substituted', () => {
    const numbering = abstract0([
      { fmt: 'decimal', text: '%1.' },
      { fmt: 'decimal', text: '%1.%2' },
      { fmt: 'decimal', text: '%1.%2.%3' },
    ]);
    const body = listBody([
      { numId: '1', ilvl: 0 }, // 1.
      { numId: '1', ilvl: 1 }, // 1.1
      { numId: '1', ilvl: 1 }, // 1.2
      { numId: '1', ilvl: 2 }, // 1.2.1
      { numId: '1', ilvl: 0 }, // 2.   (advancing l0 resets l1+l2)
      { numId: '1', ilvl: 1 }, // 2.1  (l1 restarted at 1)
    ]);
    expect(strings(body, numbering)).toEqual(['1.', '1.1', '1.2', '1.2.1', '2.', '2.1']);
  });

  it('level skip: jumping straight to ilvl 2 initialises skipped levels to start', () => {
    const numbering = abstract0([
      { fmt: 'decimal', text: '%1.' },
      { fmt: 'decimal', text: '%1.%2' },
      { fmt: 'decimal', text: '%1.%2.%3' },
    ]);
    const body = listBody([{ numId: '1', ilvl: 2 }]); // no l0/l1 seen first
    expect(strings(body, numbering)).toEqual(['1.1.1']);
  });

  it('startOverride: counter begins at the overridden value', () => {
    const numbering =
      `<?xml version="1.0"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/>' +
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>' +
      '</w:numbering>';
    const body = listBody([
      { numId: '1', ilvl: 0 }, // 5.
      { numId: '1', ilvl: 0 }, // 6.
    ]);
    expect(strings(body, numbering)).toEqual(['5.', '6.']);
  });

  it('lvlRestart=0: the level never restarts when a shallower level advances', () => {
    const numbering = abstract0([
      { fmt: 'decimal', text: '%1.' },
      { fmt: 'decimal', text: '(%2)', extra: '<w:lvlRestart w:val="0"/>' },
    ]);
    const body = listBody([
      { numId: '1', ilvl: 0 }, // 1.
      { numId: '1', ilvl: 1 }, // (1)
      { numId: '1', ilvl: 1 }, // (2)
      { numId: '1', ilvl: 0 }, // 2.
      { numId: '1', ilvl: 1 }, // (3)  — continuous, NOT reset to (1)
    ]);
    expect(strings(body, numbering)).toEqual(['1.', '(1)', '(2)', '2.', '(3)']);
  });

  // PINS A KNOWN DEFECT, not the contract, and the `describe` above is called
  // CONTRACT — so read this one as *"what this build does today"*.
  //
  // `render-list-numbers.ts` applies `<w:isLgl/>` to the paragraph's OWN token as
  // well as its ancestors. **Word applies it to the ancestors and keeps the own
  // level's declared format** — measured against the reference rendering, a
  // `decimalZero` sub-level under `<w:isLgl/>` gives `1.01` where this gives
  // `1.1`. So the third row below is `1.1.a` in Word (`upperRoman` and
  // `lowerLetter` ancestors converted; the own `lowerLetter` kept) and `1.1.1`
  // here. Recorded in that module's fidelity bounds, in
  // `spec/features/16-render-core.md` and in ADR-0042; the fix is 08c's, **and
  // this assertion is what will go red when it lands — update it then, do not
  // "restore" it.**
  //
  // It matters beyond painting: Compare's `numberingKey` mirrored the same fold
  // and both detectors went blind together, so an edit to the own level's format
  // produced no op at all. The key now
  // carries the own level's DECLARED format as well as the folded one.
  it('isLgl: this build renders ALL substituted levels as decimal (Word keeps the own format)', () => {
    const numbering = abstract0([
      { fmt: 'upperRoman', text: '%1.' },
      { fmt: 'lowerLetter', text: '%1.%2' },
      { fmt: 'lowerLetter', text: '%1.%2.%3', extra: '<w:isLgl/>' },
    ]);
    const body = listBody([
      { numId: '1', ilvl: 0 }, // I.        (upperRoman, no isLgl)
      { numId: '1', ilvl: 1 }, // I.a       (lowerLetter ancestor + own)
      { numId: '1', ilvl: 2 }, // 1.1.1 here; 1.1.a in Word — see above
    ]);
    expect(strings(body, numbering)).toEqual(['I.', 'I.a', '1.1.1']);
  });

  it('lowerLetter z→aa rollover across a real scheme (with startOverride)', () => {
    const numbering =
      `<?xml version="1.0"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/>' +
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="26"/></w:lvlOverride></w:num>' +
      '</w:numbering>';
    const body = listBody([
      { numId: '1', ilvl: 0 }, // z)   (26)
      { numId: '1', ilvl: 0 }, // aa)  (27 — rollover)
    ]);
    expect(strings(body, numbering)).toEqual(['z)', 'aa)']);
  });

  // ── Edges pinned by the renderer (left open by the contract) ──────────

  it('lvlRestart=1 (one-based, ECMA-376 §17.9.10): restarts on an ilvl-0 advance but NOT on an ilvl-1 advance', () => {
    const numbering = abstract0([
      { fmt: 'decimal', text: '%1.' },
      { fmt: 'decimal', text: '%1.%2' },
      { fmt: 'decimal', text: '%1.%2.%3', extra: '<w:lvlRestart w:val="1"/>' },
    ]);
    const body = listBody([
      { numId: '1', ilvl: 0 }, // 1.
      { numId: '1', ilvl: 1 }, // 1.1
      { numId: '1', ilvl: 2 }, // 1.1.1
      { numId: '1', ilvl: 1 }, // 1.2     (ilvl-1 advance — does NOT restart l2)
      { numId: '1', ilvl: 2 }, // 1.2.2   (l2 continues)
      { numId: '1', ilvl: 0 }, // 2.      (ilvl-0 advance — one-based index 1 ≤ val 1 ⇒ restarts l2)
      { numId: '1', ilvl: 2 }, // 2.1.1   (l2 re-enters at start; l1 was default-reset too)
    ]);
    expect(strings(body, numbering)).toEqual([
      '1.', '1.1', '1.1.1', '1.2', '1.2.2', '2.', '2.1.1',
    ]);
  });

  it('default restart applies even when the advancing level is at its first sighting', () => {
    const numbering = abstract0([
      { fmt: 'decimal', text: '%1.' },
      { fmt: 'decimal', text: '%1.%2' },
    ]);
    const body = listBody([
      { numId: '1', ilvl: 1 }, // 1.1   (l0 skipped — initialises to start)
      { numId: '1', ilvl: 1 }, // 1.2
      { numId: '1', ilvl: 0 }, // 2.    (first ADVANCE of l0: init was 1, so 2; resets l1)
      { numId: '1', ilvl: 1 }, // 2.1
    ]);
    expect(strings(body, numbering)).toEqual(['1.1', '1.2', '2.', '2.1']);
  });

  it('unknown numId or undefined level: the paragraph emits nothing (absent from the map)', () => {
    const numbering = abstract0([{ fmt: 'decimal', text: '%1.' }]); // numId 1, level 0 only
    const body = listBody([
      { numId: '9', ilvl: 0 }, // unknown numId
      { numId: '1', ilvl: 3 }, // known numId, level not defined
      { numId: '1', ilvl: 0 }, // valid — still numbered, counters unaffected by the skips
    ]);
    const rendered = render(body, numbering);
    expect(rendered.size).toBe(1);
    // `path` is the ordinals behind the token (ADR-0042) — [1] here, the level's start.
    expect(rendered.get('body-para-2')).toEqual({ listString: '1.', listLevel: 0, path: [1] });
    expect(strings(body, numbering)).toEqual(['', '', '1.']);
  });
});

/**
 * Counters belong to the abstract DEFINITION, not to the `<w:num>`.
 *
 * Read from the format rather than from a claim about Word: `<w:startOverride>`
 * exists to restart a level for one instance, and it would be a no-op if every
 * fresh instance restarted by itself. Word writes a second `<w:num>` over the
 * same `<w:abstractNum>` WITH an override when the user asks to restart, and
 * WITHOUT one when the user asks to continue — so both of those must work.
 *
 * Keyed per `numId` instead, this module invented restarts Word does not render
 * and hid real ones, and because Compare's labels come from here, both showed up
 * as ADR-0008 defects: an insert in one group moved the other group's numbers and
 * reported nothing, while a clause moved between instances reported a cascade of
 * fabricated numbers.
 */
describe('renderListNumbers — counters belong to the definition', () => {
  /** Two instances over one abstract; instance 2 carries `override` verbatim. */
  const twoInstances = (override: string) =>
    `<?xml version="1.0"?><w:numbering ${W}>` +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    `<w:num w:numId="2"><w:abstractNumId w:val="0"/>${override}</w:num>` +
    '</w:numbering>';

  const FOUR = listBody([
    { numId: '1', ilvl: 0 },
    { numId: '1', ilvl: 0 },
    { numId: '2', ilvl: 0 },
    { numId: '2', ilvl: 0 },
  ]);

  it('continues the sequence across two instances with no override', () => {
    expect(strings(FOUR, twoInstances(''))).toEqual(['1.', '2.', '3.', '4.']);
  });

  it('restarts at the override when the second instance carries one', () => {
    const override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>';
    expect(strings(FOUR, twoInstances(override))).toEqual(['1.', '2.', '1.', '2.']);
  });

  it('restarts at a value the definition never reaches, and counts on from it', () => {
    const override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride>';
    expect(strings(FOUR, twoInstances(override))).toEqual(['1.', '2.', '9.', '10.']);
  });

  it('applies the override ONCE, on the instance\'s first use of that level', () => {
    // An instance that leaves and returns does not restart a second time — the
    // override is a restart, not a per-paragraph start value.
    const alternating = listBody([
      { numId: '1', ilvl: 0 },
      { numId: '2', ilvl: 0 },
      { numId: '1', ilvl: 0 },
      { numId: '2', ilvl: 0 },
    ]);
    const override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>';
    expect(strings(alternating, twoInstances(override))).toEqual(['1.', '5.', '6.', '7.']);
  });

  it('restarts on the instance\'s first touch of a level, whether it advances or READS it', () => {
    // Found in review as a regression: the restart
    // fired only when a paragraph SAT at the overridden level, while an ancestor
    // token for that level was read straight off the shared counter. An instance
    // that restarts at 1 and opens with a sub-item therefore rendered the
    // previous instance's parent number:
    //
    //     before   1.  2.  2.1  1.  2.      (and `2.1 2.2` where the instance
    //     after    1.  2.  1.1  2.  3.       never uses the overridden level)
    //
    // The `1.1` then `2.` is the module's existing contract for a skipped
    // shallower level — it initialises and KEEPS the value, so the instance's
    // own first level-0 paragraph advances from it.
    const override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>';
    const deepFirst = listBody([
      { numId: '1', ilvl: 0 },
      { numId: '1', ilvl: 0 },
      { numId: '2', ilvl: 1 },
      { numId: '2', ilvl: 0 },
      { numId: '2', ilvl: 0 },
    ]);
    expect(strings(deepFirst, twoInstances(override))).toEqual(['1.', '2.', '1.1', '2.', '3.']);

    // And an instance that never uses the overridden level at all still restarts
    // it — the override is a property of the instance, not of a paragraph.
    const onlyDeep = listBody([
      { numId: '1', ilvl: 0 },
      { numId: '1', ilvl: 0 },
      { numId: '2', ilvl: 1 },
      { numId: '2', ilvl: 1 },
    ]);
    expect(strings(onlyDeep, twoInstances(override))).toEqual(['1.', '2.', '1.1', '1.2']);
  });

  it('shares the sequence when the two instances reach the definition by different routes', () => {
    // One instance points at the delegating definition, the other at the one the
    // list style names. Same list, so one sequence.
    const viaStyle =
      `<?xml version="1.0"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="HouseList"/></w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="HouseList"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
      '</w:numbering>';
    expect(strings(FOUR, viaStyle)).toEqual(['1.', '2.', '3.', '4.']);
  });
});
