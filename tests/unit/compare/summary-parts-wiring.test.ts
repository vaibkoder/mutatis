/**
 * Compare — the auxiliary-part diffs reach the user.
 *
 * `driver.ts` called `diffPartsIfEnabled(...)` and `diffStyles(...)` and threw
 * both return values away, under a comment claiming they were "included in
 * summary note". Nothing consumed them, so `compare.excludeHeaders: false` —
 * the switch whose whole purpose is to compare headers — produced no output of
 * any kind. The work was done and discarded.
 *
 * These run through the DRIVER on purpose. The defect was not in either diff
 * (both were correct); it was in the two lines that dropped the results, and
 * only a test that goes through those lines can fail when they are reverted.
 *
 * ## Recorded limit
 * The acceptance asked for this on a clean document pair. It cannot be:
 * on such a pair both sides carry NO header parts at all, the
 * same three footer references and the same `titlePg` — the class measures zero
 * . So this is a synthetic pair built to carry exactly
 * the shape the real one no longer has.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort } from '@shared/host/index.js';
import { partsPackage, para, p, fnRef, W_NS } from '../../fixtures/parts-docx.js';

/** Set per test; `undefined` leaves the shipped default. */
let excludeHeaders: boolean | undefined;

vi.mock('@shared/config/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@shared/config/index.js')>();
  return {
    ...actual,
    getCompareConfig: (profile: Parameters<typeof actual.getCompareConfig>[0]) => ({
      ...actual.getCompareConfig(profile),
      ...(excludeHeaders === undefined ? {} : { excludeHeaders }),
    }),
  };
});

const { compareDriverState } = await import('@features/compare/driver.js');

/**
 * A .docx carrying a body, a default header, a styles part and a footnote the
 * body references — wired through the relationships, as Word writes it
 * (ADR-0047: a part is found by the section that references it, so a header
 * file nothing references is not a header).
 */
function makeDocx(
  name: string,
  bodyText: string,
  headerText: string,
  styleSz: string,
  footnoteText = 'See the Disclosure Letter for the agreed exceptions.',
): File {
  const bundle = partsPackage({
    body: [p(bodyText, fnRef(2))],
    sections: [{ headers: { default: 'header1.xml' } }],
    parts: { 'header1.xml': [para(headerText)] },
    footnotes: { 2: [footnoteText] },
  });
  bundle.setPart(
    'word/styles.xml',
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults/>` +
      `<w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:sz w:val="${styleSz}"/></w:rPr></w:style>` +
      '</w:styles>',
  );
  return new File([new Uint8Array(bundle.toUint8Array())], name);
}

async function runCompare(
  footnotes?: [string, string],
  headers: [string, string] = ['Project Sample — draft', 'Project Sample — execution'],
) {
  const d = compareDriverState;
  d.uploadDocx(
    'original',
    makeDocx(
      'orig.docx', 'The Seller shall deliver the Goods.', headers[0], '22',
      footnotes?.[0],
    ),
  );
  d.uploadDocx(
    'revised',
    makeDocx(
      'rev.docx', 'The Seller shall deliver the Goods.', headers[1], '24',
      footnotes?.[1],
    ),
  );
  await d.runCompare();
  expect(d.status.value).toBe('done');
  return d.summary.value!;
}

beforeEach(() => {
  setActiveHost({ writeBack: vi.fn().mockResolvedValue(undefined) } as unknown as HostPort);
});

afterEach(() => {
  _resetActiveHost();
  vi.restoreAllMocks();
  excludeHeaders = undefined;
});

describe('auxiliary-part diffs reach the change summary', () => {
  it('reports header changes at the default (§7.5 — headers compared by default, ADR-0047)', async () => {
    const summary = await runCompare();

    expect(summary.partChanges.headers).toBeGreaterThan(0);
  });

  it('reports NO header changes when compare.excludeHeaders is ON', async () => {
    excludeHeaders = true;
    const summary = await runCompare();

    // Not processed-and-hidden: not processed.
    expect(summary.partChanges.headers).toBe(0);
  });

  it('reports a style DEFINITION change, which is not gated by that switch at all', async () => {
    excludeHeaders = true;
    const summary = await runCompare();

    expect(summary.styleDefinitionNote).not.toBeNull();
    expect(summary.styleDefinitionNote).toContain('Style definitions changed');
    // The bodies are word-for-word identical, so nothing here came from the body
    // pipeline — this note is the styles diff arriving, and before the wiring it
    // was computed and dropped on the floor.
    expect(summary.total).toBe(0);
  });
});

/**
 * §9.8 — a footnote's CONTENT is substantive and compared BY DEFAULT.
 *
 * Two separate things were wrong and neither had a test. `compare.footnotes`
 * defaults to `'compare'`, so this class is live with nothing switched on — yet
 * `ChangeSummary.partChanges` was documented, in the spec and in `model.ts`, as
 * *"zero at the default config"*. And the change reached the reader only as a
 * line in the "Other parts" block: it was absent from `summary.total`, which
 * §8.2 defines as the full count, and from every `byCategory` bucket.
 */
describe('a footnote content change, at the DEFAULT config (§9.8)', () => {
  const ORIGINAL = 'See the Disclosure Letter for the agreed exceptions.';
  const REVISED = 'See the Disclosure Letter and the Tax Deed for the agreed exceptions.';
  /** The header unchanged, so every change counted here is the footnote's. */
  const SAME_HEADER: [string, string] = ['Project Sample — draft', 'Project Sample — draft'];

  it('is reported with nothing switched on', async () => {
    const summary = await runCompare([ORIGINAL, REVISED], SAME_HEADER);

    expect(excludeHeaders, 'this test must run at the default').toBeUndefined();
    expect(summary.partChanges.headers).toBe(0);
    expect(summary.partChanges.footers).toBe(0);
    expect(summary.partChanges.footnotes).toBeGreaterThan(0);
  });

  it('counts into `total`, classified by its content — never as formatting, because §9.8 says substantive', async () => {
    const summary = await runCompare([ORIGINAL, REVISED], SAME_HEADER);
    const footnotes = summary.partChanges.footnotes;

    // The body text is identical on both sides, so every counted change here
    // came from outside the body. Before this, `total` was 0 and the reader was
    // told "Total changes: 0" beside "Footnote changes: 1".
    expect(summary.total).toBe(footnotes);
    // ADR-0047: a note's ops go through the body's classifier, so "the Tax
    // Deed" added is filed by what it is (a defined term), not wholesale.
    expect(summary.byCategory.formatting).toBe(0);
    expect(summary.byCategory.whitespacePunct).toBe(0);
    expect(summary.byCategory.definedTerm + summary.byCategory.substantiveText).toBe(footnotes);
  });

  it('reports nothing when the footnotes are identical', async () => {
    const summary = await runCompare([ORIGINAL, ORIGINAL], SAME_HEADER);
    expect(summary.partChanges.footnotes).toBe(0);
    expect(summary.total).toBe(0);
  });
});
