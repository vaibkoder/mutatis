/**
 * A header, footer or note change is a row of the change list.: classified by content, badged with its location,
 * navigable inside its own part, and exported with its own text.
 *
 * The hazard this guards (`mergePartsIntoSummary`'s old comment): a part op's
 * range indexes its OWN part's paragraphs, so a reader that resolves it against
 * the body's arrays lands on an unrelated body paragraph with the same index.
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { render } from 'preact';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import type { HostPort, SearchScope } from '@shared/host/index.js';
import { ChangeList } from '@features/compare/components/ChangeList.js';
import { buildChangeSet } from '@features/compare/change-set-export.js';
import { partLabel } from '@features/ai-export/index.js';
import { headerFooterLocation, noteLocation } from '@features/compare/parts-diff.js';
import type { DiffOp } from '@features/compare/model.js';

import { partsPackage, p, run, fnRef, type PackageSpec } from '../../fixtures/parts-docx.js';

let config: Record<string, unknown> = {};
vi.mock('@shared/config/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@shared/config/index.js')>();
  return {
    ...actual,
    getCompareConfig: (profile: Parameters<typeof actual.getCompareConfig>[0]) => ({
      ...actual.getCompareConfig(profile),
      ...config,
    }),
  };
});

const { createCompareDriver } = await import('@features/compare/driver.js');

const findAndReveal = vi.fn(async (_text: string, _scope?: SearchScope) => {});

beforeEach(() => {
  findAndReveal.mockClear();
  setActiveHost({ writeBack: vi.fn(async () => {}), findAndReveal } as unknown as HostPort);
});

afterEach(() => {
  _resetActiveHost();
  config = {};
});

const file = (spec: PackageSpec, name: string): File =>
  new File([new Uint8Array(partsPackage(spec).toUint8Array())], name);

async function compare(original: PackageSpec, revised: PackageSpec) {
  const driver = createCompareDriver();
  driver.uploadDocx('original', file(original, 'original.docx'));
  driver.uploadDocx('revised', file(revised, 'revised.docx'));
  await driver.runCompare();
  expect(driver.errorMessage.value).toBeNull();
  return driver;
}

const NOTES: PackageSpec = {
  body: [
    p('The Seller', fnRef(1), ' shall deliver', fnRef(2), ' the Goods on the Completion Date.'),
    p('The Buyer shall pay the Price', fnRef(3), ' within 30 days.'),
    'Clause three.',
  ],
  footnotes: {
    1: ['As defined in the Framework Agreement.'],
    2: ['Delivery is to the Buyer’s premises.'],
    3: ['Exclusive of GST.'],
  },
};

/** Note 2 edited (now 1), note 1 removed, a note added after "Clause three". */
const NOTES_REVISED: PackageSpec = {
  body: [
    p('The Seller shall deliver', fnRef(1), ' the Goods on the Completion Date.'),
    p('The Buyer shall pay the Price', fnRef(2), ' within 30 days.'),
    p('Clause three', fnRef(3), '.'),
  ],
  footnotes: {
    1: ['Delivery is to the Seller’s premises.'],
    2: ['Exclusive of GST.'],
    3: ['Inserted at the Buyer’s request.'],
  },
};

const noteRows = (driver: Awaited<ReturnType<typeof compare>>): DiffOp[] =>
  (driver.ops.value ?? []).filter(op => op.partId.startsWith('footnote:'));

describe('part changes in the change list', () => {
  it('a footnote content change is one navigable row, classified by content, labelled with its note', async () => {
    const driver = await compare(NOTES, NOTES_REVISED);
    const edited = noteRows(driver).filter(op => op.type === 'replace');
    expect(edited).toHaveLength(1);
    expect(edited[0]!.classification).not.toBe('formatting');
    expect(driver.parts.value.get(edited[0]!.partId)?.location.label).toBe('Footnote 1');

    driver.goToChange(edited[0]!.id);
    expect(findAndReveal).toHaveBeenCalledTimes(1);
    expect(findAndReveal.mock.calls[0]).toEqual([expect.stringMatching(/Seller/), { story: 'footnote' }]);
  });

  it('an added and a removed note are labelled so', async () => {
    const driver = await compare(NOTES, NOTES_REVISED);
    const labels = [...new Set(noteRows(driver).map(op => driver.parts.value.get(op.partId)?.location.label))];
    expect(labels).toEqual(expect.arrayContaining(['Footnote 1 (removed)', 'Footnote 3 (added)']));
  });

  it('the body’s rows carry no location and navigate the main text', async () => {
    const driver = await compare(NOTES, NOTES_REVISED);
    const body = (driver.ops.value ?? []).filter(op => op.type !== 'equal' && op.partId === 'document');
    expect(body.length).toBeGreaterThan(0);
    expect(driver.parts.value.has('document')).toBe(false);
    driver.goToChange(body[0]!.id);
    expect(findAndReveal.mock.calls[0]?.[1]).toBeUndefined();
  });

  it('a footer row reads its OWN paragraph, never the body paragraph at the same index', async () => {
    config = { excludeFooters: false };
    const styled = (style: string): string =>
      `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${run('Strictly confidential')}</w:p>`;
    const spec = (style: string): PackageSpec => ({
      body: ['Clause one applies.', 'Clause two applies.'],
      sections: [{ footers: { default: 'footer1.xml' } }],
      parts: { 'footer1.xml': [styled(style)] },
    });
    const driver = await compare(spec('Footer'), spec('FooterBold'));
    const rows = (driver.ops.value ?? []).filter(op => op.partId.startsWith('footer:'));
    expect(rows).toHaveLength(1);
    expect(rows[0]!.type).toBe('property');
    expect(driver.parts.value.get(rows[0]!.partId)?.location.label).toBe('Footer');

    driver.goToChange(rows[0]!.id);
    expect(findAndReveal).toHaveBeenCalledWith(
      'Strictly confidential', { story: 'footer', section: 0, type: 'default' },
    );
  });

  it('the summary still counts every part change once', async () => {
    const driver = await compare(NOTES, NOTES_REVISED);
    const nonEqual = (driver.ops.value ?? []).filter(op => op.type !== 'equal');
    expect(driver.summary.value?.total).toBe(nonEqual.length);
  });
});

describe('location labels', () => {
  it('names a header or footer by its type, and its section only when there are several', () => {
    expect(headerFooterLocation('footer', 0, 'default', 1).label).toBe('Footer');
    expect(headerFooterLocation('header', 0, 'first', 1).label).toBe('First-page header');
    expect(headerFooterLocation('header', 1, 'even', 3).label).toBe('Even-page header, section 2');
    expect(headerFooterLocation('header', 1, 'even', 3).scope).toEqual({ story: 'header', section: 1, type: 'even' });
  });

  it('names a note by its number and whether it was added or removed', () => {
    expect(noteLocation('footnote', 3, 'kept').label).toBe('Footnote 3');
    expect(noteLocation('footnote', 3, 'inserted').label).toBe('Footnote 3 (added)');
    expect(noteLocation('endnote', 1, 'deleted').label).toBe('Endnote 1 (removed)');
    expect(noteLocation('endnote', null, null).label).toBe('Endnotes');
    expect(noteLocation('footnote', 2, 'kept').hint).toMatch(/WordApi 1\.5.*main text/);
  });

  it('the export’s fallback label knows endnotes', () => {
    expect(partLabel('endnote:3')).toBe('Endnotes');
    expect(partLabel('endnotes')).toBe('Endnotes');
  });
});

// ---------------------------------------------------------------------------
// The badge, and the export
// ---------------------------------------------------------------------------

const op = (id: string, partId: string, start: number): DiffOp => ({
  id, type: 'delete', origRange: { start, end: start }, revRange: null, partId,
  classification: 'substantiveText', classes: ['substantiveText'], author: 'Compare', date: '2026-10-26T00:00:00Z',
});

let host: HTMLDivElement | null = null;
afterEach(() => {
  if (host) render(null, host);
  host?.remove();
  host = null;
});

describe('the location badge', () => {
  it('shows a part row’s location with its hint, and nothing on a body row', () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    const location = noteLocation('footnote', 2, 'kept');
    render(
      <ChangeList
        ops={[op('op-1', 'document', 0), op('footnote:7/op-1', 'footnote:7', 0)]}
        filters={{ hideFormattingOnly: false, collapseRenumbering: false }}
        locations={new Map([['footnote:7', { location }]])}
        onGoToChange={() => {}}
      />,
      host,
    );
    const badges = [...host.querySelectorAll('.dp-compare__change-location')];
    expect(badges.map(b => b.textContent)).toEqual(['Footnote 2']);
    expect(badges[0]!.getAttribute('title')).toBe(location.hint);
  });
});

describe('the change-set export', () => {
  it('prints a part op’s location and its own paragraph’s text', () => {
    const rec = (text: string) => ({ normText: text }) as never;
    const env = buildChangeSet(
      [
        op('op-1', 'document', 0),
        op('footer:s0:default/op-1', 'footer:s0:default', 0),
        op('footnote:9/op-1', 'footnote:9', 0),
      ],
      [rec('Clause one applies.')],
      [],
      new Map([
        ['footer:s0:default', {
          originalParas: [rec('Strictly confidential')], revisedParas: [],
          location: headerFooterLocation('footer', 0, 'default', 1),
        }],
        ['footnote:9', {
          originalParas: [rec('Exclusive of GST.')], revisedParas: [],
          location: noteLocation('footnote', 2, 'deleted'),
        }],
      ]),
    );
    const entries = env!.change_set!.changes;
    expect(entries.map(e => [e.clause, e.original])).toEqual([
      ['Body', 'Clause one applies.'],
      ['Footer', 'Strictly confidential'],
      ['Footnote 2 (removed)', 'Exclusive of GST.'],
    ]);
  });
});

