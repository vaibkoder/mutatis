/**
 * Compare at scale — the batch driver and the artefact exits (ADR-0037; §5.3–5.5).
 *
 * The driver holds `File` handles, never bytes: a hundred-variant batch must
 * not cost a hundred `ArrayBuffer`s before the first comparison starts. And it
 * is instantiable from the first line, with no module-level singleton — the
 * single-pair driver had one and it was the prerequisite that blocked this
 * whole feature, so repeating it in the module built for running many
 * comparisons would be perverse.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  styleNumberedFile,
  HEAD,
  SUB_A,
  SUB_B,
  INSERTED,
  type StyledPara,
} from '../../fixtures/style-numbered-docx.js';
import { createBatchDriver } from '@features/compare/batch/driver.js';
import { PizZipBundle } from '@shared/ooxml/bundle.js';
import { emptyClassificationTally } from '@features/compare/batch/model.js';
import type { ArtefactHandle, BatchRow } from '@features/compare/batch/model.js';
import {
  downloadArtefact,
  downloadRowArchive,
  isDownloadable,
  redlineOf,
} from '@features/compare/batch/artefacts.js';

const TEMPLATE_TEXT = [
  'The Seller shall pay the Purchase Price on Completion.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

/**
 * A different house template, so "which template did this run against?" has a
 * visible answer: it shares no sentence with `TEMPLATE_TEXT`, and a redline
 * built from it says so in its own deleted text.
 */
const OTHER_TEMPLATE_TEXT = [
  'The Contractor shall complete the Works by the Long Stop Date.',
  'Liquidated damages accrue at the rate stated in the Schedule.',
].join('\n');

function docxFile(name: string, text: string): File {
  const bytes = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return new File([bytes], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/**
 * A `.docx` carrying a real tracked insertion, INSIDE the first paragraph.
 *
 * `synthesiseDocx` emits no revision markup, so a flatten policy has no effect
 * on its output and any test asserting acceptAll ≠ rejectAll would pass
 * vacuously. The placement matters as much as the markup: a whole tracked
 * PARAGRAPH rejects down to an empty paragraph, which is still one change
 * against a template that has no paragraph there — so the counts would match
 * either way. Inside an existing paragraph, accepting changes that paragraph's
 * text and rejecting restores it exactly, which is a difference a count can see.
 */
function trackedChangeFile(name: string): File {
  const bundle = synthesiseDocx(TEMPLATE_TEXT);
  const xml = bundle.part('word/document.xml') as string;
  bundle.setPart(
    'word/document.xml',
    xml.replace(
      '</w:p>',
      '<w:ins w:id="900" w:author="Counterparty" w:date="2026-01-01T00:00:00Z">' +
        '<w:r><w:t xml:space="preserve"> and in cleared funds</w:t></w:r>' +
        '</w:ins></w:p>',
    ),
  );
  return new File([new Uint8Array(bundle.toUint8Array())], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function row(id: string, over: Partial<BatchRow> = {}): BatchRow {
  return {
    id,
    label: `${id}.docx`,
    status: 'done',
    changeCount: 3,
    byClassification: emptyClassificationTally(),
    quarantinedOriginal: 0,
    quarantinedRevised: 0,
    artefacts: [artefact(`${id} (redline).docx`)],
    artefactOnDemand: false,
    durationMs: 5,
    ...over,
  };
}

function artefact(fileName: string): ArtefactHandle {
  const blob = new Blob([new Uint8Array([1, 2, 3, 4])]);
  return { kind: 'redlineDocx', fileName, byteLength: 4, blob };
}

describe('createBatchDriver — instances are independent', () => {
  it('shares no state between two drivers', () => {
    const a = createBatchDriver();
    const b = createBatchDriver();

    a.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    a.addVariants([docxFile('acme.docx', TEMPLATE_TEXT)]);
    a.progress.value = 40;

    expect(b.template.value).toBeNull();
    expect(b.variants.value).toEqual([]);
    expect(b.progress.value).toBe(0);
  });

  it('exports no module-level singleton to scribble on', async () => {
    const mod = await import('@features/compare/batch/driver.js');
    expect(Object.keys(mod)).toEqual(['createBatchDriver']);
  });
});

describe('slot management holds files, not bytes', () => {
  it('records a template slot from a File without reading it', () => {
    const driver = createBatchDriver();
    const file = docxFile('house-nda.docx', TEMPLATE_TEXT);
    const spy = vi.spyOn(file, 'arrayBuffer');

    driver.setTemplate(file);

    expect(driver.template.value).toMatchObject({ label: 'house-nda.docx' });
    expect(driver.template.value?.byteLength).toBe(file.size);
    expect(spy).not.toHaveBeenCalled();
  });

  it('queues variants in order and gives each a distinct slot id', () => {
    const driver = createBatchDriver();
    driver.addVariants([
      docxFile('a.docx', TEMPLATE_TEXT),
      docxFile('b.docx', TEMPLATE_TEXT),
    ]);

    expect(driver.variants.value.map(v => v.label)).toEqual(['a.docx', 'b.docx']);
    expect(new Set(driver.variants.value.map(v => v.id)).size).toBe(2);
  });

  it('removes and clears variants', () => {
    const driver = createBatchDriver();
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT), docxFile('b.docx', TEMPLATE_TEXT)]);
    const first = driver.variants.value[0]?.id as string;

    driver.removeVariant(first);
    expect(driver.variants.value.map(v => v.label)).toEqual(['b.docx']);

    driver.clearVariants();
    expect(driver.variants.value).toEqual([]);
  });

  it('clearing resets the whole batch, not just the queue', async () => {
    // Otherwise the rows of a finished batch survive their files and the panel
    // offers buttons for pairs that can no longer be re-run.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT.replace('Completion.', 'Friday.'))]);
    await driver.runBatch();

    expect(driver.rows.value).toHaveLength(1);
    expect(driver.index.value).not.toBeNull();
    expect(driver.status.value).toBe('done');

    driver.clearVariants();

    expect(driver.variants.value).toEqual([]);
    expect(driver.rows.value).toEqual([]);
    expect(driver.index.value).toBeNull();
    expect(driver.status.value).toBe('idle');
    expect(driver.progress.value).toBe(0);
    expect(driver.errorMessage.value).toBeNull();
  });
});

describe('runBatch', () => {
  it('refuses without a template, with a message rather than a throw', async () => {
    const driver = createBatchDriver();
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT)]);

    await driver.runBatch();

    expect(driver.status.value).toBe('error');
    expect(driver.errorMessage.value).toMatch(/template/i);
  });

  it('refuses with no variants', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));

    await driver.runBatch();

    expect(driver.status.value).toBe('error');
    expect(driver.errorMessage.value).toMatch(/at least one variant/i);
  });

  it('runs a real batch end to end and publishes rows and an index', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
      docxFile('borealis.docx', TEMPLATE_TEXT),
    ]);

    await driver.runBatch();

    expect(driver.status.value).toBe('done');
    expect(driver.progress.value).toBe(100);
    expect(driver.rows.value).toHaveLength(2);
    expect(driver.index.value?.mode).toBe('one-to-many');
    expect(driver.index.value?.templateLabel).toBe('house-nda.docx');

    const acme = driver.rows.value.find(r => r.label === 'acme.docx');
    const borealis = driver.rows.value.find(r => r.label === 'borealis.docx');
    expect(acme?.changeCount).toBeGreaterThan(0);
    expect(borealis?.changeCount).toBe(0);
  });

  it('publishes rows live as pairs finish, not only at the end', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT)]);

    const seen: number[] = [];
    const stop = driver.rows.subscribe(rows => seen.push(rows.length));
    await driver.runBatch();
    stop();

    // 0 (reset) then at least one intermediate publish before the final set.
    expect(seen.length).toBeGreaterThan(2);
  });

  it('is re-entrant per instance — a second run while one is going is a no-op', async () => {
    const driver = createBatchDriver();
    driver.status.value = 'running';

    await driver.runBatch();

    // Untouched: it returned at the guard rather than clobbering the live run.
    expect(driver.errorMessage.value).toBeNull();
    expect(driver.rows.value).toEqual([]);
  });
});

describe('artefact exits', () => {
  let clicked: { name: string; blob: Blob }[];
  let created: string[];

  beforeEach(() => {
    clicked = [];
    created = [];
    vi.stubGlobal('URL', {
      ...URL,
      createObjectURL: (blob: Blob) => {
        created.push('url');
        lastBlob = blob;
        return 'blob:test';
      },
      revokeObjectURL: () => {},
    });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
      this: HTMLAnchorElement,
    ) {
      clicked.push({ name: this.download, blob: lastBlob });
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  let lastBlob: Blob;

  it('downloads one row by its artefact name', async () => {
    await downloadArtefact(row('acme'));

    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('acme (redline).docx');
    expect(created).toHaveLength(1);
  });

  it('re-runs the pair for a row whose redline was not retained', async () => {
    // Past the artefact ceiling a row keeps no Blob. The bytes are still
    // obtainable — that is what makes the ceiling a disclosed degradation
    // rather than a hundred results the user can never reach (spec §4.3 step 6).
    const produced = new Blob([new Uint8Array([80, 75, 3, 4])]);
    const produce = vi.fn(async () => ({
      kind: 'redlineDocx' as const,
      fileName: 'acme (redline).docx',
      byteLength: 4,
      blob: produced,
    }));

    await downloadArtefact(row('acme', { artefacts: [], artefactOnDemand: true }), produce);

    expect(produce).toHaveBeenCalledTimes(1);
    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('acme (redline).docx');
  });

  it('does nothing for an on-demand row when no producer is supplied', async () => {
    await downloadArtefact(row('acme', { artefacts: [], artefactOnDemand: true }));
    expect(clicked).toHaveLength(0);
  });

  it('never re-runs a pair whose redline is already retained', async () => {
    const produce = vi.fn();
    await downloadArtefact(row('acme'), produce);
    expect(produce).not.toHaveBeenCalled();
    expect(clicked).toHaveLength(1);
  });

  it('archives every downloadable row under one name derived from the template', async () => {
    await downloadRowArchive([row('a'), row('b'), row('c')], 'house-nda.docx');

    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('house-nda — redlines.zip');
    expect(clicked[0]?.blob.size).toBeGreaterThan(0);
  });

  it('skips rows that yield nothing rather than writing an empty entry', async () => {
    // An empty .docx in the archive would read as "compared, no changes",
    // which is a different and wrong claim from "not retained".
    const rows = [
      row('good'),
      row('failed', { status: 'failed', artefacts: [], reason: 'corrupt' }),
    ];

    await downloadRowArchive(rows, 'house-nda.docx');

    expect(clicked).toHaveLength(1);
    expect(rows.filter(isDownloadable)).toHaveLength(1);
  });

  it('re-runs on-demand rows lazily, one at a time, as the archive is written', async () => {
    // Producing them all up front would rebuild the exact memory peak that
    // made them on-demand in the first place.
    let live = 0;
    let peak = 0;
    const produce = vi.fn(async (r: BatchRow) => {
      live += 1;
      peak = Math.max(peak, live);
      await Promise.resolve();
      live -= 1;
      return {
        kind: 'redlineDocx' as const,
        fileName: `${r.label} (redline).docx`,
        byteLength: 4,
        blob: new Blob([new Uint8Array([80, 75, 3, 4])]),
      };
    });

    const rows = [
      row('retained'),
      row('a', { artefacts: [], artefactOnDemand: true }),
      row('b', { artefacts: [], artefactOnDemand: true }),
    ];

    await downloadRowArchive(rows, 'house-nda.docx', produce);

    expect(produce).toHaveBeenCalledTimes(2);
    expect(peak).toBe(1);
    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('house-nda — redlines.zip');
  });

  it('writes nothing at all when no row is downloadable', async () => {
    await downloadRowArchive([row('failed', { status: 'failed', artefacts: [] })], 'x.docx');
    expect(clicked).toHaveLength(0);
  });
});

describe('artefact predicates', () => {
  it('finds the redline artefact of a row', () => {
    expect(redlineOf(row('a'))?.fileName).toBe('a (redline).docx');
    expect(redlineOf(row('a', { artefacts: [] }))).toBeUndefined();
  });

  it('treats an on-demand row as downloadable — its redline is obtainable, just not retained', () => {
    expect(isDownloadable(row('a'))).toBe(true);
    expect(isDownloadable(row('a', { artefacts: [], artefactOnDemand: true }))).toBe(true);
    // Not compared, or compared and failed: there is nothing to produce.
    expect(isDownloadable(row('a', { status: 'failed' }))).toBe(false);
    expect(isDownloadable(row('a', { status: 'cancelled', artefacts: [] }))).toBe(false);
    expect(isDownloadable(row('a', { artefacts: [], artefactOnDemand: false }))).toBe(false);
  });
});

describe('the driver\'s own on-demand producer (B11/B12, end to end)', () => {
  // The tests above prove the plumbing with an injected stub. This proves the
  // real thing: `createBatchDriver` re-running the pair, assembling the package
  // and handing over actual `.docx` bytes. That packaging half is exactly what
  // B2 was reopened for, so testing it against a stub would restate the bug.
  const clicked: { name: string; blob: Blob }[] = [];

  beforeEach(() => {
    clicked.length = 0;
    vi.stubGlobal('URL', {
      ...URL,
      createObjectURL: (blob: Blob) => { lastBlob = blob; return 'blob:x'; },
      revokeObjectURL: () => {},
    });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
      this: HTMLAnchorElement,
    ) {
      clicked.push({ name: this.download, blob: lastBlob });
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  let lastBlob: Blob;

  /** A driver whose artefact ceiling is crossed immediately. */
  async function ranOverBudget() {
    const driver = createBatchDriver({ artefactBudgetBytes: 1 });
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
      docxFile('beta.docx', TEMPLATE_TEXT.replace('on Completion.', 'within ten days.')),
    ]);
    await driver.runBatch();
    return driver;
  }

  it('retains nothing past the ceiling, yet still produces real .docx bytes on download', async () => {
    const driver = await ranOverBudget();

    expect(driver.status.value).toBe('done');
    const rows = driver.index.value?.rows ?? [];
    expect(rows).toHaveLength(2);
    expect(rows.every(r => r.status === 'done')).toBe(true);
    expect(rows.every(r => r.artefacts.length === 0)).toBe(true);
    expect(rows.every(r => r.artefactOnDemand)).toBe(true);
    expect(driver.index.value?.disclosures).toHaveLength(1);

    await driver.downloadRow(rows[0]?.id as string);

    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('acme (redline).docx');
    // A real package, not an empty Blob: `PK\x03\x04` is the ZIP local header.
    const head = new Uint8Array(await (clicked[0]?.blob as Blob).arrayBuffer()).slice(0, 4);
    expect(Array.from(head)).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect((clicked[0]?.blob as Blob).size).toBeGreaterThan(1000);
    expect(driver.errorMessage.value).toBeNull();
  });

  it('archives every on-demand row by re-running each of them', async () => {
    const driver = await ranOverBudget();
    await driver.downloadAll();

    expect(clicked).toHaveLength(1);
    expect(clicked[0]?.name).toBe('house-nda — redlines.zip');
    expect((clicked[0]?.blob as Blob).size).toBeGreaterThan(2000);
  });

  it('re-runs under the batch\'s own flatten choice, not whatever is set now', async () => {
    // Spec §8 case 7: the index records what the batch ran under, and a re-run
    // that read the live signals could assemble a redline the index denies.
    //
    // The fixture has to carry a REAL tracked insertion. A `synthesiseDocx`
    // document contains no `w:ins`/`w:del`, so acceptAll and rejectAll flatten
    // it identically and this test would pass with the fix ripped out — which
    // is what a review caught it doing.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([trackedChangeFile('acme.docx')]);
    await driver.runBatch();

    expect(driver.index.value?.flatten).toEqual({ template: 'acceptAll', variant: 'acceptAll' });
    const rows = driver.index.value?.rows ?? [];
    const underAcceptAll = rows[0]?.changeCount as number;
    expect(underAcceptAll).toBeGreaterThan(0);

    // Flipping the live signals must not reach a re-run of a finished batch.
    driver.flattenTemplate.value = 'rejectAll';
    driver.flattenVariant.value = 'rejectAll';

    const detail = await driver.openPair(rows[0]?.id as string);
    expect(detail?.summary.total).toBe(underAcceptAll);
  });

  it('the flatten fixture really does change the answer — the guard above can fail', async () => {
    // The control for the test above: run the SAME pair under rejectAll from
    // the start and confirm the count differs. Without this, "the numbers
    // matched" proves nothing.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([trackedChangeFile('acme.docx')]);
    driver.flattenVariant.value = 'rejectAll';
    await driver.runBatch();
    const underReject = driver.index.value?.rows[0]?.changeCount as number;

    const other = createBatchDriver();
    other.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    other.addVariants([trackedChangeFile('acme.docx')]);
    await other.runBatch();
    const underAccept = other.index.value?.rows[0]?.changeCount as number;

    // Accepting the insertion changes the first paragraph; rejecting it
    // restores the template exactly.
    expect(underAccept).toBeGreaterThan(0);
    expect(underReject).toBe(0);
    expect(underReject).not.toBe(underAccept);
  });

  it('re-runs against the batch\'s own template, not whichever one is chosen now', async () => {
    // The sibling of the flatten pin, and the easier one to lose. Choosing a
    // new template after a batch finishes must not silently re-point every
    // drill-down and every on-demand redline at it: the rows still carry the
    // old template's counts, so a re-run against the new one would produce a
    // redline the index on screen flatly denies. `removeVariant`'s refusal
    // pins the variant files; nothing but this pins the template.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
    ]);
    await driver.runBatch();

    const rowId = driver.index.value?.rows[0]?.id as string;
    const asRun = driver.index.value?.rows[0]?.changeCount as number;
    expect(asRun).toBeGreaterThan(0);

    // The file input is live after a batch — this is a user action, not a bug.
    driver.setTemplate(docxFile('construction.docx', OTHER_TEMPLATE_TEXT));

    const detail = await driver.openPair(rowId);
    expect(detail?.summary.total).toBe(asRun);
  });

  it('the template fixture really does change the answer — the guard above can fail', async () => {
    // Control. Without it, "the counts matched" could just mean the two
    // templates happen to compare identically against this variant.
    const other = createBatchDriver();
    other.setTemplate(docxFile('construction.docx', OTHER_TEMPLATE_TEXT));
    other.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
    ]);
    await other.runBatch();

    const against = createBatchDriver();
    against.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    against.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
    ]);
    await against.runBatch();

    expect(other.index.value?.rows[0]?.changeCount)
      .not.toBe(against.index.value?.rows[0]?.changeCount);
  });

  it('builds the on-demand redline from the pinned template too', async () => {
    // `openPair` and `produceArtefact` share `comparePair`, but the artefact
    // path also carries the template BUNDLE — the package the redline is
    // assembled into. A redline whose body came from one template and whose
    // styles, numbering and headers came from another would be a real document
    // that no comparison ever produced.
    const driver = await ranOverBudget();
    driver.setTemplate(docxFile('construction.docx', OTHER_TEMPLATE_TEXT));

    const rows = driver.index.value?.rows ?? [];
    await driver.downloadRow(rows[0]?.id as string);

    expect(clicked).toHaveLength(1);
    const bundle = new PizZipBundle(await (clicked[0]?.blob as Blob).arrayBuffer());
    const xml = bundle.part('word/document.xml') as string;
    expect(xml).toContain('Purchase Price');
    expect(xml).not.toContain('Long Stop Date');
  });

  it('clears a stale failure when the next download works', async () => {
    // Reported by review: the message is only ever set, never cleared, so one
    // failed download leaves its error standing beside every later successful
    // one. A user reading "could not be produced" next to a file that just
    // saved has no way to tell which claim is current.
    const driver = await ranOverBudget();
    const rows = driver.index.value?.rows ?? [];

    await driver.downloadRow('no-such-row-id');
    driver.errorMessage.value = 'That redline could not be produced.';

    await driver.downloadRow(rows[0]?.id as string);

    expect(clicked).toHaveLength(1);
    expect(driver.errorMessage.value).toBeNull();
  });

  it('says a failed pair has no redline, rather than that one could not be produced', async () => {
    // The two are different claims. "Could not be produced" describes a
    // packaging failure over a comparison that worked; this pair never
    // compared at all, and its row already says why.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([new File([new Uint8Array([1, 2, 3])], 'broken.docx')]);
    await driver.runBatch();

    const row = driver.index.value?.rows[0];
    expect(row?.status).toBe('failed');

    await driver.downloadRow(row?.id as string);

    expect(clicked).toHaveLength(0);
    expect(driver.errorMessage.value).toContain('broken.docx');
    expect(driver.errorMessage.value).toContain('did not complete');
    expect(driver.errorMessage.value).not.toMatch(/could not be produced/);
  });

  it('refuses to strand a compared row by removing its file', async () => {
    const driver = await ranOverBudget();
    const rows = driver.index.value?.rows ?? [];
    const id = rows[0]?.id as string;

    driver.removeVariant(id);

    // Still present, still re-runnable — and the refusal is stated, not silent.
    expect(driver.errorMessage.value).toMatch(/already been compared/i);
    expect(await driver.openPair(id)).not.toBeNull();
  });
});

/**
 * ADR-0042 at the drill-down seam.
 *
 * `comparePair`'s `side()` is the driver's OWN preprocess — the runner's records
 * are gone by the time a row is opened, so "Open changes" and any on-demand
 * redline re-run the pair from the files still held here. That makes it a fourth
 * numbering seam, and deleting its `buildNumberingContext(...)` left the whole
 * suite green while making the drill-down disagree with the row it drilled into:
 * the row says one change, the change list shows none.
 *
 * The redline BYTES cannot see this. A property op authors exactly as `equal`
 * (ADR-0042), so the `.docx` is byte-identical either way — which is the point
 * of the design and the reason this has to be asserted on `openPair`.
 */
describe('ADR-0042 — the numbering context reaches the drill-down', () => {
  async function ranPair(template: readonly StyledPara[], variant: readonly StyledPara[]) {
    const driver = createBatchDriver();
    driver.setTemplate(styleNumberedFile('house-spa.docx', template));
    driver.addVariants([styleNumberedFile('acme.docx', variant)]);
    await driver.runBatch();

    const row = driver.index.value?.rows[0];
    expect(row?.status).toBe('done');
    return { row: row!, detail: await driver.openPair(row!.id) };
  }

  it('opens a style-inherited re-levelling as a NUMBERING change, not a bare style one', async () => {
    const { row, detail } = await ranPair([HEAD, SUB_A], [HEAD, ['Clause1', SUB_A[1]]]);

    expect(row.changeCount).toBe(1);
    expect(detail?.summary.total).toBe(1);
    const property = detail?.ops.filter(o => o.type === 'property') ?? [];
    expect(property).toHaveLength(1);

    // The `w:pStyle` half needs no numbering context — the style ids differ on
    // their own, so a count of 1 passes with the wiring deleted. Only the
    // numbering half can see it: with no context every paragraph resolves to
    // `numbering: null`, the identity compares equal and both labels are null.
    const change = property[0]?.propertyChange;
    expect(change?.numberingIdentity).toBe(true);
    expect(change?.from.numberLabel).toBe('(a)');
    expect(change?.to.numberLabel).toBe('2.');
  });

  it('agrees with the row it drilled into on a pure renumbering', async () => {
    // Two sub-clauses shifted by one insertion. The RUNNER counts this correctly
    // (it has its own context); the drill-down re-runs the pair through `side()`.
    // Drop the context there and the two numbers stop matching — one product,
    // two answers, on one screen.
    const { row, detail } = await ranPair([HEAD, SUB_A, SUB_B], [HEAD, INSERTED, SUB_A, SUB_B]);

    expect(detail?.summary.total).toBe(row.changeCount);
    expect(detail?.summary.byCategory.crossRefRenumber).toBe(2);
    expect(detail?.summary.byCategory.formatting).toBe(0);
  });
});
