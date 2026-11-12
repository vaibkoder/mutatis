/**
 * Compare at scale — the web batch surface (ADR-0037; spec §5.4; acceptance I2).
 *
 * Three things this file exists to prove, in descending order of how quietly
 * they would otherwise fail:
 *
 *  1. **The live driver is injected.** `BatchComparePanel` defaults to a no-op
 *     stub, so a surface that forgets to inject renders perfectly and does
 *     nothing at all. That bug has shipped before,
 *     and a rendering test does
 *     not catch it — hence an assertion aimed squarely at it.
 *  2. **Opening a row re-runs that pair** rather than materialising a redline,
 *     which on web means triggering a download.
 *  3. **The add-in is untouched** — the batch surface is not reachable from
 *     `taskpane/**`.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { setActiveHost, _resetActiveHost } from '@shared/host/index.js';
import { WebHostPort } from '@host-web/index.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import { createBatchDriver } from '@features/compare/batch/index.js';
import { WebBatchCompare } from '../../../src/web/WebBatchCompare.js';
import { WebApp } from '../../../src/web/WebApp.js';
import { BatchComparePanel } from '@features/compare/batch/components/BatchComparePanel.js';
import { makeStubBatchDriverState } from '@features/compare/batch/ui-types.js';
import { signal } from '@preact/signals';

const TEMPLATE_TEXT = [
  'The Seller shall pay the Purchase Price on Completion.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

function docxFile(name: string, text: string): File {
  const bytes = new Uint8Array(synthesiseDocx(text).toUint8Array());
  return new File([bytes], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/**
 * Let pending microtasks and timers settle inside `act`.
 *
 * `openPair` reads two files and runs the pipeline — several async hops — so a
 * bare `act(async () => {})` returns before the drill-down has rendered.
 */
async function settle(times = 5): Promise<void> {
  for (let i = 0; i < times; i += 1) {
    await act(async () => { await new Promise(r => setTimeout(r, 0)); });
  }
}

let container: HTMLDivElement;

beforeEach(() => {
  setActiveHost(new WebHostPort());
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  render(null, container);
  container.remove();
  _resetActiveHost();
  vi.restoreAllMocks();
});

describe('the batch surface is wired to a live driver, not the stub', () => {
  it('drives the injected driver when a template is chosen', async () => {
    const driver = createBatchDriver();
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    const inputs = container.querySelectorAll('input[type="file"]');
    expect(inputs).toHaveLength(2);

    // The stub's setTemplate is a no-op, so a template landing in driver state
    // is proof the live one is in play.
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    await act(async () => {});

    expect(driver.template.value?.label).toBe('house-nda.docx');
    expect(container.textContent).toContain('house-nda.docx');
  });

  it('WebBatchCompare creates its own driver when none is passed', async () => {
    await act(async () => { render(<WebBatchCompare />, container); });
    // It renders its own live driver rather than falling through to the stub;
    // the run button is present and correctly disabled with nothing chosen.
    const run = container.querySelector('.dp-batch__run') as HTMLButtonElement;
    expect(run).toBeTruthy();
    expect(run.disabled).toBe(true);
  });

  it('keeps one driver across re-renders instead of dropping the queue', async () => {
    const driver = createBatchDriver();
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT)]);
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    expect(driver.variants.value).toHaveLength(1);
    expect(container.textContent).toContain('a.docx');
  });
});

describe('the panel reports every pair, including the ones that went wrong', () => {
  it('renders a row per variant with its status and change count', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
      docxFile('borealis.docx', TEMPLATE_TEXT),
      new File([new Uint8Array([1, 2, 3])], 'broken.docx'),
    ]);

    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });

    const rows = container.querySelectorAll('.dp-batch__row');
    expect(rows).toHaveLength(3);

    const text = container.textContent ?? '';
    expect(text).toContain('acme.docx');
    expect(text).toContain('borealis.docx');
    // The failure is shown with its reason, not omitted.
    expect(text).toContain('broken.docx');
    expect(text).toContain('Could not compare');
    // And the unchanged variant says so rather than looking like a failure.
    expect(text).toContain('No changes');
  });

  it('offers a download only for rows that actually have a redline retained', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five days.')),
      new File([new Uint8Array([1, 2, 3])], 'broken.docx'),
    ]);

    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });

    const downloads = Array.from(container.querySelectorAll('button'))
      .filter(b => b.textContent === 'Download redline');
    expect(downloads).toHaveLength(1);
  });
});

describe('drill-down re-runs one pair rather than materialising a redline', () => {
  it('shows the change list without triggering a download', async () => {
    const clicks = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});

    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five Business Days.')),
    ]);

    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });

    const open = Array.from(container.querySelectorAll('button'))
      .find(b => b.textContent === 'Open changes') as HTMLButtonElement;
    expect(open).toBeTruthy();

    await act(async () => { open.click(); });
    await settle();

    expect(container.querySelector('.dp-batch__drilldown')).toBeTruthy();
    // The whole point: showing what changed must not save a file.
    expect(clicks).not.toHaveBeenCalled();
  });

  it('recovers the ops the batch dropped', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([
      docxFile('acme.docx', TEMPLATE_TEXT.replace('on Completion.', 'within five Business Days.')),
    ]);
    await act(async () => { await driver.runBatch(); });

    const rowId = driver.rows.value[0]?.id as string;
    const detail = await driver.openPair(rowId);

    expect(detail?.ops.some(op => op.type !== 'equal')).toBe(true);
    expect(detail?.summary.total).toBe(driver.rows.value[0]?.changeCount);
  });

  it('declines to open a row that was never compared', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([new File([new Uint8Array([1, 2, 3])], 'broken.docx')]);
    await act(async () => { await driver.runBatch(); });

    const rowId = driver.rows.value[0]?.id as string;
    expect(await driver.openPair(rowId)).toBeNull();
  });
});

describe('the web app exposes the batch view', () => {
  it('renders a "Compare many" tab that mounts the batch surface', async () => {
    await act(async () => { render(<WebApp host={new WebHostPort()} />, container); });

    const tab = Array.from(container.querySelectorAll('[role="tab"]'))
      .find(b => b.textContent === 'Compare many') as HTMLButtonElement;
    expect(tab).toBeTruthy();

    await act(async () => { tab.click(); });

    expect(container.querySelector('.dp-batch')).toBeTruthy();
  });
});

/**
 * A template/variant pair that lands changes in THREE different classes.
 *
 * The chip test used to run on a one-change, one-class fixture, and a review
 * showed it could not fail: deleting `structural` from the chip list outright
 * left all 26 tests green, because the fixture never produced a structural
 * change to drop. A wall that only holds for the one class you happened to
 * exercise is not a wall. These three lines produce, in order, a numerical
 * change, an obligation-language change, and a structural insertion.
 */
const MULTI_CLASS_TEMPLATE = [
  'The Seller shall pay the Purchase Price within 30 days of Completion.',
  'The Buyer may inspect the Property on reasonable notice.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

const MULTI_CLASS_VARIANT = [
  'The Seller shall pay the Purchase Price within 60 days of Completion.',
  'The Buyer shall inspect the Property on reasonable notice.',
  'This Agreement is governed by the laws of England and Wales.',
  'Time is of the essence in respect of the payment obligation.',
].join('\n');

describe('classification chips (spec §5.4)', () => {
  it('renders a chip per non-zero class, and they sum to the row\'s change count', async () => {
    // The panel's own load-bearing claim. A chip list that silently dropped a
    // class would show fewer changes than the row counts — the ADR-0008 wall
    // failing quietly at the one place a reader actually looks.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', MULTI_CLASS_TEMPLATE));
    driver.addVariants([docxFile('acme.docx', MULTI_CLASS_VARIANT)]);
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    const row = driver.index.value?.rows[0];
    expect(row?.status).toBe('done');
    expect(row?.changeCount).toBeGreaterThan(0);

    const chips = Array.from(container.querySelectorAll('.dp-batch__chip'));
    expect(chips.length).toBeGreaterThan(0);

    const shown = chips.reduce(
      (n, c) => n + Number(/^\s*(\d+)/.exec(c.textContent ?? '')?.[1] ?? 0),
      0,
    );
    expect(shown).toBe(row?.changeCount);

    // One chip per non-zero class — no class is rendered twice or omitted.
    const nonZero = Object.entries(row?.byClassification ?? {}).filter(([, v]) => v > 0);
    expect(chips).toHaveLength(nonZero.length);

    // And the fixture really does exercise more than one bucket, so dropping
    // ANY of these three from the chip list fails this test rather than only
    // the one class a narrower fixture happened to produce.
    expect(nonZero.map(([k]) => k).sort()).toEqual(
      ['numericalValue', 'obligationLanguage', 'structural'],
    );
  });

  it('renders no chips for a variant identical to the template', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([docxFile('same.docx', TEMPLATE_TEXT)]);
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    expect(driver.index.value?.rows[0]?.changeCount).toBe(0);
    expect(container.querySelectorAll('.dp-batch__chip')).toHaveLength(0);
    expect(container.textContent).toContain('No changes');
  });

  it('states the flatten choice and the limitations once a batch has finished', async () => {
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT.replace('Completion.', 'Friday.'))]);
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });
    await act(async () => { await driver.runBatch(); });
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    // Spec §8 case 7 and spec §10, both stated in-product.
    expect(container.textContent).toMatch(/Tracked changes accepted first/i);
    expect(container.textContent).toMatch(/list of comparisons, not an analysis across them/i);
  });

  it('locks the file inputs once a batch has produced rows', async () => {
    // The rows were produced against THIS template and THESE versions; letting
    // either change would make every later re-run compare something else.
    const driver = createBatchDriver();
    driver.setTemplate(docxFile('house-nda.docx', TEMPLATE_TEXT));
    driver.addVariants([docxFile('a.docx', TEMPLATE_TEXT.replace('Completion.', 'Friday.'))]);
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    const inputsBefore = container.querySelectorAll<HTMLInputElement>('input[type="file"]');
    expect(Array.from(inputsBefore).every(i => !i.disabled)).toBe(true);

    await act(async () => { await driver.runBatch(); });
    await act(async () => { render(<WebBatchCompare driver={driver} />, container); });

    const inputsAfter = container.querySelectorAll<HTMLInputElement>('input[type="file"]');
    expect(Array.from(inputsAfter).every(i => i.disabled)).toBe(true);
  });
});

describe('I2 — the Word add-in is untouched', () => {
  /** Every `.ts`/`.tsx` under `src/taskpane/`, not just the shell. */
  function taskpaneSources(dir: string): string[] {
    const out: string[] = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = resolve(dir, entry.name);
      if (entry.isDirectory()) out.push(...taskpaneSources(full));
      else if (/\.tsx?$/.test(entry.name)) out.push(full);
    }
    return out;
  }

  const ROOT = resolve(__dirname, '../../..');
  const files = taskpaneSources(resolve(ROOT, 'src/taskpane'));

  it('covers the whole task-pane tree, not just the shell', () => {
    // ADR-0003 caps the add-in at two documents in memory, which is what makes
    // N-way batching structurally wrong there. Checking one file would let the
    // batch surface reach the add-in through any other.
    expect(files.length).toBeGreaterThanOrEqual(5);
    expect(files.some(f => f.endsWith('Shell.tsx'))).toBe(true);
  });

  for (const file of files) {
    const name = file.slice(ROOT.length + 1);
    it(`${name} cannot reach the batch surface`, () => {
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');

      expect(code).not.toMatch(/compare\/batch/);
      expect(code).not.toMatch(/BatchCompare|createBatchDriver|runOneToManyBatch/);
    });
  }
});

describe('the variant cap the panel shows is the one it was given (spec §7)', () => {
  it('renders the driver\'s resolved cap, not the built-in default', async () => {
    // The display half of the wiring tested in `batch.driver-config.test.ts`.
    // A panel that reached for `MAX_BATCH_VARIANTS` directly would render 100
    // beside a driver that refuses past 3, and the only place a user would
    // find out is a batch that fails at document four.
    const state = { ...makeStubBatchDriverState(signal), maxVariants: 3 };
    await act(async () => { render(<BatchComparePanel driver={state} />, container); });

    expect(container.textContent).toContain('0 of 3');
    expect(container.textContent).not.toContain('of 100');
  });
});
