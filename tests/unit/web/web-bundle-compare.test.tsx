/**
 * The many-to-many web surface (ADR-0037; spec §6; acceptance M7).
 *
 * The hazard this file exists for is recorded in `batch/ui-types.ts` and has
 * bitten the project once already: a panel rendered against its stub driver
 * silently does nothing, and unit tests that merely render it do not notice.
 * Here that would be worse than a dead button — the stub's `confirm()` is a
 * no-op, so a surface wired to the stub would look like a working gate that
 * simply never opens.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { render } from 'preact';
import { act } from 'preact/test-utils';

import { WebBundleCompare } from '../../../src/web/WebBundleCompare.js';
import { createBundleDriver } from '@features/compare/batch/bundle-driver.js';
import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';

const SPA = [
  'This Share Purchase Agreement is made between the Seller and the Buyer.',
  'The Seller shall sell and the Buyer shall purchase the Sale Shares.',
  'The Purchase Price shall be paid in full on the Completion Date.',
  'The Seller warrants that it has good title to the Sale Shares.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

function docxFile(name: string, text: string): File {
  return new File([new Uint8Array(synthesiseDocx(text).toUint8Array())], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  render(null, container);
  container.remove();
  vi.restoreAllMocks();
});

/** Let the run drain. jsdom has no Worker, so it runs in-process on microtasks. */
async function settle(driver: { status: { value: string } }): Promise<void> {
  for (let i = 0; i < 200 && driver.status.value === 'running'; i += 1) {
    await act(async () => { await new Promise(r => setTimeout(r, 0)); });
  }
}

const runBtn = (): HTMLButtonElement | undefined =>
  Array.from(container.querySelectorAll<HTMLButtonElement>('button'))
    .find(b => b.textContent === 'Run comparisons');

async function loadedSurface() {
  const driver = createBundleDriver();
  await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

  driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
  driver.addDocuments('b', [
    docxFile('SPA v7.docx', SPA.replace('in full on the Completion Date', 'in cleared funds on Completion')),
  ]);
  await act(async () => { await driver.proposePairs(); });
  await act(async () => { render(<WebBundleCompare driver={driver} />, container); });
  return driver;
}

describe('the bundle surface is wired to a live driver, not the stub', () => {
  it('drives the injected driver when documents are chosen', async () => {
    const driver = createBundleDriver();
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    const inputs = container.querySelectorAll<HTMLInputElement>('input[type="file"]');
    expect(inputs).toHaveLength(2);

    // The stub has no addDocuments at all, so state changing is proof of life.
    driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });
    expect(container.textContent).toContain('SPA v6.docx');
  });

  it('creates its own live driver when none is passed', async () => {
    // The default path is the one WebApp uses. Driving it through the real
    // file input is the only way to tell a live driver from a stub: a stub
    // would render identically and do nothing.
    await act(async () => { render(<WebBundleCompare />, container); });

    const inputs = container.querySelectorAll<HTMLInputElement>('input[type="file"]');
    const chosen = [docxFile('SPA v6.docx', SPA)];
    Object.defineProperty(inputs[0]!, 'files', { value: chosen, configurable: true });
    await act(async () => { inputs[0]!.dispatchEvent(new Event('change', { bubbles: true })); });

    expect(container.textContent).toContain('SPA v6.docx');
    expect(container.textContent).toContain('1 of ');
  });
});

describe('M7 — the run button is gated on the confirmation, and so is the driver', () => {
  it('offers no run button until a pairing has been proposed', async () => {
    const driver = createBundleDriver();
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });
    expect(runBtn()).toBeUndefined();
  });

  it('disables the run button while the pairing is unconfirmed', async () => {
    await loadedSurface();
    expect(runBtn()).toBeTruthy();
    expect(runBtn()?.disabled).toBe(true);
  });

  it('enables it only after Confirm, and then it runs', async () => {
    const driver = await loadedSurface();

    await act(async () => {
      (container.querySelector('.dp-pairing__confirm') as HTMLButtonElement).click();
    });
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    expect(driver.confirmed.value).toBe(true);
    expect(runBtn()?.disabled).toBe(false);

    await act(async () => { runBtn()?.click(); });
    await settle(driver);
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    expect(driver.index.value?.mode).toBe('many-to-many');
    expect(driver.rows.value).toHaveLength(1);
  });

  it('refuses at the driver even if the button is bypassed', async () => {
    // The button is the courtesy; this is the invariant. A future component
    // that forgot the `disabled` attribute still cannot run an unconfirmed
    // pairing.
    const driver = await loadedSurface();
    expect(driver.confirmed.value).toBe(false);

    await driver.runConfirmed();

    expect(driver.index.value).toBeNull();
    expect(driver.errorMessage.value).toMatch(/confirm the pairing/i);
  });

  it('renders the grid with its rows rather than a bare button', async () => {
    await loadedSurface();
    expect(container.querySelector('.dp-pairing')).toBeTruthy();
    expect(container.querySelectorAll('.dp-pairing__row').length).toBeGreaterThan(0);
    expect(container.textContent).toMatch(/proposed, not decided/i);
  });
});

describe('I2 — the bundle surface cannot reach the Word add-in', () => {
  const ROOT = resolve(__dirname, '../../..');

  function taskpaneSources(dir: string): string[] {
    const out: string[] = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = resolve(dir, entry.name);
      if (entry.isDirectory()) out.push(...taskpaneSources(full));
      else if (/\.tsx?$/.test(entry.name)) out.push(full);
    }
    return out;
  }

  it('is referenced nowhere under src/taskpane', () => {
    for (const file of taskpaneSources(resolve(ROOT, 'src/taskpane'))) {
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');
      expect(code).not.toMatch(/WebBundleCompare|createBundleDriver|PairingGrid/);
    }
  });

  it('is host-free and network-free', () => {
    for (const rel of [
      'src/web/WebBundleCompare.tsx',
      'src/features/compare/batch/bundle-driver.ts',
      'src/features/compare/batch/components/PairingGrid.tsx',
    ]) {
      const code = readFileSync(resolve(ROOT, rel), 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/[^\n]*/g, '');
      expect(/\bOffice\./.test(code)).toBe(false);
      expect(/\bWord\./.test(code)).toBe(false);
      expect(/\bfetch\s*\(/.test(code)).toBe(false);
      expect(/\bnew\s+(XMLHttpRequest|WebSocket|EventSource)\b/.test(code)).toBe(false);
    }
  });
});

describe('opening a row shows its changes without saving a file', () => {
  it('renders the change list for the pair that was opened', async () => {
    const driver = await loadedSurface();

    await act(async () => {
      (container.querySelector('.dp-pairing__confirm') as HTMLButtonElement).click();
    });
    await act(async () => { runBtn()?.click(); });
    await settle(driver);
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    const openBtn = container.querySelector('.dp-batch__open') as HTMLButtonElement;
    expect(openBtn).toBeTruthy();

    await act(async () => { openBtn.click(); });
    for (let i = 0; i < 8; i += 1) {
      await act(async () => { await new Promise(r => setTimeout(r, 0)); });
    }
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    const drilldown = container.querySelector('.dp-batch__drilldown');
    expect(drilldown).toBeTruthy();
    expect(drilldown?.textContent).toContain('▸');
  });

  it('offers no "Open changes" button before anything has run', async () => {
    await loadedSurface();
    expect(container.querySelector('.dp-batch__open')).toBeNull();
  });
});

describe('the bundle surface reports a run the way §5 requires', () => {
  /** Confirm, run, drain, re-render. */
  async function ranSurface() {
    const driver = await loadedSurface();
    await act(async () => {
      (container.querySelector('.dp-pairing__confirm') as HTMLButtonElement).click();
    });
    await act(async () => { runBtn()?.click(); });
    await settle(driver);
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });
    return driver;
  }

  it('shows each row\'s classification chips and a readable status', async () => {
    // §5.4. The many-to-many surface hand-rolled its own row line instead of
    // reusing the one-to-many one, and the copy drifted: no chips, and the raw
    // status enum where a label belongs. A reader scanning a hundred redlines
    // is looking for "which of these changed an obligation", and the chips are
    // the only thing on the row that answers it.
    await ranSurface();

    expect(container.querySelector('.dp-batch__row-chips')).toBeTruthy();
    expect(container.querySelector('.dp-batch__row-status')?.textContent).toBe('Compared');
  });

  it('renders each batch disclosure once', async () => {
    // Two blocks rendered the same list — the second surface growing a second
    // copy of the same model, one component along from the row line that
    // occasioned all this. A disclosure printed twice reads as two findings.
    const driver = createBundleDriver({ artefactBudgetBytes: 1 });
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });
    driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
    driver.addDocuments('b', [docxFile('SPA v7.docx',
      SPA.replace('in full on the Completion Date', 'in cleared funds on Completion'))]);
    await act(async () => { await driver.proposePairs(); });
    driver.confirm();
    await act(async () => { await driver.runConfirmed(); });
    await settle(driver);
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    const disclosures = driver.index.value?.disclosures ?? [];
    expect(disclosures.length).toBeGreaterThan(0);
    const rendered = Array.from(container.querySelectorAll('.dp-batch__disclosure'))
      .map(p => p.textContent ?? '');
    for (const text of disclosures) {
      expect(rendered.filter(t => t === text)).toHaveLength(1);
    }
  });

  it('states the flatten choice its redlines were produced under', async () => {
    // §8 case 7 / §10. It changes what every redline in the run says, and the
    // index has carried it all along — it was simply never rendered here.
    await ranSurface();

    const note = container.querySelector('.dp-batch__note')?.textContent ?? '';
    expect(note).toMatch(/Tracked changes accepted first/);
    // Named for the mode: a bundle's side A is not a template.
    expect(note).toMatch(/side-A document/);
    expect(note).not.toMatch(/the template/);
    const limits = Array.from(container.querySelectorAll('.dp-batch__limits li'))
      .map(li => li.textContent ?? '');
    expect(limits.length).toBeGreaterThan(0);
    // §10's one pairing-specific limitation belongs on the surface that pairs.
    expect(limits.some(t => /lexical and structural/.test(t))).toBe(true);
  });

  it('says a batch was cancelled, and that what it finished is kept', async () => {
    // §5.5. Under a progress bar that reads as complete, the rows' own
    // Cancelled chips are easy to miss.
    const driver = await loadedSurface();
    driver.addDocuments('a', [docxFile('Tax Deed v2.docx', SPA.replace('Seller', 'Covenantor'))]);
    driver.addDocuments('b', [docxFile('Tax Deed (execution).docx',
      SPA.replace('Seller', 'Covenantor').replace('in full', 'in cleared funds'))]);
    await act(async () => { await driver.proposePairs(); });
    driver.confirm();

    const stop = driver.rows.subscribe(() => {
      if (driver.rows.value.some(r => r.status === 'done')) driver.cancelRun();
    });
    await act(async () => { await driver.runConfirmed(); });
    stop();
    await act(async () => { render(<WebBundleCompare driver={driver} />, container); });

    const notices = Array.from(container.querySelectorAll('.dp-batch__disclosure'))
      .map(p => p.textContent ?? '');
    expect(notices.some(t => /Batch cancelled after \d+ of \d+/.test(t))).toBe(true);
    expect(notices.some(t => /already produced are kept/.test(t))).toBe(true);
  });
});
