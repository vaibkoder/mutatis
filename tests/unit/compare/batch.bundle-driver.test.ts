/**
 * The many-to-many driver (ADR-0037; spec §6; acceptance M7; 17b task 8).
 *
 * The driver is where the confirmation gate is actually enforced. A disabled
 * button is a courtesy; this is the invariant, so it is tested as one — the
 * runner must record zero invocations until `confirm()` has been called, and
 * a confirmation must not survive an edit to the documents it was made about.
 */
import { describe, it, expect, vi } from 'vitest';

import { synthesiseDocx } from '@features/compare/io/text/synthesise-docx.js';
import {
  styleNumberedFile,
  HEAD,
  SUB_A,
  SUB_B,
  INSERTED,
  type StyledPara,
} from '../../fixtures/style-numbered-docx.js';
import { createBundleDriver } from '@features/compare/batch/bundle-driver.js';
import * as writeout from '@features/compare/output/writeout.js';
import * as preprocess from '@features/compare/preprocess/index.js';
import * as pipeline from '@features/compare/pipeline.js';
import { MAX_BUNDLE_DOCUMENTS, MAX_DOCX_BYTES } from '@features/compare/batch/limits.js';

const SPA = [
  'This Share Purchase Agreement is made between the Seller and the Buyer.',
  'The Seller shall sell and the Buyer shall purchase the Sale Shares.',
  'The Purchase Price shall be paid in full on the Completion Date.',
  'The Seller warrants that it has good title to the Sale Shares.',
  'This Agreement is governed by the laws of England and Wales.',
].join('\n');

const DEED = [
  'This Deed of Tax Covenant is made between the Covenantor and the Covenantee.',
  'The Covenantor covenants to pay an amount equal to any Tax Liability.',
  'A claim under this Deed must be notified within seven years.',
  'Payment shall be made in cleared funds within five Business Days.',
  'This Deed is governed by the laws of Singapore.',
].join('\n');

function docxFile(name: string, text: string): File {
  return new File([new Uint8Array(synthesiseDocx(text).toUint8Array())], name, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/** Two bundles of two documents each, correctly pairable. */
async function loaded() {
  const driver = createBundleDriver();
  driver.addDocuments('a', [
    docxFile('Project Falcon SPA v6.docx', SPA),
    docxFile('Project Falcon Tax Deed v2.docx', DEED),
  ]);
  driver.addDocuments('b', [
    docxFile('Project Falcon SPA v7.docx', SPA.replace('in full on the Completion Date', 'in cleared funds on Completion')),
    docxFile('Project Falcon Tax Deed (execution).docx', DEED.replace('seven years', 'five years')),
  ]);
  await driver.proposePairs();
  return driver;
}

describe('proposing a pairing', () => {
  it('sketches both bundles and proposes each document against its counterpart', async () => {
    const driver = await loaded();
    const proposal = driver.proposal.value;

    expect(proposal).not.toBeNull();
    expect(proposal?.totals.pairs).toBe(2);
    expect(proposal?.totals.unmatchedA).toBe(0);
    expect(proposal?.totals.unmatchedB).toBe(0);
    for (const row of proposal?.rows ?? []) expect(row.label).toMatch(/ ▸ /);
  });

  it('refuses with a message when a side is empty, rather than proposing nothing', async () => {
    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('a.docx', SPA)]);
    await driver.proposePairs();

    expect(driver.proposal.value).toBeNull();
    expect(driver.errorMessage.value).toMatch(/both sides/i);
  });

  it('surfaces a cap breach as a message rather than a partial pairing', async () => {
    const driver = createBundleDriver();
    driver.addDocuments('a', Array.from(
      { length: MAX_BUNDLE_DOCUMENTS + 1 },
      (_, i) => docxFile(`a${i}.docx`, SPA),
    ));
    driver.addDocuments('b', [docxFile('b.docx', SPA)]);
    await driver.proposePairs();

    expect(driver.proposal.value).toBeNull();
    expect(driver.errorMessage.value).toMatch(/cap/i);
  });

  it('exposes the cap it will actually enforce', () => {
    expect(createBundleDriver().maxBundleDocuments).toBe(MAX_BUNDLE_DOCUMENTS);
  });
});

describe('M7 — nothing runs before the pairing is confirmed', () => {
  it('refuses to run an unconfirmed pairing, with a reason', async () => {
    const driver = await loaded();
    expect(driver.confirmed.value).toBe(false);

    await driver.runConfirmed();

    // Not merely "the button was disabled" — the driver itself refused.
    expect(driver.rows.value).toEqual([]);
    expect(driver.index.value).toBeNull();
    expect(driver.status.value).toBe('idle');
    expect(driver.errorMessage.value).toMatch(/confirm the pairing/i);
  });

  it('runs once confirmed, and produces one redline per pair', async () => {
    const driver = await loaded();
    driver.confirm();
    expect(driver.confirmed.value).toBe(true);

    await driver.runConfirmed();

    expect(driver.status.value).toBe('done');
    expect(driver.index.value?.mode).toBe('many-to-many');
    expect(driver.rows.value).toHaveLength(2);
    expect(driver.rows.value.every(r => r.status === 'done')).toBe(true);
    expect(driver.rows.value.every(r => r.artefacts[0]?.kind === 'redlineDocx')).toBe(true);
  });

  it('will not confirm when there is nothing to confirm', () => {
    const driver = createBundleDriver();
    driver.confirm();
    expect(driver.confirmed.value).toBe(false);
  });
});

describe('a confirmation is about a specific set of documents', () => {
  it('is discarded when a document is added', async () => {
    const driver = await loaded();
    driver.confirm();

    driver.addDocuments('b', [docxFile('Side Letter.docx', DEED)]);

    // A confirmation that survived the edit would let a run proceed on a
    // pairing the user confirmed for a different bundle.
    expect(driver.confirmed.value).toBe(false);
    expect(driver.proposal.value).toBeNull();
  });

  it('is discarded when a document is removed', async () => {
    const driver = await loaded();
    driver.confirm();

    driver.removeDocument('a', driver.sideA.value[0]?.id as string);

    expect(driver.confirmed.value).toBe(false);
    expect(driver.proposal.value).toBeNull();
  });

  it('leaves nothing behind when the bundles are cleared', async () => {
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();

    driver.clearBundles();

    expect(driver.sideA.value).toEqual([]);
    expect(driver.sideB.value).toEqual([]);
    expect(driver.rows.value).toEqual([]);
    expect(driver.index.value).toBeNull();
    expect(driver.proposal.value).toBeNull();
    expect(driver.confirmed.value).toBe(false);
    expect(driver.status.value).toBe('idle');
  });
});

describe('re-pointing a pair', () => {
  it('uses the user\'s choice verbatim instead of re-solving', async () => {
    const driver = await loaded();
    const before = driver.proposal.value;
    const spaRow = before?.rows.find(r => r.label.startsWith('Project Falcon SPA v6')) ;
    const deedB = before?.documents.find(d => d.side === 'b' && d.fileName.includes('Tax Deed'));

    driver.repoint(spaRow?.id as string, deedB?.docId as string);

    const after = driver.proposal.value;
    const repointed = after?.rows.find(r => r.a === spaRow?.a);
    expect(repointed?.b).toBe(deedB?.docId);

    // A solver that "improved" this back would be overriding the user.
    expect(after?.rows.filter(r => r.b === deedB?.docId)).toHaveLength(1);
  });

  it('frees the document the other row was holding rather than showing it twice', async () => {
    // Pairing is one-to-one. A local edit that ignored that would show one
    // document twice and another not at all, with every row still plausible.
    const driver = await loaded();
    const before = driver.proposal.value;
    const spaRow = before?.rows.find(r => r.label.startsWith('Project Falcon SPA v6'));
    const deedB = before?.documents.find(d => d.side === 'b' && d.fileName.includes('Tax Deed'));

    driver.repoint(spaRow?.id as string, deedB?.docId as string);
    const after = driver.proposal.value;

    const bIds = (after?.rows ?? []).map(r => r.b).filter((b): b is string => b !== null);
    expect(new Set(bIds).size).toBe(bIds.length);
    // Every document still appears exactly once across the grid.
    const seen = [
      ...(after?.rows ?? []).map(r => r.a).filter(Boolean),
      ...(after?.rows ?? []).map(r => r.b).filter(Boolean),
    ];
    expect(new Set(seen).size).toBe(4);
  });

  it('can mark a pair unmatched', async () => {
    const driver = await loaded();
    const row = driver.proposal.value?.rows[0];

    driver.repoint(row?.id as string, null);

    const after = driver.proposal.value;
    expect(after?.totals.pairs).toBe(1);
    expect(after?.totals.unmatchedA).toBe(1);
    expect(after?.totals.unmatchedB).toBe(1);
  });

  it('offers every side-B document as a candidate, best first', async () => {
    const driver = await loaded();
    const row = driver.proposal.value?.rows[0];
    const candidates = driver.candidatesFor(row?.id as string);

    expect(candidates).toHaveLength(2);
    expect(candidates[0]?.score).toBeGreaterThanOrEqual(candidates[1]?.score as number);
  });

  it('ignores a re-point once the pairing is confirmed', async () => {
    const driver = await loaded();
    const before = driver.proposal.value;
    driver.confirm();

    driver.repoint(before?.rows[0]?.id as string, null);

    expect(driver.proposal.value).toBe(before);
    expect(driver.confirmed.value).toBe(true);
  });
});

describe('the run itself', () => {
  it('reports progress and finishes at 100', async () => {
    const driver = await loaded();
    driver.confirm();

    const seen: number[] = [];
    const unsubscribe = driver.progress.subscribe(v => seen.push(v));
    await driver.runConfirmed();
    unsubscribe();

    expect(driver.progress.value).toBe(100);
    expect(driver.index.value?.totals.done).toBe(2);
    // It moved, rather than jumping straight from 0 to 100 at the end.
    expect(new Set(seen).size).toBeGreaterThan(1);
  });

  it('keeps completed rows when the run is cancelled, and says how far it got', async () => {
    const driver = await loaded();
    driver.confirm();

    // Cancel the moment the first pair lands, deterministically.
    let stopped = false;
    const stopWhenFirstCompletes = (): void => {
      if (stopped) return;
      if (driver.rows.value.some(r => r.status === 'done')) {
        stopped = true;
        driver.cancelRun();
      }
    };
    const unsubscribe = driver.rows.subscribe(stopWhenFirstCompletes);
    await driver.runConfirmed();
    unsubscribe();

    const index = driver.index.value;
    expect(index?.totals.pairs).toBe(2);
    expect(index?.totals.done).toBeGreaterThanOrEqual(1);
    // Every pair is still a row, and completed ones keep their redline.
    expect(index?.rows).toHaveLength(2);
    for (const row of index?.rows.filter(r => r.status === 'done') ?? []) {
      expect(row.artefacts[0]?.byteLength).toBeGreaterThan(0);
    }
    // The bar shows how far it actually got, not a misleading 100.
    if ((index?.totals.cancelled ?? 0) > 0) expect(driver.progress.value).toBeLessThan(100);
  });

  it('refuses to re-open the pairing mid-run, and finishes describing the one it ran', async () => {
    // A review found `unconfirm()` reachable during a run: the grid's "change
    // the pairing" button had no idea a batch was in flight. Taking it cleared
    // `confirmed`, which re-enabled `repoint()`, which called `discardResults()`
    // — and then `runConfirmed`'s own continuation wrote the OLD pairing's rows
    // and index back over the cleared state. The screen ended with a complete,
    // downloadable redline sitting under a grid saying that pairing did not
    // exist. A run consumes the pairing it was given; cancel is the way out.
    const driver = await loaded();
    const before = driver.proposal.value?.rows.map(r => r.id);
    driver.confirm();

    let tried = false;
    let stillConfirmed = false;
    const meddle = (): void => {
      if (tried) return;
      if (!driver.rows.value.some(r => r.status === 'done')) return;
      tried = true;
      driver.unconfirm();
      // Captured HERE: the later calls write `confirmed` by hand, so a check
      // at the end of the test would be satisfied by those instead and this
      // guard would have no test at all.
      stillConfirmed = driver.confirmed.value;
      driver.repoint(driver.proposal.value?.rows[0]?.id as string, null);
      // `confirmed` is a public writable signal, so `unconfirm()` is not the
      // only way to it. Forcing it directly is what a future surface would do
      // by accident; `repoint` must still refuse.
      driver.confirmed.value = false;
      driver.repoint(driver.proposal.value?.rows[0]?.id as string, null);
      driver.confirmed.value = true;
      // And every other path into `invalidate()`. Each one ends in
      // `discardResults()`, which resets `status` — so a mid-run edit does not
      // merely clear the screen, it makes `running` lie and lets a second run
      // start alongside the first.
      driver.addDocuments('b', [docxFile('Side Letter.docx', DEED)]);
      driver.removeDocument('a', driver.sideA.value[0]?.id as string);
      driver.clearBundles();
      void driver.proposePairs();
      // And the sixth writer of `status`: `runConfirmed`'s own flatten-drift
      // branch calls `invalidate()`, so re-entering it mid-run used to clear
      // the proposal and make `running` say false while the run continued.
      driver.flattenA.value = 'rejectAll';
      void driver.runConfirmed();
      driver.flattenA.value = 'acceptAll';
    };
    const unsubscribe = driver.rows.subscribe(meddle);
    await driver.runConfirmed();
    unsubscribe();

    expect(tried).toBe(true);
    expect(stillConfirmed).toBe(true);
    expect(driver.confirmed.value).toBe(true);
    expect(driver.errorMessage.value).toMatch(/cancel the run/i);
    expect(driver.proposal.value?.rows.map(r => r.id)).toEqual(before);
    expect(driver.sideA.value).toHaveLength(2);
    expect(driver.sideB.value).toHaveLength(2);
    // And the run still describes the pairing it was given, in full.
    expect(driver.index.value?.rows).toHaveLength(2);
    expect(driver.index.value?.rows.map(r => r.label))
      .toEqual(driver.rows.value.map(r => r.label));
  });

  it('will not compare a pair the user has not confirmed', async () => {
    // Spec §6.1: "nothing is compared before the user confirms the pairing."
    // `runConfirmed` is not the only public method that starts one — `openPair`
    // is the other, and its guard checked `runnable` and not the gate. A caller
    // holding a proposal could get a complete, plausible change list for a
    // pairing no human had read, which is exactly the artefact this feature's
    // confirmation step exists to prevent.
    const driver = await loaded();
    const spy = vi.spyOn(pipeline, 'runComparePipeline');

    expect(driver.confirmed.value).toBe(false);
    const detail = await driver.openPair(driver.proposal.value?.rows[0]?.id as string);

    expect(detail).toBeNull();
    expect(spy).not.toHaveBeenCalled();

    // And the control: confirmed and run, the very same call works.
    driver.confirm();
    await driver.runConfirmed();
    spy.mockClear();
    expect(await driver.openPair(driver.rows.value[0]?.id as string)).not.toBeNull();
    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });

  it('gives every redline in one batch the same date, including one opened mid-run', async () => {
    // `runner.ts` injects the date so every redline in a batch shares one, and
    // `index.value` is null for the whole DURATION of a run — so a re-run that
    // read it fell through to `new Date()` and stamped a redline produced
    // mid-batch differently from every other one in the same batch. The rows
    // are rendered while the batch runs, and each carries its own buttons.
    const driver = await loaded();
    driver.confirm();
    const spy = vi.spyOn(pipeline, 'runComparePipeline');

    let midRun: Promise<unknown> | null = null;
    const stop = driver.rows.subscribe(() => {
      if (midRun !== null) return;
      const done = driver.rows.value.find(r => r.status === 'done');
      if (done !== undefined) midRun = driver.openPair(done.id);
    });
    await driver.runConfirmed();
    stop();
    await midRun;

    const batchDate = driver.index.value?.generatedAt as string;
    expect(midRun).not.toBeNull();
    expect(spy.mock.calls.length).toBeGreaterThan(0);
    for (const call of spy.mock.calls) expect(call[0]?.date).toBe(batchDate);
    spy.mockRestore();
  });

  it('re-runs a pair under the flatten policy the index states, not the live one', async () => {
    // Spec §8 case 7. A drill-down and an on-demand redline must reproduce the
    // row they belong to. `flattenA`/`flattenB` are public writable signals, so
    // reading them at re-run time let a redline be assembled under a flatten
    // choice the index explicitly denies, and a change list disagree with its
    // own row's `changeCount`. The one-to-many driver pins the same three
    // values; here the flatten policy was the one still moving.
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();
    expect(driver.index.value?.flatten).toEqual({ template: 'acceptAll', variant: 'acceptAll' });

    driver.flattenA.value = 'rejectAll';
    driver.flattenB.value = 'rejectAll';

    const spy = vi.spyOn(preprocess, 'preprocessSide');
    await driver.openPair(driver.rows.value[0]?.id as string);

    expect(spy.mock.calls.length).toBeGreaterThan(0);
    expect(spy.mock.calls.map(c => c[2])).toEqual(spy.mock.calls.map(() => 'acceptAll'));
    spy.mockRestore();
  });

  it('refuses an edit while the bundles are being read', async () => {
    // `proposePairs` resolves its sketches against the CURRENT bundle lists, so
    // a document added while it runs would be missing from the grid it
    // produces — and "every document appears exactly once" (§6.6) is the one
    // property the grid cannot be wrong about. The surface disables its inputs
    // during that window; the driver is public API and must refuse anyway.
    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
    driver.addDocuments('b', [docxFile('SPA v7.docx',
      SPA.replace('in full on the Completion Date', 'in cleared funds on Completion'))]);

    const inFlight = driver.proposePairs();
    expect(driver.busy.value).toBe(true);
    driver.addDocuments('a', [docxFile('Side Letter.docx', DEED)]);
    driver.clearBundles();
    await inFlight;

    expect(driver.errorMessage.value).toMatch(/still being read/i);
    expect(driver.sideA.value).toHaveLength(1);
    expect(driver.proposal.value?.documents).toHaveLength(2);
  });

  it('keeps a batch\'s results when a re-point changes nothing', async () => {
    // `repointPairing` returns the proposal untouched when there is nothing to
    // re-point. Discarding a batch's redlines unconditionally throws away every
    // on-demand artefact for a request that changed nothing.
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();
    const before = driver.rows.value.length;
    driver.unconfirm();

    driver.repoint('a-does-not-exist::', null);
    expect(driver.rows.value).toHaveLength(before);

    // And re-pointing a row at the counterpart it already has. The proposal is
    // rebuilt for that one, so identity on the result does not catch it — the
    // check has to be on the value, and the cost of missing it is every
    // on-demand redline in the batch.
    const row = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    driver.repoint(row?.id as string, row?.b as string);

    expect(driver.rows.value).toHaveLength(before);
    expect(driver.index.value).not.toBeNull();
  });

  it('will not start a run while the bundles are still being read', async () => {
    // During that window `confirmed` is still true and `proposal` still holds
    // the PREVIOUS one, so a run started then is un-confirmed mid-flight by the
    // propose that follows it: a complete, downloadable batch under a grid
    // saying nothing was confirmed.
    //
    // The state is not reachable through the driver today — a re-propose that
    // reuses every sketch never awaits, so `busy` is never observable — but
    // that is an implementation detail of the reuse check, not a property, and
    // `busy` is a public signal. Set directly, which is what a propose that
    // gained an await would do on its own.
    const driver = await loaded();
    driver.confirm();
    driver.busy.value = true;

    await driver.runConfirmed();

    expect(driver.status.value).not.toBe('running');
    expect(driver.rows.value).toHaveLength(0);
    expect(driver.index.value).toBeNull();
  });

  it('clears a previous run\'s results when a new pairing is proposed', async () => {
    // Rows and an index describe the pairing that produced them. Proposing
    // again is a new question, and every other path into `invalidate()` already
    // says so; this one did not, so a completed run's rows sat under a fresh,
    // unconfirmed grid.
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();
    expect(driver.rows.value.length).toBeGreaterThan(0);
    expect(driver.index.value).not.toBeNull();

    await driver.proposePairs();

    expect(driver.rows.value).toHaveLength(0);
    expect(driver.index.value).toBeNull();
    expect(driver.confirmed.value).toBe(false);
  });

  it('never offers an unreadable document as a re-point target, and refuses one anyway', async () => {
    // The assignment never sees an unreadable document, so a re-point onto one
    // is discarded — but only AFTER the proposal has been rebuilt without the
    // row that held the real counterpart. The result was a correct pairing
    // destroyed in silence, and the surviving row claiming "no counterpart in
    // the other bundle" about a document sitting two rows down at 100%.
    const corrupt = new File([new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00])],
      'Project Falcon SPA v8 (corrupt).docx',
      { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [
      docxFile('Project Falcon SPA v7.docx',
        SPA.replace('in full on the Completion Date', 'in cleared funds on Completion')),
      corrupt,
    ]);
    await driver.proposePairs();

    const pair = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    const before = driver.proposal.value?.rows.map(r => r.id);
    const corruptId = driver.proposal.value?.documents
      .find(d => d.fileName.includes('corrupt'))?.docId as string;

    // Not offered: neither as a candidate nor in the proposal's pairable set.
    expect(driver.candidatesFor(pair?.id as string).map(c => c.b)).not.toContain(corruptId);
    expect(driver.proposal.value?.documents.find(d => d.docId === corruptId)?.unreadable)
      .toMatch(/could not be read/);

    // And refused if asked directly — `repoint` is public API.
    driver.repoint(pair?.id as string, corruptId);
    expect(driver.proposal.value?.rows.map(r => r.id)).toEqual(before);
    expect(driver.proposal.value?.rows.find(r => r.id === pair?.id)?.b).toBe(pair?.b);
  });

  it('tells an over-cap document with no counterpart about the cap, not about corruption', async () => {
    // It is never read (§4.4), so it reaches the unmatched branch looking
    // exactly like a file that could not be read. Saying so would send the user
    // looking for corruption in a perfectly good document.
    const huge = docxFile('Warranty Schedule.docx', DEED);
    Object.defineProperty(huge, 'size', { value: MAX_DOCX_BYTES + 1 });

    const driver = createBundleDriver();
    driver.addDocuments('a', [huge]);
    driver.addDocuments('b', [docxFile('Board Minutes March.docx', SPA)]);
    await driver.proposePairs();

    const row = driver.proposal.value?.rows.find(r => r.a !== null && r.b === null);
    expect(row?.reason).toMatch(/per-document cap/);
    expect(row?.reason).not.toMatch(/content could not be read/i);
    expect(driver.proposal.value?.disclosures.some(d => /larger than this build/i.test(d)))
      .toBe(true);
  });

  it('keeps the grid when a document\'s BYTES cannot be read, not only when they cannot be parsed', async () => {
    // The failure boundary is one document, which means everything done to that
    // document. `File.arrayBuffer()` rejects when the picked file has moved,
    // been truncated or lost permission since the picker ran — and with the
    // read outside the per-document `try`, that rejection escaped both loops
    // and left `proposal` null: no grid, and every other document with it.
    const moved = docxFile('Side Letter (moved).docx', DEED);
    Object.defineProperty(moved, 'arrayBuffer', {
      value: () => Promise.reject(
        new DOMException('The requested file could not be read', 'NotReadableError')),
    });

    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA), moved]);
    driver.addDocuments('b', [docxFile('Project Falcon SPA v7.docx',
      SPA.replace('in full on the Completion Date', 'in cleared funds on Completion'))]);
    await driver.proposePairs();

    const proposal = driver.proposal.value;
    expect(proposal).not.toBeNull();
    expect(proposal?.rows.find(r => r.a !== null && r.b !== null)?.label)
      .toBe('Project Falcon SPA v6.docx ▸ Project Falcon SPA v7.docx');

    const bad = proposal?.rows.find(r => r.label === 'Side Letter (moved).docx');
    expect(bad?.reason).toMatch(/could not be read/);
    expect(bad?.reason).toMatch(/NotReadableError|could not be read/);
    expect(proposal?.documents).toHaveLength(3);
  });

  it('never reads a document that cannot be batched at all', async () => {
    // A `.pdf` is paired on its name by design (§2.1) and never compared, so
    // reading its bytes buys nothing — the same argument as the cap below, and
    // the same fix: branch before the read.
    const pdf = new File([new Uint8Array([0x25, 0x50, 0x44, 0x46])], 'SPA (clean).pdf',
      { type: 'application/pdf' });
    const reads = vi.spyOn(pdf, 'arrayBuffer');

    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
    driver.addDocuments('b', [pdf]);
    await driver.proposePairs();

    expect(reads).not.toHaveBeenCalled();
    const row = driver.proposal.value?.rows.find(r => r.b !== null);
    expect(row?.runnable).toBe(false);
  });

  it('never reads a document that is over the per-document cap', async () => {
    // It is going to be shown with the cap named and never compared, so
    // reading it would spend every one of its bytes and then hold its records
    // for the driver's life — the peak spec §4.3 exists to bound, paid for a
    // comparison that will not happen.
    const huge = docxFile('Project Falcon SPA v7 (uncompressed).docx', SPA);
    Object.defineProperty(huge, 'size', { value: MAX_DOCX_BYTES + 1 });
    const reads = vi.spyOn(huge, 'arrayBuffer');

    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [huge]);
    await driver.proposePairs();

    expect(reads).not.toHaveBeenCalled();

    // And it is still paired, still shown, and told why it will not run.
    const row = driver.proposal.value?.rows.find(r => r.b !== null);
    expect(row?.a).not.toBeNull();
    expect(row?.runnable).toBe(false);
    expect(row?.reason).toMatch(/per-document cap/);
  });

  it('keeps the grid — and the rest of the bundle — when one document cannot be read', async () => {
    // Spec §8 case 2. One `try` around the whole ingest loop turned a single
    // corrupt file into a null proposal: no grid at all, and every other
    // document in both bundles gone with it. The one-to-many runner has always
    // continued past a bad file; this path did not.
    const corrupt = new File([new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00])],
      'Side Letter (corrupt).docx',
      { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA), corrupt]);
    driver.addDocuments('b', [docxFile('Project Falcon SPA v7.docx',
      SPA.replace('in full on the Completion Date', 'in cleared funds on Completion'))]);
    await driver.proposePairs();

    const proposal = driver.proposal.value;
    expect(proposal).not.toBeNull();

    // The good pair is still proposed.
    const pair = proposal?.rows.find(r => r.a !== null && r.b !== null);
    expect(pair?.label).toBe('Project Falcon SPA v6.docx ▸ Project Falcon SPA v7.docx');

    // The bad one is SHOWN, with a reason naming it — not dropped, and not
    // paired with anything.
    const bad = proposal?.rows.find(r => r.label === 'Side Letter (corrupt).docx');
    expect(bad?.b).toBeNull();
    expect(bad?.band).toBe('unmatched');
    expect(bad?.reason).toMatch(/could not be read/);
    expect(proposal?.disclosures.some(d => /could not be read/i.test(d))).toBe(true);

    // Every document appears exactly once, as the grid's rule requires.
    expect(proposal?.documents).toHaveLength(3);
    expect(proposal?.rows).toHaveLength(2);

    // And the run still produces the redline the good pair earns.
    driver.confirm();
    await driver.runConfirmed();
    expect(driver.rows.value.filter(r => r.status === 'done')).toHaveLength(1);
  });

  it('degrades to on-demand artefacts past the ceiling rather than losing a pair', async () => {
    const driver = createBundleDriver({ artefactBudgetBytes: 1 });
    driver.addDocuments('a', [docxFile('SPA v6.docx', SPA)]);
    driver.addDocuments('b', [docxFile('SPA v7.docx', SPA.replace('in full', 'in cleared funds'))]);
    await driver.proposePairs();
    driver.confirm();
    await driver.runConfirmed();

    const row = driver.rows.value[0];
    expect(row?.status).toBe('done');
    expect(row?.changeCount).toBeGreaterThan(0);
    expect(row?.artefactOnDemand).toBe(true);
    expect(driver.index.value?.disclosures).toHaveLength(1);
  });
});

describe('results belong to the pairing that produced them', () => {
  // Found by review. Adding or removing a document already invalidated the
  // proposal and the confirmation, and that was tested; re-pointing a row did
  // not clear the ROWS from a previous run. The grid would then show one
  // pairing while the results beneath it described a superseded one, with
  // nothing on screen to say which was current.
  it('clears a previous run\'s rows when a pair is re-pointed', async () => {
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();

    expect(driver.rows.value.length).toBeGreaterThan(0);
    expect(driver.index.value).not.toBeNull();

    driver.unconfirm();
    const row = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    driver.repoint(row?.id as string, null);

    expect(driver.rows.value).toEqual([]);
    expect(driver.index.value).toBeNull();
    expect(driver.status.value).toBe('idle');
    expect(driver.progress.value).toBe(0);
  });

  it('clears them when a document changes, too', async () => {
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();
    expect(driver.rows.value.length).toBeGreaterThan(0);

    driver.addDocuments('a', [docxFile('Late arrival.docx', DEED)]);

    expect(driver.rows.value).toEqual([]);
    expect(driver.index.value).toBeNull();
    expect(driver.confirmed.value).toBe(false);
  });
});

describe('downloadRow will not re-run a pair the grid declared unrunnable', () => {
  // An exported-API hazard rather than a live bug: the panel gates its button
  // on `isDownloadable`. But `downloadRow` is public, and without this guard an
  // unrunnable row's id would reach `produceArtefact`, which reads the
  // sketches — and a cross-format pair's side-B sketch has no records at all,
  // so it would package a delete-everything "redline" for a comparison
  // ADR-0017 gates.
  it('reports that there is no redline instead of producing one', async () => {
    // The real path: a Word document against a PDF. The pair is proposed and
    // shown (M8), marked unrunnable, and run as an `unrunnable` row.
    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [
      new File([new Uint8Array([1, 2, 3, 4])], 'Project Falcon SPA v7.pdf',
        { type: 'application/pdf' }),
    ]);
    await driver.proposePairs();

    const paired = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    expect(paired?.runnable).toBe(false);

    driver.confirm();
    await driver.runConfirmed();

    const row = driver.rows.value[0];
    expect(row?.status).toBe('unrunnable');

    await driver.downloadRow(row?.id as string);

    expect(driver.errorMessage.value).toMatch(/no redline/i);
    expect(driver.errorMessage.value).not.toMatch(/could not be produced/i);
  });
});

describe('a sketch is rebuilt when the flatten policy it was built under changes', () => {
  // Latent today — the bundle surface has no flatten control — but the shape of
  // the bug is the one spec §8 case 7 exists to prevent. `proposePairs` reused
  // any sketch that already existed, so changing flatten after a proposal left
  // records preprocessed under the OLD policy while the index stated the NEW
  // one: an index claiming a flatten choice its redlines were not produced
  // under.
  /**
   * A `.docx` carrying a real tracked insertion INSIDE its first paragraph.
   *
   * `synthesiseDocx` emits no revision markup, so under it acceptAll and
   * rejectAll flatten identically and any test of this would pass with the fix
   * ripped out — which is exactly what the first version of this test did.
   */
  function trackedChangeFile(name: string, text: string): File {
    const bundle = synthesiseDocx(text);
    const xml = bundle.part('word/document.xml') as string;
    bundle.setPart('word/document.xml', xml.replace(
      '</w:p>',
      '<w:ins w:id="900" w:author="Counterparty" w:date="2026-01-01T00:00:00Z">'
        + '<w:r><w:t xml:space="preserve"> and in cleared funds on Completion</w:t></w:r>'
        + '</w:ins></w:p>',
    ));
    return new File([new Uint8Array(bundle.toUint8Array())], name, {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  it('does not reuse records preprocessed under the previous policy', async () => {
    const driver = createBundleDriver();
    driver.addDocuments('a', [trackedChangeFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [docxFile('Project Falcon SPA v7.docx', SPA)]);

    await driver.proposePairs();
    const underAccept = driver.proposal.value?.rows[0]?.score?.content as number;

    driver.flattenA.value = 'rejectAll';
    await driver.proposePairs();
    const underReject = driver.proposal.value?.rows[0]?.score?.content as number;

    // Accepting the insertion changes the first paragraph's text, so its hash
    // changes and the Jaccard against side B moves. Rejecting restores the
    // paragraph exactly and the two sides agree. If the sketch were reused on
    // presence alone these two numbers would be identical — and the index would
    // then state a flatten policy the records were not built under.
    expect(underAccept).not.toBeCloseTo(underReject, 6);
    expect(underReject).toBe(1);

    driver.confirm();
    await driver.runConfirmed();
    expect(driver.index.value?.flatten).toEqual({ template: 'rejectAll', variant: 'acceptAll' });
    expect(driver.status.value).toBe('done');
  });
});

describe('opening one pair (spec §2.1 capability 8)', () => {
  // The batch drops every pair's ops the moment it completes — two
  // document-sized record arrays per pair is what breaks first at N — so the
  // change list is recovered by re-running that one pair, not remembered.
  it('returns the ops and the summary for a completed pair', async () => {
    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();

    const row = driver.rows.value.find(r => r.status === 'done');
    const detail = await driver.openPair(row?.id as string);

    expect(detail?.rowId).toBe(row?.id);
    expect(detail?.label).toBe(row?.label);
    expect(detail?.ops.length).toBeGreaterThan(0);
    // The same total the row reported. A drill-down that disagreed with the row
    // above it would leave the reader with two numbers and no way to choose.
    expect(detail?.summary.total).toBe(row?.changeCount);
  });

  it('re-runs the pipeline WITHOUT packaging a redline', async () => {
    // The wall this path exists behind. On web, materialising a redline means
    // triggering a download, so "show me what changed" must not save a file —
    // the same rule the one-to-many driver keeps.
    const clicked: string[] = [];
    vi.stubGlobal('URL', { ...URL, createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    const spy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(
      function (this: HTMLAnchorElement) { clicked.push(this.download); },
    );

    try {
      const driver = await loaded();
      driver.confirm();
      await driver.runConfirmed();
      await driver.openPair(driver.rows.value[0]?.id as string);
      expect(clicked).toEqual([]);
    } finally {
      spy.mockRestore();
      vi.unstubAllGlobals();
    }
  });

  it('declines a pair the grid marked unrunnable rather than comparing nothing', async () => {
    const driver = createBundleDriver();
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [
      new File([new Uint8Array([1, 2, 3, 4])], 'Project Falcon SPA v7.pdf',
        { type: 'application/pdf' }),
    ]);
    await driver.proposePairs();
    driver.confirm();
    await driver.runConfirmed();

    expect(await driver.openPair(driver.rows.value[0]?.id as string)).toBeNull();
  });

  it('declines an id it has never heard of', async () => {
    const driver = await loaded();
    expect(await driver.openPair('no-such-row')).toBeNull();
  });
});

describe('the remaining wiring a review found unpinned', () => {
  it('refuses to run under a flatten policy the pairing was not proposed under', async () => {
    // X-6 was fixed in `proposePairs` only, which left the run path open:
    // confirm, then change the policy, and the records were preprocessed under
    // one policy while the index stated the other. A confirmation is a
    // statement about specific documents UNDER a specific policy, so this is
    // treated like any other change to the inputs.
    const driver = await loaded();
    driver.confirm();
    driver.flattenA.value = 'rejectAll';

    await driver.runConfirmed();

    expect(driver.index.value).toBeNull();
    expect(driver.rows.value).toEqual([]);
    expect(driver.confirmed.value).toBe(false);
    expect(driver.proposal.value).toBeNull();
    expect(driver.errorMessage.value).toMatch(/tracked-changes setting changed/i);
  });

  it('opening a pair never packages a redline', async () => {
    // The structural form of "the drill-down does not take the artefact path".
    // The previous test asserted only that no download fired, which is a weaker
    // claim: a review inserted `await produceArtefact(id)` into `openPair` and
    // every test stayed green. `assembleRedlinePackage` is the step that turns
    // a comparison into a file, so spying on it catches the artefact path
    // whether or not anything is ultimately saved.
    //
    // Counting `arrayBuffer()` reads no longer works for this: since the
    // records moved out of the sketches (spec §4.3), re-running a pair legitimately
    // re-reads both documents to preprocess them.
    const packaged = vi.spyOn(writeout, 'assembleRedlinePackage');

    const driver = await loaded();
    driver.confirm();
    await driver.runConfirmed();
    packaged.mockClear();

    const detail = await driver.openPair(driver.rows.value[0]?.id as string);

    expect(detail?.ops.length).toBeGreaterThan(0);
    expect(packaged).not.toHaveBeenCalled();
  });

  it('and the artefact path does package one — so the spy above is live', async () => {
    // The control. Without it, "never called" could just mean the spy was never
    // wired to the module the driver actually imports.
    const packaged = vi.spyOn(writeout, 'assembleRedlinePackage');

    const driver = createBundleDriver({ artefactBudgetBytes: 1 });
    driver.addDocuments('a', [docxFile('Project Falcon SPA v6.docx', SPA)]);
    driver.addDocuments('b', [docxFile('Project Falcon SPA v7.docx',
      SPA.replace('in full on the Completion Date', 'in cleared funds on Completion'))]);
    await driver.proposePairs();
    driver.confirm();
    await driver.runConfirmed();
    packaged.mockClear();

    await driver.downloadRow(driver.rows.value[0]?.id as string);

    expect(packaged).toHaveBeenCalled();
  });

  it('rebuilds a re-pointed proposal without losing the other rows', async () => {
    // The re-point path goes back through `proposePairing`, so the whole
    // proposal is recomputed. Everything it was not asked to change must come
    // back identical — `batch.driver-config.test.ts` covers the half of this
    // that needs a non-default profile to see.
    const driver = await loaded();
    const before = driver.proposal.value;
    const row = before?.rows.find(r => r.a !== null && r.b !== null);
    const other = before?.rows.find(r => r.a !== null && r.b !== null && r.id !== row?.id);

    driver.repoint(row?.id as string, null);

    const after = driver.proposal.value;
    expect(after?.documents).toEqual(before?.documents);
    expect(after?.rows.find(r => r.a === row?.a)?.b).toBeNull();
    expect(after?.rows.find(r => r.a === other?.a)?.b).toBe(other?.b);
  });

  it('offers candidates scored under the profile\'s settings', async () => {
    // `candidatesFor` resolves its own settings rather than receiving them, and
    // nothing pinned that it resolves from the driver's config at all.
    const driver = await loaded();
    const row = driver.proposal.value?.rows.find(r => r.a !== null && r.b !== null);
    const candidates = driver.candidatesFor(row?.id as string);

    expect(candidates.length).toBe(driver.sideB.value.length);
    // The proposal's own score for this pair must be the one `candidatesFor`
    // reports for the same two documents — two different scores for one pair
    // would mean the grid and the re-point control disagree.
    const same = candidates.find(c => c.b === row?.b);
    expect(same?.score).toBeCloseTo(row?.score?.score as number, 12);
  });
});

/**
 * ADR-0042 at the M:N seam.
 *
 * `preprocessBundleDocument` (`batch/pairing/sketch.ts`) is the ONLY place a
 * bundle document's `ParaRecord[]` comes from, so it is the only place the
 * many-to-many run can learn about numbering. Deleting the
 * `buildNumberingContext(...)` argument there left the full suite green — 206
 * files, zero failures — while restoring exactly the pre-ADR-0042 silence: a
 * re-levelled clause reported as no change at all in the bundle panel, and as
 * one change in the single-pair panel, on the same two documents.
 *
 * The assertion is the DRILL-DOWN, not the record shape: `openPair` is what the
 * user reads, and the divergence this guards against is two panels of one
 * product answering differently.
 */
describe('ADR-0042 — the numbering context reaches the many-to-many seam', () => {
  /**
   * Two bundles of one document each. The B-side clause is the SAME SENTENCE at
   * a different list level — no `w:numPr` anywhere in either body, so every
   * field the bundle path read before ADR-0042 is identical and the pair is a
   * perfect content match (which is also what makes the pairing trivial).
   */
  async function ranPair(a: readonly StyledPara[], b: readonly StyledPara[]) {
    const driver = createBundleDriver();
    driver.addDocuments('a', [styleNumberedFile('Project Falcon SPA v6.docx', a)]);
    driver.addDocuments('b', [styleNumberedFile('Project Falcon SPA v7.docx', b)]);
    await driver.proposePairs();
    driver.confirm();
    await driver.runConfirmed();

    const row = driver.rows.value.find(r => r.status === 'done');
    expect(row).toBeDefined();
    return { row: row!, detail: await driver.openPair(row!.id) };
  }

  it('reports a style-inherited re-levelling as a NUMBERING change, not a bare style one', async () => {
    const { row, detail } = await ranPair([HEAD, SUB_A], [HEAD, ['Clause1', SUB_A[1]]]);

    expect(row.changeCount).toBe(1);
    expect(detail?.summary.total).toBe(1);
    const property = detail?.ops.filter(o => o.type === 'property') ?? [];
    expect(property).toHaveLength(1);

    // The `w:pStyle` half of this op needs NO numbering context — the style ids
    // differ on their own, so a count of 1 passes with the wiring deleted. What
    // the wiring carries is the numbering half, and only these assertions can
    // see it: with no context every paragraph resolves to `numbering: null`, so
    // the identity compares equal and both labels are null.
    const change = property[0]?.propertyChange;
    expect(change?.numberingIdentity).toBe(true);
    expect(change?.from.numberLabel).toBe('(a)');
    expect(change?.to.numberLabel).toBe('2.');
  });

  it('resolves numbering on BOTH sides, so a pure renumbering reads as one', async () => {
    // A sub-clause inserted at the head: `(a)`→`(b)` and `(b)`→`(c)` follow it.
    // Identity unchanged, so §9.4's answer is `crossRefRenumber` — and dropping
    // the context from EITHER side makes that side's paragraphs identity-less,
    // which reads as `formatting` instead.
    const { detail } = await ranPair([HEAD, SUB_A, SUB_B], [HEAD, INSERTED, SUB_A, SUB_B]);

    expect(detail?.summary.byCategory.crossRefRenumber).toBe(2);
    expect(detail?.summary.byCategory.formatting).toBe(0);
    expect(detail?.summary.total).toBe(3); // two shifted labels + the insert
  });

  it('reports nothing when the two sides are the same document', async () => {
    const { row, detail } = await ranPair([HEAD, SUB_A, SUB_B], [HEAD, SUB_A, SUB_B]);

    expect(row.changeCount).toBe(0);
    expect(detail?.summary.total).toBe(0);
  });
});
