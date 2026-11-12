/**
 * The pairing grid and its confirmation gate (ADR-0037; spec §6.6;
 * acceptance M7).
 *
 * M7 is the one that matters and it is asserted the hard way: a spy standing in
 * for the batch runner must record ZERO invocations until the grid has been
 * confirmed. Everything else here is the same rule seen from a different angle
 * — every document has a row, a band never removes one, and a proposal can be
 * argued with rather than only accepted.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render } from 'preact';
import { act } from 'preact/test-utils';
import { signal } from '@preact/signals';

import { PairingGrid } from '@features/compare/batch/components/PairingGrid.js';
import { makeStubPairingGridState } from '@features/compare/batch/pairing/ui-types.js';
import type { PairingGridState } from '@features/compare/batch/pairing/ui-types.js';
import { proposePairing } from '@features/compare/batch/pairing/propose.js';
import type { PairingProposal } from '@features/compare/batch/pairing/model.js';
import type { PairingSketch } from '@features/compare/batch/pairing/sketch.js';

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

const AT = '2026-09-10T09:00:00.000Z';

function sketch(
  docId: string,
  fileName: string,
  hashes: readonly string[],
  over: Partial<PairingSketch> = {},
): PairingSketch {
  return {
    docId,
    side: docId.startsWith('a') ? 'a' : 'b',
    sourceKind: 'docx',
    fileName,
    byteLength: 4096,
    paraHashes: new Set(hashes),
    definedTerms: new Set(['seller', 'buyer']),
    partyNames: new Set(['seller', 'buyer']),
    quarantinedCount: 0,
    ...over,
  };
}

const body = (p: string, n: number, drift = 0): string[] =>
  Array.from({ length: n }, (_, i) => (i < drift ? `${p}-new-${i}` : `${p}-${i}`));

/** A bundle with one confident pair, one ambiguous pair, and one orphan. */
function realisticProposal(): PairingProposal {
  return proposePairing({
    sideA: [
      sketch('a-spa', 'Project Falcon SPA v6.docx', body('spa', 60)),
      sketch('a-dl', 'Disclosure Letter v3.docx', body('dl', 30)),
      sketch('a-old', 'Escrow Deed.docx', body('escrow', 20)),
    ],
    sideB: [
      sketch('b-spa', 'Project Falcon SPA v7.docx', body('spa', 60, 3)),
      sketch('b-dl6', 'Disclosure Letter v6.docx', body('dl', 30, 1)),
      sketch('b-dl7', 'Disclosure Letter v7.docx', body('dl', 30, 2)),
    ],
    generatedAt: AT,
  });
}

/** A live-ish state: real signals, spied actions. */
function liveState(proposal: PairingProposal | null): PairingGridState {
  const state = makeStubPairingGridState(signal);
  state.proposal.value = proposal;
  return {
    ...state,
    candidatesFor: () => [],
    repoint: vi.fn(),
    confirm: vi.fn(() => { state.confirmed.value = true; }),
    unconfirm: vi.fn(() => { state.confirmed.value = false; }),
  };
}

const textOf = (el: Element | null): string => el?.textContent ?? '';

describe('M7 — nothing runs before the grid is confirmed', () => {
  it('records zero runner invocations until Confirm is pressed', async () => {
    // The surface's own gate, modelled exactly as the real one is: the run
    // action is only ever reachable while `confirmed` is true.
    const runBatch = vi.fn();
    const state = liveState(realisticProposal());

    function Surface() {
      return (
        <div>
          <PairingGrid state={state} />
          <button
            type="button"
            class="test-run"
            disabled={!state.confirmed.value}
            onClick={() => runBatch()}
          >
            Run comparisons
          </button>
        </div>
      );
    }

    await act(async () => { render(<Surface />, container); });
    const run = container.querySelector('.test-run') as HTMLButtonElement;

    // Before confirming: disabled, and clicking it anyway does nothing.
    expect(run.disabled).toBe(true);
    await act(async () => { run.click(); });
    expect(runBatch).not.toHaveBeenCalled();

    // Interacting with the grid must not run anything either.
    const select = container.querySelector('.dp-pairing__repoint select') as HTMLSelectElement;
    await act(async () => {
      select.value = '';
      select.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(runBatch).not.toHaveBeenCalled();

    // Only after confirming.
    const confirm = container.querySelector('.dp-pairing__confirm') as HTMLButtonElement;
    await act(async () => { confirm.click(); });
    expect(state.confirmed.value).toBe(true);

    const runAfter = container.querySelector('.test-run') as HTMLButtonElement;
    expect(runAfter.disabled).toBe(false);
    await act(async () => { runAfter.click(); });
    expect(runBatch).toHaveBeenCalledTimes(1);
  });

  it('locks the re-point controls once confirmed, and unlocks them again', async () => {
    const state = liveState(realisticProposal());
    await act(async () => { render(<PairingGrid state={state} />, container); });

    const selects = () =>
      Array.from(container.querySelectorAll<HTMLSelectElement>('.dp-pairing__repoint select'));
    expect(selects().every(s => !s.disabled)).toBe(true);

    await act(async () => {
      (container.querySelector('.dp-pairing__confirm') as HTMLButtonElement).click();
    });
    expect(selects().every(s => s.disabled)).toBe(true);
    expect(textOf(container.querySelector('.dp-pairing__confirmed'))).toMatch(/confirmed/i);

    // And the user can go back — confirming is a gate, not a trapdoor.
    await act(async () => {
      const back = Array.from(container.querySelectorAll('button'))
        .find(b => /change the pairing/i.test(b.textContent ?? '')) as HTMLButtonElement;
      back.click();
    });
    expect(selects().every(s => !s.disabled)).toBe(true);
  });
});

describe('the grid shows everything', () => {
  it('gives every document from both bundles a row', () => {
    const proposal = realisticProposal();
    render(<PairingGrid state={liveState(proposal)} />, container);

    const rows = container.querySelectorAll('.dp-pairing__row');
    expect(rows).toHaveLength(proposal.rows.length);
    // Three on side A and three on side B, in two pairs and two orphans.
    expect(proposal.documents).toHaveLength(6);
  });

  it('shows rivals with their scores on a row that needs checking', () => {
    render(<PairingGrid state={liveState(realisticProposal())} />, container);
    const propose = container.querySelector('.dp-pairing__row--propose-with-rivals');

    expect(propose).toBeTruthy();
    const rivals = textOf(propose?.querySelector('.dp-pairing__rivals') ?? null);
    expect(rivals).toMatch(/next best/i);
    // A name and a percentage — an argument, not an assertion.
    expect(rivals).toMatch(/Disclosure Letter v\d+\.docx \(\d+%\)/);
  });

  it('shows the workings, not just a total', () => {
    render(<PairingGrid state={liveState(realisticProposal())} />, container);
    const workings = textOf(container.querySelector('.dp-pairing__workings'));
    expect(workings).toMatch(/of paragraphs match/);
    expect(workings).toMatch(/filenames \d+%/);
  });

  it('shows an unmatched document with its reason rather than omitting it', () => {
    render(<PairingGrid state={liveState(realisticProposal())} />, container);
    const unmatched = container.querySelector('.dp-pairing__row--unmatched');

    expect(unmatched).toBeTruthy();
    expect(textOf(unmatched)).toMatch(/no counterpart/i);
  });

  it('offers every side-B document when re-pointing, not only the top rivals', () => {
    // The document the user wants may be the one the solver ranked last.
    const proposal = realisticProposal();
    render(<PairingGrid state={liveState(proposal)} />, container);
    const select = container.querySelector('.dp-pairing__repoint select') as HTMLSelectElement;

    const values = Array.from(select.options).map(o => o.value);
    expect(values).toContain('');            // "leave unmatched"
    for (const doc of proposal.documents.filter(d => d.side === 'b')) {
      expect(values).toContain(doc.docId);
    }
  });

  it('surfaces a cross-format pair with its pairing and its reason', () => {
    const proposal = proposePairing({
      sideA: [sketch('a-spa', 'Project Falcon SPA.docx', body('spa', 40))],
      sideB: [sketch('b-spa', 'Project Falcon SPA.pdf', [], { sourceKind: 'pdf' })],
      generatedAt: AT,
    });
    render(<PairingGrid state={liveState(proposal)} />, container);

    const text = textOf(container.querySelector('.dp-pairing__rows'));
    expect(text).toMatch(/Project Falcon SPA\.docx/);
    expect(text).toMatch(/Project Falcon SPA\.pdf/);
    expect(text).toMatch(/out of scope/i);
    // Paired and scored — shown, not withheld.
    expect(container.querySelector('.dp-pairing__band')?.textContent).toMatch(/\d+%/);
  });

  it('says plainly that the pairings are proposed rather than decided', () => {
    render(<PairingGrid state={liveState(realisticProposal())} />, container);
    expect(textOf(container.querySelector('.dp-pairing__intro')))
      .toMatch(/proposed, not decided/i);
  });
});

describe('the grid renders in isolation without doing anything', () => {
  it('prompts for documents when there is no proposal', () => {
    render(<PairingGrid />, container);
    expect(textOf(container.querySelector('.dp-pairing__note'))).toMatch(/add documents/i);
    expect(container.querySelectorAll('.dp-pairing__row')).toHaveLength(0);
  });

  it('says it is working while the sketches are being built', () => {
    const state = makeStubPairingGridState(signal);
    state.busy.value = true;
    render(<PairingGrid state={state} />, container);
    expect(textOf(container.querySelector('.dp-pairing__note'))).toMatch(/proposing a pairing/i);
  });
});

describe('re-pointing reports the user\'s choice verbatim', () => {
  it('passes the chosen document id, and null for unmatched', async () => {
    const state = liveState(realisticProposal());
    await act(async () => { render(<PairingGrid state={state} />, container); });
    const select = container.querySelector('.dp-pairing__repoint select') as HTMLSelectElement;
    const rowId = (container.querySelector('.dp-pairing__row') as HTMLElement);
    expect(rowId).toBeTruthy();

    const target = Array.from(select.options).find(o => o.value !== '' && o.value !== select.value);
    await act(async () => {
      select.value = (target as HTMLOptionElement).value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(state.repoint).toHaveBeenCalledWith(expect.any(String), (target as HTMLOptionElement).value);

    await act(async () => {
      select.value = '';
      select.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(state.repoint).toHaveBeenLastCalledWith(expect.any(String), null);
  });
});

describe('an unmatched row is checkable, and a run makes the grid read-only', () => {
  it('names the OTHER document in a side-B orphan\'s rival list, not itself', () => {
    // The rivals of an unmatched side-B document vary in `.a`, not `.b` — they
    // are the side-A documents the solver considered and turned down. Reading
    // `.b` printed the row's own file name back at itself, twice, at two
    // different scores: the one row type whose whole justification IS its
    // rivals showed no evidence at all.
    render(<PairingGrid state={liveState(realisticProposal())} />, container);

    const orphan = Array.from(container.querySelectorAll('.dp-pairing__row'))
      .find(li => textOf(li.querySelector('.dp-pairing__label')) === 'Disclosure Letter v7.docx');
    const rivals = textOf(orphan?.querySelector('.dp-pairing__rivals') ?? null);

    expect(rivals).toMatch(/Next best/);
    expect(rivals).toMatch(/Disclosure Letter v3\.docx/);
    expect(rivals).not.toMatch(/Disclosure Letter v7\.docx/);
  });

  it('withdraws "change the pairing" and locks the selects while comparisons run', async () => {
    // A run consumes the pairing it was given. Re-opening the grid mid-batch
    // left finished rows and a written index describing a pairing the grid no
    // longer showed — and the run's own continuation then wrote the old pairing
    // back over the cleared state.
    const base = liveState(realisticProposal());
    const running = signal(false);
    const state: PairingGridState = { ...base, running };

    await act(async () => { render(<PairingGrid state={state} />, container); });
    (container.querySelector('.dp-pairing__confirm') as HTMLButtonElement).click();
    await act(async () => {});
    expect(container.querySelector('.dp-batch__link')).toBeTruthy();

    await act(async () => { running.value = true; });

    expect(container.querySelector('.dp-batch__link')).toBeNull();
    expect(textOf(container.querySelector('.dp-pairing__note'))).toMatch(/cancel the run/i);
    // And the selects stay shut throughout — `confirmed` is what locks them,
    // and a run never clears it.
    for (const sel of container.querySelectorAll('.dp-pairing__repoint select')) {
      expect((sel as HTMLSelectElement).disabled).toBe(true);
    }
  });
});

describe('an unreadable document is shown, and is not a pairing candidate', () => {
  const unreadableProposal = (): PairingProposal => proposePairing({
    sideA: [
      sketch('a-spa', 'Project Falcon SPA v6.docx', body('spa', 60)),
      { ...sketch('a-bad', 'Escrow Deed (truncated).docx', []),
        unreadable: '"Escrow Deed (truncated).docx" could not be read as a Word document.' },
    ],
    sideB: [
      sketch('b-spa', 'Project Falcon SPA v7.docx', body('spa', 60, 3)),
      { ...sketch('b-bad', 'Project Falcon SPA v8 (corrupt).docx', []),
        unreadable: '"Project Falcon SPA v8 (corrupt).docx" could not be read as a Word document.' },
    ],
    generatedAt: AT,
  });

  it('gives it a row with its reason', () => {
    render(<PairingGrid state={liveState(unreadableProposal())} />, container);

    const text = textOf(container.querySelector('.dp-pairing__rows'));
    expect(text).toMatch(/Project Falcon SPA v8 \(corrupt\)\.docx/);
    expect(text).toMatch(/could not be read as a Word document/);
  });

  it('scores candidates only for the rows that actually offer a control', async () => {
    // `candidatesFor` is a full Jaccard against every document on the other
    // side. Called from the row body it ran for every row on every render, and
    // the surface re-renders on each row update during a run.
    const state = liveState(unreadableProposal());
    const scored: string[] = [];
    const counted: PairingGridState = {
      ...state,
      candidatesFor: (rowId: string) => { scored.push(rowId); return []; },
    };

    await act(async () => { render(<PairingGrid state={counted} />, container); });

    expect(scored).toHaveLength(1);
    expect(container.querySelectorAll('.dp-pairing__repoint')).toHaveLength(1);
  });

  it('scores nothing at all once the pairing is confirmed', async () => {
    // A locked control cannot use the scores, and the surface re-renders on
    // every row update for the whole length of a run — which is exactly when
    // the grid is locked. At the §4.4 caps that was fifty full Jaccard
    // matrices per completed pair, on the main thread §4.2 leaves a core for.
    const state = liveState(unreadableProposal());
    const scored: string[] = [];
    const counted: PairingGridState = {
      ...state,
      candidatesFor: (rowId: string) => { scored.push(rowId); return []; },
    };

    await act(async () => { render(<PairingGrid state={counted} />, container); });
    scored.length = 0;
    await act(async () => { counted.confirmed.value = true; });

    expect(container.querySelectorAll('.dp-pairing__repoint')).toHaveLength(1);
    expect(scored).toHaveLength(0);
  });

  it('keeps it out of every re-point control', async () => {
    // Offering it — with a percentage — let a user pick something the
    // assignment never sees, which silently discarded the choice AFTER the row
    // holding the real counterpart had been rebuilt away.
    const state = liveState(unreadableProposal());
    await act(async () => { render(<PairingGrid state={state} />, container); });

    const options = Array.from(container.querySelectorAll('.dp-pairing__repoint option'))
      .map(o => o.textContent ?? '');
    expect(options.some(o => /Project Falcon SPA v7/.test(o))).toBe(true);
    expect(options.some(o => /corrupt/.test(o))).toBe(false);

    // And an unreadable side-A document has no control of its own: there is
    // nothing it could be pointed at that the assignment would honour — and it
    // is never SCORED against the other bundle either, which is why the control
    // owns that call rather than every row's body.
    const rows = Array.from(container.querySelectorAll('.dp-pairing__row'));
    const bad = rows.find(li =>
      textOf(li.querySelector('.dp-pairing__label')) === 'Escrow Deed (truncated).docx');
    expect(bad).toBeTruthy();
    expect(bad?.querySelector('.dp-pairing__repoint')).toBeNull();
    expect(rows.filter(li => li.querySelector('.dp-pairing__repoint'))).toHaveLength(1);
  });
});
