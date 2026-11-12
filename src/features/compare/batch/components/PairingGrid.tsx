/**
 * The pairing grid and its confirmation gate (ADR-0037; spec §6.6).
 *
 * The one rule: **nothing is compared before the user confirms.** Everything
 * about this component follows from it.
 *
 *  - Every document from both bundles has a row, including the ones that pair
 *    with nothing. A document the user handed us is never absent from the grid.
 *  - A confidence band changes how a row LOOKS — pre-ticked, flagged, or shown
 *    as unmatched — and never whether it appears. There is no filter here, no
 *    "hide the confident ones", no threshold that collapses a row away.
 *  - Outside `auto-pair`, a row shows its top rivals **with their scores**, so
 *    a proposal can be checked rather than rubber-stamped.
 *  - The run button lives on the mounting surface and is gated on `confirmed`.
 *    This component never runs anything itself.
 *
 * The reason for the ceremony: a mis-pairing produces a complete, well-formed,
 * fully-populated redline of one document against an unrelated one. It does not
 * look wrong. That is the pairing analogue of suppressing a change, and ADR-0008
 * forbids it.
 */
import { signal } from '@preact/signals';

import type { PairingBand, PairingRow } from '../pairing/model.js';
import type { PairScore } from '../pairing/score.js';
import type { PairingGridState } from '../pairing/ui-types.js';
import { makeStubPairingGridState } from '../pairing/ui-types.js';

export interface PairingGridProps {
  state?: PairingGridState;
}

const BAND_LABEL: Record<PairingBand, string> = {
  'auto-pair': 'Confident',
  'propose-with-rivals': 'Check this',
  unmatched: 'No counterpart',
};

/** A score as a percentage, which is how people read a similarity. */
const pct = (n: number): string => `${Math.round(n * 100)}%`;

/** Why this pair scored what it did — the workings, not just the total. */
function Workings({ score }: { score: PairScore }) {
  return (
    <span class="dp-pairing__workings">
      {score.contentAvailable
        ? `${pct(score.content)} of paragraphs match`
        : 'no text to compare'}
      {' · '}
      {`filenames ${pct(score.filename)}`}
    </span>
  );
}

/**
 * The rivals for one row, named by the OTHER document in each candidate.
 *
 * `anchor` is the row's own document. Which end of a `PairScore` varies is not
 * a constant: for a side-A row the rivals differ in `.b`, and for a side-B
 * document with no counterpart they differ in `.a`. Reading `.b` unconditionally
 * printed the row's own file name twice, at two different scores — so the one
 * row type whose rivals are the whole evidence that "unmatched" is right showed
 * no evidence at all.
 */
function RivalList({ rivals, anchor, nameOf }: {
  rivals: readonly PairScore[];
  anchor: string;
  nameOf: (id: string) => string;
}) {
  if (rivals.length === 0) return null;
  return (
    <span class="dp-pairing__rivals">
      Next best:{' '}
      {rivals.map((r, i) => {
        const other = r.a === anchor ? r.b : r.a;
        return (
          <span key={other}>
            {i > 0 ? ', ' : ''}
            {nameOf(other)} ({pct(r.score)})
          </span>
        );
      })}
    </span>
  );
}

/**
 * The re-point control for one row.
 *
 * Its own component so `candidatesFor` — a full Jaccard of this document
 * against every document on the other side — is called only where the scores
 * can actually be used: on a row that offers the control, and only while the
 * control is live. It used to run from `GridRow`'s body, for every row on every
 * render, and the surface re-renders on each row update DURING a run — when the
 * grid is locked and every one of those scores is unreadable. At the §4.4 caps
 * that was fifty score matrices per completed pair, on the main thread §4.2
 * goes out of its way to leave a core for.
 *
 * It is still one matrix per unlocked row per render, which is the cost of
 * showing a score against every option in the list. That is bounded by the
 * user's own editing, not by the batch.
 */
function RepointControl({ row, state, locked, sideBNames }: {
  row: PairingRow;
  state: PairingGridState;
  locked: boolean;
  sideBNames: readonly { id: string; label: string }[];
}) {
  const scoreOf = locked
    ? new Map<string, number>()
    : new Map(state.candidatesFor(row.id).map(c => [c.b, c.score]));
  return (
    <label class="dp-pairing__repoint">
      <span class="dp-pairing__repoint-label">Compare against</span>
      <select
        disabled={locked}
        value={row.b ?? ''}
        onChange={e => {
          const value = (e.currentTarget as HTMLSelectElement).value;
          state.repoint(row.id, value === '' ? null : value);
        }}
      >
        <option value="">Nothing — leave unmatched</option>
        {sideBNames.map(b => (
          <option key={b.id} value={b.id}>
            {b.label}
            {scoreOf.has(b.id) ? ` (${pct(scoreOf.get(b.id) as number)})` : ''}
          </option>
        ))}
      </select>
    </label>
  );
}

function GridRow({ row, state, nameOf, sideBNames, pairable }: {
  row: PairingRow;
  state: PairingGridState;
  nameOf: (id: string) => string;
  sideBNames: readonly { id: string; label: string }[];
  /** Whether this row's document can be paired at all. */
  pairable: boolean;
}) {
  // `confirmed` alone: a run keeps it true from start to finish, and the only
  // route back to false — the "change the pairing" button — is withdrawn for
  // the duration of a run, below. Adding `|| running` here as well would be a
  // branch nothing can reach and no test can pin.
  const locked = state.confirmed.value;

  return (
    <li class={`dp-pairing__row dp-pairing__row--${row.band}`}>
      <span class="dp-pairing__label" title={row.label}>{row.label}</span>

      <span class={`dp-pairing__band dp-pairing__band--${row.band}`}>
        {BAND_LABEL[row.band]}
        {row.score !== null && ` · ${pct(row.score.score)}`}
      </span>

      {row.score !== null && <Workings score={row.score} />}

      {/* Rivals are shown for everything except a confident pair — including
          for an unmatched document, where they are the pairings the solver
          considered and turned down. "Unmatched" should be checkable too. */}
      {row.band !== 'auto-pair' && (
        <RivalList rivals={row.rivals} anchor={(row.a ?? row.b) as string} nameOf={nameOf} />
      )}

      {row.reason !== undefined && (
        <span class="dp-pairing__reason">{row.reason}</span>
      )}

      {/* Re-pointing offers EVERY side-B document, not just the top rivals:
          the one the user wants may be the one the solver ranked fifth. */}
      {row.a !== null && pairable && (
        <RepointControl row={row} state={state} locked={locked} sideBNames={sideBNames} />
      )}
    </li>
  );
}

export function PairingGrid({ state }: PairingGridProps) {
  const grid = state ?? makeStubPairingGridState(signal);
  const proposal = grid.proposal.value;

  if (grid.busy.value) {
    return (
      <section class="dp-pairing" aria-label="Proposed document pairing">
        <p class="dp-pairing__note">Reading both bundles and proposing a pairing…</p>
      </section>
    );
  }

  if (proposal === null) {
    return (
      <section class="dp-pairing" aria-label="Proposed document pairing">
        <p class="dp-pairing__note">
          Add documents to both sides to see a proposed pairing.
        </p>
      </section>
    );
  }

  // Names come from the proposal's own document list, never from splitting a
  // row label: a file name containing the separator would corrupt every name.
  const nameFor = new Map(proposal.documents.map(d => [d.docId, d.fileName]));
  // Only READABLE side-B documents are offered. An unreadable one is not a
  // pairing candidate — the assignment never sees it — so offering it (with a
  // percentage, no less) let a user pick something that was then silently
  // discarded, after the row holding the real counterpart had already been
  // rebuilt away. A control that can be used and cannot work is worse than no
  // control.
  const sideBNames = proposal.documents
    .filter(d => d.side === 'b' && d.unreadable === undefined)
    .map(d => ({ id: d.docId, label: d.fileName }));
  const unreadable = new Set(
    proposal.documents.filter(d => d.unreadable !== undefined).map(d => d.docId),
  );
  const nameOf = (id: string): string => nameFor.get(id) ?? id;
  const totals = proposal.totals;

  return (
    <section class="dp-pairing" aria-label="Proposed document pairing">
      <p class="dp-pairing__intro">
        These pairings are <strong>proposed, not decided</strong>. Check them before
        anything is compared — a wrongly paired document still produces a complete
        redline, and it would look just like a real one.
      </p>

      <p class="dp-pairing__totals">
        {totals.pairs} pair{totals.pairs === 1 ? '' : 's'}
        {' · '}{totals.autoPair} confident
        {' · '}{totals.needsReview} to check
        {' · '}{totals.unmatchedA + totals.unmatchedB} unmatched
        {totals.unrunnable > 0 && ` · ${totals.unrunnable} cannot be compared`}
      </p>

      <ul class="dp-pairing__rows">
        {proposal.rows.map(row => (
          <GridRow
            key={row.id}
            row={row}
            state={grid}
            nameOf={nameOf}
            sideBNames={sideBNames}
            pairable={row.a === null || !unreadable.has(row.a)}
          />
        ))}
      </ul>

      {proposal.disclosures.map(text => (
        <p key={text} class="dp-pairing__disclosure">{text}</p>
      ))}

      <div class="dp-pairing__actions">
        {grid.confirmed.value ? (
          <>
            <span class="dp-pairing__confirmed" role="status">
              Pairing confirmed.
            </span>
            {grid.running.value ? (
              <span class="dp-pairing__note">
                Comparisons are running against it. Cancel the run to change it.
              </span>
            ) : (
              <button type="button" class="dp-batch__link" onClick={() => grid.unconfirm()}>
                Change the pairing
              </button>
            )}
          </>
        ) : (
          <button
            type="button"
            class="dp-pairing__confirm"
            onClick={() => grid.confirm()}
          >
            Confirm this pairing
          </button>
        )}
      </div>
    </section>
  );
}
