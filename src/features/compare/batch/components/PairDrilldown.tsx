/**
 * One pair's change list, recovered by re-running that pair (ADR-0037;
 * spec §2.1 capability 8, §5.4, §6.7).
 *
 * Shared by both batch surfaces. It existed twice — copied verbatim from the
 * one-to-many panel into the many-to-many one, differing in a single error
 * string — which a review caught against the reuse-before-build rule
 * rule. Two copies of a rule like "the drill-down is never pre-filtered" is one
 * copy that will eventually be edited alone.
 *
 * The `open` callback belongs to the caller because the two drivers are
 * different types; everything downstream of the returned `PairDetail` is the
 * same, and lives here.
 *
 * Host-free: no Office JS, no network.
 */
import { useState } from 'preact/hooks';

import { ChangeList } from '../../components/ChangeList.js';
import type { PairDetail } from '../ui-types.js';

export interface PairDrilldownState {
  detail: PairDetail | null;
  opening: string | null;
  error: string | null;
  open: (rowId: string) => void;
  close: () => void;
}

/**
 * Wire a driver's `openPair` to the drill-down's own state.
 *
 * `unavailableMessage` is the one thing the two surfaces word differently —
 * "run the batch again" against "run the comparisons again" — because they are
 * different actions on screen.
 */
export function usePairDrilldown(
  openPair: (rowId: string) => Promise<PairDetail | null>,
  unavailableMessage: string,
): PairDrilldownState {
  const [detail, setDetail] = useState<PairDetail | null>(null);
  const [opening, setOpening] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  return {
    detail,
    opening,
    error,
    close: () => setDetail(null),
    open: (rowId: string): void => {
      setOpening(rowId);
      setDetail(null);
      setError(null);
      void openPair(rowId).then(
        result => {
          setDetail(result);
          setOpening(null);
          // `null` means the pair could not be reopened — most often because
          // the batch was cleared, or because the grid marked it unrunnable.
          // Saying nothing would leave a button that appears to do nothing.
          if (result === null) setError(unavailableMessage);
        },
        // A re-run reads records and runs the pipeline, so it can throw.
        // Without this the surface would sit on "Re-running…" for ever and the
        // rejection would go unhandled.
        (err: unknown) => {
          setError(err instanceof Error ? err.message : String(err));
          setOpening(null);
        },
      );
    },
  };
}

export function PairDrilldown({ state }: { state: PairDrilldownState }) {
  return (
    <>
      {state.opening !== null && (
        <p class="dp-batch__drilldown-note">Re-running that comparison…</p>
      )}

      {state.error !== null && (
        <p class="dp-batch__error" role="alert">
          That comparison could not be re-run: {state.error}
        </p>
      )}

      {state.detail !== null && (
        <section class="dp-batch__drilldown" aria-label={`Changes in ${state.detail.label}`}>
          <header class="dp-batch__drilldown-head">
            <h3>{state.detail.label}</h3>
            <button type="button" class="dp-batch__link" onClick={state.close}>
              Close
            </button>
          </header>
          <ChangeList
            ops={state.detail.ops}
            // The FULL change set. Filters are a navigation aid on the
            // single-pair surface; a drill-down that arrived pre-filtered would
            // show less than the row's own count claims (ADR-0008).
            filters={{ hideFormattingOnly: false, collapseRenumbering: false }}
            onGoToChange={() => { /* no host document to navigate to on web */ }}
          />
        </section>
      )}
    </>
  );
}
