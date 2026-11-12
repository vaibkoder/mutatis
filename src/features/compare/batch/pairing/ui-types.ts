/**
 * UI contract for the pairing grid (ADR-0037; spec §6.6).
 *
 * The rule this contract exists to make structural: **nothing is compared
 * before the user confirms the pairing.** `confirmed` starts false, only
 * `confirm()` sets it, and the surface that mounts the grid must gate the run
 * on it. A mis-pairing does not fail loudly — comparing the SPA against the
 * disclosure letter produces a complete, well-formed, fully-populated redline
 * that a reader has no cue to distrust — so the gate is the feature, not a
 * courtesy.
 *
 * Same stub hazard as `batch/ui-types.ts`, and it bites harder here: a grid
 * rendered against the stub confirms nothing and runs nothing, which looks like
 * a broken button rather than a missing wire. Any surface mounting this grid
 * must inject the real state and must have a test asserting that it did.
 */
import type { ReadonlySignal, Signal } from '@preact/signals';

import type { PairingProposal } from './model.js';
import type { PairScore } from './score.js';

export interface PairingGridState {
  /** `null` until both bundles have been sketched. */
  proposal: Signal<PairingProposal | null>;
  /** False until the user has confirmed. Nothing may run while it is false. */
  confirmed: Signal<boolean>;
  /** Building sketches is real work; the grid says so rather than sitting blank. */
  busy: Signal<boolean>;
  /**
   * True while comparisons are actually running.
   *
   * The pairing is not editable then. `confirmed` alone is not enough to say
   * so: a run keeps it true, so the "change the pairing" affordance stayed live
   * for the whole batch, and taking it mid-run left finished rows and a written
   * index sitting under a grid that no longer described them. A run consumes
   * the pairing it was given; the way to change one is to cancel first.
   */
  running: ReadonlySignal<boolean>;

  /**
   * Every candidate for a row, best first — the full list, not the top two.
   * The row shows two rivals; re-pointing needs all of them, because the
   * document the user actually wants may be the one the solver ranked fifth.
   */
  candidatesFor(rowId: string): readonly PairScore[];

  /** Re-point a row at a different side-B document, or at nothing (`null`). */
  repoint(rowId: string, b: string | null): void;

  /** Confirm the grid as it stands. The only thing that opens the gate. */
  confirm(): void;

  /** Return to editing after confirming. Refused while a run is in flight. */
  unconfirm(): void;
}

/** A no-op state, for rendering the grid in isolation. See the hazard above. */
export function makeStubPairingGridState(
  signal: <T>(value: T) => Signal<T>,
): PairingGridState {
  return {
    proposal: signal<PairingProposal | null>(null),
    confirmed: signal(false),
    busy: signal(false),
    running: signal(false),
    candidatesFor: () => [],
    repoint: () => {},
    confirm: () => {},
    unconfirm: () => {},
  };
}
