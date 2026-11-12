import { buildPairingSketch } from '@features/compare/batch/pairing/sketch.js';
import { proposePairing, repointPairing } from '@features/compare/batch/pairing/propose.js';
import { getCompareConfig, SG_PROFILE, UK_PROFILE } from '@shared/config/index.js';
import type { PairingSketch } from '@features/compare/batch/pairing/sketch.js';
import type { PairingProposal } from '@features/compare/batch/pairing/model.js';

/**
 * Build a pairing proposal for a many-to-many bundle.
 *
 * IMPORT DISCIPLINE: deep paths only. `@features/compare/batch/index.js` is a
 * barrel that re-exports `BatchComparePanel` — a Preact `.tsx`. Importing the
 * barrel would drag Preact into a React bundle.
 *
 * The maths underneath is pure data and reused verbatim: sketching, scoring,
 * Hungarian assignment, banding and the rival lists are all upstream's. Only
 * the orchestration around them is ours, because upstream's is signals-based.
 */

export interface PairingDocInput {
  readonly docId: string;
  readonly side: 'a' | 'b';
  readonly fileName: string;
  readonly bytes: ArrayBuffer;
}

export interface PairingResult {
  readonly proposal: PairingProposal;
  /**
   * The sketches are returned alongside the proposal so a re-point can be
   * recomputed WITHOUT re-reading document bytes. `repointPairing` needs the
   * same `ProposePairingInput` that produced the proposal, and sketches are
   * structured-clone-safe (Sets survive), so they cross the worker boundary
   * once and every subsequent re-point is instant and local.
   */
  readonly sideA: readonly PairingSketch[];
  readonly sideB: readonly PairingSketch[];
}

export function buildPairing(
  docs: readonly PairingDocInput[],
  jurisdiction: 'UK' | 'SG',
  generatedAt: string,
): PairingResult {
  const profile = jurisdiction === 'SG' ? SG_PROFILE : UK_PROFILE;
  const config = getCompareConfig(profile);

  const sketches = docs.map((doc) =>
    buildPairingSketch({
      docId: doc.docId,
      side: doc.side,
      fileName: doc.fileName,
      bytes: doc.bytes,
      config,
      profile,
      // Pre-existing tracked changes are accepted before sketching so the
      // content signal reflects the document as it reads, not as it was marked
      // up. Same policy the compare pipeline uses by default.
      flatten: 'acceptAll',
    }),
  );

  const sideA = sketches.filter((s) => s.side === 'a');
  const sideB = sketches.filter((s) => s.side === 'b');

  return {
    proposal: proposePairing({ sideA, sideB, generatedAt, config }),
    sideA,
    sideB,
  };
}

/** Recompute a proposal after the user re-points one row. Pure; no bytes needed. */
export function repoint(
  result: PairingResult,
  jurisdiction: 'UK' | 'SG',
  rowId: string,
  b: string | null,
): PairingResult {
  const profile = jurisdiction === 'SG' ? SG_PROFILE : UK_PROFILE;
  const config = getCompareConfig(profile);

  return {
    ...result,
    proposal: repointPairing(
      {
        sideA: result.sideA,
        sideB: result.sideB,
        generatedAt: result.proposal.generatedAt,
        config,
      },
      result.proposal,
      rowId,
      b,
    ),
  };
}
