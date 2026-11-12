// src/features/repair/scan.ts
import { wordRun } from '@shared/office/index.js';
import { detectMetadata } from './detect/metadata.js';
import { detectStructure, detectHeuristics } from './detect/structure.js';
import { detectTracked } from './detect/tracked.js';
import { detectOoxml } from './detect/ooxml-scan.js';
import type { RepairReport } from './types.js';
import type { StyleProfile } from './profiles.js';
import type { DocumentHeuristics } from './guards.js';
import type { DocxBundle } from '@shared/ooxml/index.js';

export interface ScanState {
  report: RepairReport;
  /** The bundle from the OOXML scan — pass to apply to avoid double-read. */
  bundle: DocxBundle;
  /** Document formatting heuristics for the court-profile guards (spec §7.2). */
  heuristics: DocumentHeuristics;
}

/**
 * Orchestrates the full document scan.
 *
 * Phase 1: typed detectors (metadata, structure, tracked changes) run inside a
 *          single wordRun so they share one Office request context.
 * Phase 2: OOXML scan runs outside wordRun — it calls ooxmlService.readDocx()
 *          which uses its own Office.context.document handle.
 *
 * The returned ScanState carries both the RepairReport (phase:'scan') and the
 * DocxBundle so the apply phase can re-use the already-read ZIP without a
 * second round-trip to the host document.
 *
 * @param documentName  Display name for the report header.
 * @param profile       Style profile against which structural issues are checked.
 * @param generatedAt   ISO 8601 timestamp — caller provides so tests can inject a fixed value.
 */
export async function scan(
  documentName: string,
  profile: StyleProfile,
  generatedAt: string,
): Promise<ScanState> {
  // ── Phase 1: typed detectors ─────────────────────────────────────────────
  const { typedFindings, heuristics } = await wordRun(async ctx => {
    const meta      = await detectMetadata(ctx);
    const structure = await detectStructure(ctx, profile);
    const tracked   = await detectTracked(ctx);
    const h         = await detectHeuristics(ctx);
    return {
      typedFindings: [...meta, ...structure, ...(tracked !== null ? [tracked] : [])],
      heuristics: h,
    };
  });

  // ── Phase 2: OOXML scan ──────────────────────────────────────────────────
  // Runs separately; ooxmlService.readDocx() manages its own Office.context.
  const ooxmlResult = await detectOoxml();

  // ── Phase 3: merge findings ───────────────────────────────────────────────
  const allFindings = [...typedFindings, ...ooxmlResult.findings];

  // ── Phase 4: assemble report ──────────────────────────────────────────────
  const report: RepairReport = {
    documentName,
    generatedAt,
    profile: { id: profile.id, name: profile.name },
    phase: 'scan',
    findings: allFindings,
    metrics: {
      fileSizeBeforeBytes: ooxmlResult.metrics.fileSizeBeforeBytes,
      fileSizeAfterBytes: null,
      rsidCountBefore: ooxmlResult.metrics.rsidCountBefore,
      rsidCountAfter: null,
    },
    errors: [],
  };

  return { report, bundle: ooxmlResult.bundle, heuristics };
}
