// src/features/repair/apply.ts
import { wordRun } from '@shared/office/index.js';
import { applyMetadata } from './apply/metadata.js';
import { applyStandardisation } from './apply/standardise.js';
import { applySpacing } from './apply/spacing.js';
import { applyOoxmlSurgery } from './apply/ooxml-apply.js';
import { applyTracked } from './apply/tracked.js';
import type { RepairReport, RepairFinding, RepairError } from './types.js';
import type { StyleProfile } from './profiles.js';
import type { DocxBundle } from '@shared/ooxml/index.js';

export interface ApplyOptions {
  /** Which findings to act on. Findings not in this list receive result:'skipped'. */
  selectedFindingIds: string[];
  /** Non-protected section-break finding ids the user confirmed for removal. */
  confirmedSectionBreakIds: string[];
  /**
   * Orientation-changing (protected:true) section-break ids confirmed via the
   * per-break orientation-change modal. Required in addition to
   * confirmedSectionBreakIds for any protected break to be removed; absent or
   * empty means no protected break is ever removed. (spec §5.7 hard rule)
   */
  confirmedProtectedSectionBreakIds?: string[];
  blankParaMode?: 'deleteAll' | 'collapseMultiple';
  applyMargins?: boolean;
  stripInlineFormatting?: boolean;
}

/**
 * Applies selected repair findings to the document.
 *
 * Execution order (safe):
 *   1. Typed apply — metadata, standardisation, spacing, tracked changes —
 *      inside a single wordRun.
 *   2. OOXML surgery (rsid strip + section-break removal) — outside wordRun,
 *      operates directly on the bundle.
 *
 * Returns an updated RepairReport with phase:'apply' and each finding annotated
 * with its result ('done' | 'skipped' | 'error').
 *
 * @param scanReport   The RepairReport produced by scan().
 * @param bundle       The DocxBundle from ScanState — must not be re-read.
 * @param profile      Style profile used by standardisation and spacing steps.
 * @param options      Caller-controlled flags.
 * @param generatedAt  ISO 8601 timestamp for this apply-phase report.
 */
export async function apply(
  scanReport: RepairReport,
  bundle: DocxBundle,
  profile: StyleProfile,
  options: ApplyOptions,
  generatedAt: string,
): Promise<RepairReport> {
  const selected = scanReport.findings.filter(f =>
    options.selectedFindingIds.includes(f.id),
  );

  // ── Step 1: typed apply ───────────────────────────────────────────────────
  let typedResults: RepairFinding[] = [];
  const typedErrors: RepairError[] = [];

  await wordRun(async ctx => {
    const metaFindings = selected.filter(f =>
      ['clearBuiltinProps', 'deleteCustomProps', 'deleteCustomXml', 'deleteComments'].includes(f.action),
    );
    const stdFindings = selected.filter(f =>
      ['clearDirectFormatting', 'stripInlineFormatting', 'conformToProfile'].includes(f.action),
    );
    const spacingFindings = selected.filter(f =>
      ['normaliseSpacing', 'deleteBlankParas'].includes(f.action),
    );
    const trackedFindings = selected.filter(f => f.action === 'acceptAll');

    const metaDone     = await applyMetadata(ctx, metaFindings);
    const stdDone      = await applyStandardisation(ctx, stdFindings, profile, {
      ...(options.applyMargins !== undefined ? { applyMargins: options.applyMargins } : {}),
      ...(options.stripInlineFormatting !== undefined ? { stripInlineFormatting: options.stripInlineFormatting } : {}),
    });
    const spacingDone  = await applySpacing(ctx, spacingFindings, profile, {
      ...(options.blankParaMode !== undefined ? { blankParaMode: options.blankParaMode } : {}),
    });
    const trackedDone  = await applyTracked(ctx, trackedFindings, {
      ...(options.stripInlineFormatting !== undefined ? { stripInlineFormatting: options.stripInlineFormatting } : {}),
    });

    typedResults = [...metaDone, ...stdDone, ...spacingDone, ...trackedDone];

    for (const f of typedResults) {
      if (f.result === 'error') {
        typedErrors.push({
          findingId: f.id,
          code: 'GeneralException',
          message: `${f.label} failed`,
        });
      }
    }
  });

  // ── Step 2: OOXML surgery ─────────────────────────────────────────────────
  const ooxmlFindings = selected.filter(f =>
    ['stripRsid', 'removeSectionBreak'].includes(f.action),
  );
  const { findings: ooxmlResults, metricsAfter } = await applyOoxmlSurgery(
    bundle,
    ooxmlFindings,
    options.confirmedSectionBreakIds,
    options.confirmedProtectedSectionBreakIds ?? [],
  );

  const ooxmlErrors: RepairError[] = ooxmlResults
    .filter(f => f.result === 'error')
    .map(f => ({
      findingId: f.id,
      code: 'GeneralException',
      message: `${f.label} failed`,
    }));

  // ── Step 3: merge results ─────────────────────────────────────────────────
  const resultMap = new Map<string, 'done' | 'skipped' | 'error'>();
  for (const f of [...typedResults, ...ooxmlResults]) {
    if (f.result != null) {
      resultMap.set(f.id, f.result);
    }
  }

  const finalFindings: RepairFinding[] = scanReport.findings.map(f => ({
    ...f,
    result: resultMap.get(f.id) ??
      (options.selectedFindingIds.includes(f.id) ? null : 'skipped'),
  }));

  // ── Step 5: assemble apply-phase report ───────────────────────────────────
  return {
    ...scanReport,
    phase: 'apply',
    generatedAt,
    findings: finalFindings,
    metrics: {
      ...scanReport.metrics,
      fileSizeAfterBytes: metricsAfter.fileSizeAfterBytes,
      rsidCountAfter: metricsAfter.rsidCountAfter,
    },
    errors: [...typedErrors, ...ooxmlErrors],
  };
}
