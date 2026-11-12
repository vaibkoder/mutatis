/**
 * apps/web — the batch Compare surface (ADR-0037; spec/features/17 §5.4).
 *
 * One template against N versions. This lives in the web target only: ADR-0003's
 * two-documents-in-memory `getFileAsync` ceiling makes an N-way queue inside the
 * Word task pane structurally wrong, so the add-in has no equivalent surface and
 * is untouched by this feature.
 *
 * Two wirings, both load-bearing:
 *
 *  1. **The live driver is injected.** `BatchComparePanel` defaults to a no-op
 *     stub driver, so a panel mounted without this silently does nothing and no
 *     unit test notices. The project has already paid for that lesson once
 *     (`WebCompare.tsx:15-22`, memory: compare-pdf-driver-shell-wiring-gap), so
 *     `web-batch-compare.test.tsx` asserts the injection rather than trusting it.
 *  2. **Opening a row re-runs that one pair.** The batch drops each pair's ops
 *     and both `ParaRecord[]` as soon as it completes — retaining N of them is
 *     what breaks first (spec §4.3) — so the change list is recovered by
 *     `driver.openPair()`, which re-runs the pipeline alone. Notably NOT the
 *     single-pair driver's `runCompare()`: that materialises the redline, which
 *     on the web target means triggering a file download.
 */
import { useMemo } from 'preact/hooks';

import {
  BatchComparePanel,
  createBatchDriver,
  type BatchDriverState,
} from '@features/compare/batch/index.js';
import {
  PairDrilldown,
  usePairDrilldown,
} from '@features/compare/batch/components/PairDrilldown.js';

export function WebBatchCompare({ driver }: { driver?: BatchDriverState }) {
  // One driver for the life of the surface. Re-creating it per render would
  // drop the queued variants on the first unrelated state change.
  const state = useMemo(() => driver ?? createBatchDriver(), [driver]);

  const drilldown = usePairDrilldown(
    rowId => state.openPair(rowId),
    'That comparison is no longer available. Run the batch again to reopen it.',
  );

  return (
    <div class="dp-web__batch" aria-label="Compare one template against many documents">
      <BatchComparePanel driver={state} onOpenRow={drilldown.open} />

      <PairDrilldown state={drilldown} />
    </div>
  );
}
