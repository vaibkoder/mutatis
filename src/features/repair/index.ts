// src/features/repair/index.ts
// Barrel for Feature 05 — Document Repair.

export { RepairPanel } from './RepairPanel.js';
export {
  repairProfile,
  repairScanReport,
  repairApplyReport,
  repairBusy,
  repairError,
  repairChecked,
  repairGroupedFindings,
  repairNormalisationWarning,
  repairCommercialOnCourtWarning,
  repairConfirmedSectionBreaks,
  repairConfirmedProtectedBreaks,
  setProfile,
  runScan,
  runApply,
  toggleFinding,
  confirmSectionBreak,
  confirmProtectedBreak,
} from './driver.js';
