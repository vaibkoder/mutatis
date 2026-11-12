import type { Severity } from '@shared/findings/index.js';
import type { RepairSeverity } from './types.js';

export function mapRepairSeverity(repairSeverity: RepairSeverity): Severity {
  switch (repairSeverity) {
    case 'confidentiality': return 'warning';
    case 'stabilisation':   return 'info';
    case 'info':            return 'info';
  }
}
