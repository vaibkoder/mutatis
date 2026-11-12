import type { Severity } from '@shared/findings/index.js';

export type RepairSeverity = 'info' | 'stabilisation' | 'confidentiality';

export type RepairCategory =
  | 'metadata'
  | 'comments'
  | 'customXml'
  | 'styles'
  | 'spacing'
  | 'blankParas'
  | 'tracked'
  | 'sectionBreaks'
  | 'rsid';

export interface RepairFinding {
  id: string;
  category: RepairCategory;
  severity: RepairSeverity;
  label: string;
  detail?: string;
  count?: number;
  destructive: boolean;
  requiresConfirm: boolean;
  selectedByDefault: boolean;
  /** Orientation-changing section breaks must have per-break confirmation. */
  protected?: boolean;
  action: string;
  /** Populated in apply phase only. */
  result?: 'done' | 'skipped' | 'error' | null;
}

export interface RepairMetrics {
  fileSizeBeforeBytes: number | null;
  fileSizeAfterBytes: number | null;
  rsidCountBefore: number | null;
  rsidCountAfter: number | null;
}

export interface RepairError {
  findingId: string;
  code: 'GeneralException' | 'AccessDenied' | 'InvalidArgument' | string;
  message: string;
}

export interface RepairReport {
  documentName: string;
  /** ISO 8601 timestamp. */
  generatedAt: string;
  profile: { id: string; name: string };
  phase: 'scan' | 'apply';
  findings: RepairFinding[];
  metrics: RepairMetrics;
  errors: RepairError[];
}

/** Re-export the shared Severity type for convenience within the repair feature. */
export type { Severity };
