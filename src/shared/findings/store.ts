import { signal, computed } from '@preact/signals';
import type { Finding, Severity, TabId } from './types.js';
import { findingToTab } from './types.js';

type FindingsMap = Map<string, Finding>;

const _findings = signal<FindingsMap>(new Map());

/** All findings as an array (memoised by Preact signals). */
export const allFindings = computed<Finding[]>(() =>
  Array.from(_findings.value.values())
);

/** Count per tab — drives badge numbers. */
export const tabCounts = computed<Readonly<Record<TabId, number>>>(() => {
  const counts: Record<string, number> = {};
  for (const f of _findings.value.values()) {
    const tab = findingToTab(f);
    counts[tab] = (counts[tab] ?? 0) + 1;
  }
  return counts as Readonly<Record<TabId, number>>;
});

/** Add or replace a finding by ID. */
export function addFinding(finding: Finding): void {
  const next = new Map(_findings.value);
  next.set(finding.id, finding);
  _findings.value = next;
}

/** Add multiple findings (replaces any with matching IDs). */
export function addFindings(findings: Finding[]): void {
  const next = new Map(_findings.value);
  for (const f of findings) next.set(f.id, f);
  _findings.value = next;
}

/** Remove a finding by ID. */
export function removeFinding(id: string): void {
  const next = new Map(_findings.value);
  next.delete(id);
  _findings.value = next;
}

/** Remove all findings whose IDs start with a given prefix. */
export function removeByPrefix(prefix: string): void {
  const next = new Map(_findings.value);
  for (const key of next.keys()) {
    if (key.startsWith(prefix)) next.delete(key);
  }
  _findings.value = next;
}

/** Clear all findings (e.g. on full re-run). */
export function clearFindings(): void {
  _findings.value = new Map();
}

/** Findings for a specific tab, sorted Critical-first. */
export function findingsForTab(tabId: TabId): Finding[] {
  const SEV_ORDER: Record<Severity, number> = { error: 0, warning: 1, info: 2 };
  return Array.from(_findings.value.values())
    .filter(f => findingToTab(f) === tabId)
    .sort((a, b) => SEV_ORDER[a.severity] - SEV_ORDER[b.severity]);
}

/** Findings by category, within a tab, sorted by severity. */
export function findingsByCategory(tabId: TabId): Map<string, Finding[]> {
  const grouped = new Map<string, Finding[]>();
  for (const f of findingsForTab(tabId)) {
    const bucket = grouped.get(f.category) ?? [];
    bucket.push(f);
    grouped.set(f.category, bucket);
  }
  return grouped;
}
