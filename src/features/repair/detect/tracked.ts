// src/features/repair/detect/tracked.ts
import { isCapable } from '@shared/office/index.js';
import type { RepairFinding } from '../types.js';

/**
 * Detect tracked changes. Returns a finding only when wordApi16 is available
 * (the `getTrackedChanges()` API requires WordApi 1.6). Returns null if gated off.
 */
export async function detectTracked(ctx: Word.RequestContext): Promise<RepairFinding | null> {
  if (!isCapable('wordApi16')) return null;

  const tracked = ctx.document.body.getTrackedChanges();
  tracked.load('items');
  await ctx.sync();

  if (tracked.items.length === 0) return null;

  return {
    id: 'tracked.changes',
    category: 'tracked',
    severity: 'stabilisation',
    label: 'Tracked changes present',
    detail: `${tracked.items.length} tracked change${tracked.items.length === 1 ? '' : 's'} from one or more authors`,
    count: tracked.items.length,
    destructive: true,
    requiresConfirm: true,
    selectedByDefault: false,
    action: 'acceptAll',
  };
}
