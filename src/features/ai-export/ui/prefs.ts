/**
 * AI-export UI preferences (feature 11 — spec §8).
 *
 * Persists the format toggle and the one-time awareness-notice flag in
 * `localStorage`. The awareness notice is surfaced via a shared signal so any
 * copy/download action (per-panel or briefing pack) can raise the single toast.
 */
import { signal } from '@preact/signals';
import type { ExportFormat } from '../types.js';

const FORMAT_KEY = 'ai-export.format';
const NOTICE_KEY = 'ai-export.noticeShown';

/** Reactive format pref (default markdown — spec §8). Shared across all panels. */
export const exportFormat = signal<ExportFormat>(readFormat());

/** Whether the one-time awareness toast is currently showing. */
export const noticeVisible = signal(false);

function readFormat(): ExportFormat {
  return localStorage.getItem(FORMAT_KEY) === 'json' ? 'json' : 'markdown';
}

export function setExportFormat(format: ExportFormat): void {
  exportFormat.value = format;
  localStorage.setItem(FORMAT_KEY, format);
}

/**
 * Raise the awareness notice the first time any copy/download fires for this
 * browser profile. Subsequent calls are no-ops. Never blocks the action.
 */
export function maybeShowAwarenessNotice(): void {
  if (localStorage.getItem(NOTICE_KEY) === 'true') return;
  localStorage.setItem(NOTICE_KEY, 'true');
  noticeVisible.value = true;
}

export function dismissAwarenessNotice(): void {
  noticeVisible.value = false;
}
