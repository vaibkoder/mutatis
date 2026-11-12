/**
 * One-time awareness notice (feature 11 — spec §6.3, ADR-0021 principle 2).
 *
 * A non-blocking toast shown on the first copy/download action per browser
 * profile. Informational, not a consent gate — the action already proceeded.
 * Auto-dismisses after 8 s; an ✕ closes it early. Mount once in the shell.
 */
import { useEffect } from 'preact/hooks';
import { noticeVisible, dismissAwarenessNotice } from './prefs.js';

const AUTO_DISMISS_MS = 8000;

export function AwarenessNotice() {
  const visible = noticeVisible.value;

  useEffect(() => {
    if (!visible) return;
    const t = setTimeout(dismissAwarenessNotice, AUTO_DISMISS_MS);
    return () => clearTimeout(t);
  }, [visible]);

  if (!visible) return null;

  return (
    <div class="dp-ai-export__notice" role="status" aria-live="polite">
      <div class="dp-ai-export__notice-body">
        <strong>Copying document content</strong> — This copies document content to your clipboard
        for use outside Mutatis. Mutatis makes no network calls. Your firm's policies on AI and
        data handling apply.
      </div>
      <button
        type="button"
        class="dp-ai-export__notice-close"
        aria-label="Dismiss notice"
        onClick={dismissAwarenessNotice}
      >
        ✕
      </button>
    </div>
  );
}
