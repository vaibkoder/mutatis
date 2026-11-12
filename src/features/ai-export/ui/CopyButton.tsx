/**
 * Copy-for-AI button pair (feature 11 — spec §6.1).
 *
 * A primary "Copy for AI" (clipboard) button, a ⬇ download button, and a
 * JSON | MD format toggle. Generic over the export it produces: the caller
 * passes a `build` thunk that returns the envelope (or null when there is no
 * data). All actions are local-only; the first one raises the awareness notice.
 */
import { useState } from 'preact/hooks';
import type { AiExportEnvelope } from '../types.js';
import { copyExport, downloadExport } from '../index.js';
import { exportFormat, maybeShowAwarenessNotice } from './prefs.js';
import { FormatToggle } from './FormatToggle.js';

export interface CopyButtonProps {
  /** Produce the export envelope, or null when the panel has no current results. */
  build: () => AiExportEnvelope | null;
  /** Slug used in the download filename, e.g. "defined-terms" or "briefing-pack". */
  fileLabel: string;
  /** Primary button label. */
  copyLabel?: string;
  /** When set, the buttons are disabled and show this tooltip (e.g. "Run Defined Terms first"). */
  disabledReason?: string;
}

export function CopyButton({ build, fileLabel, copyLabel = 'Copy for AI', disabledReason }: CopyButtonProps) {
  const [copied, setCopied] = useState(false);
  const disabled = disabledReason !== undefined;

  async function handleCopy() {
    const env = build();
    if (!env) return;
    try {
      await copyExport(env, exportFormat.value);
      maybeShowAwarenessNotice();
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      /* clipboard denied by host — non-fatal; user can use download instead */
    }
  }

  function handleDownload() {
    const env = build();
    if (!env) return;
    downloadExport(env, exportFormat.value, fileLabel);
    maybeShowAwarenessNotice();
  }

  return (
    <div class="dp-ai-export__copy" data-file={fileLabel}>
      <button
        type="button"
        class="dp-ai-export__copy-btn"
        onClick={handleCopy}
        disabled={disabled}
        title={disabledReason}
      >
        {copied ? 'Copied ✓' : `📋 ${copyLabel}`}
      </button>
      <button
        type="button"
        class="dp-ai-export__download-btn"
        onClick={handleDownload}
        disabled={disabled}
        title={disabledReason ?? 'Download as a file'}
        aria-label="Download for AI"
      >
        ⬇
      </button>
      <FormatToggle idBase={`dp-ai-fmt-${fileLabel}`} />
    </div>
  );
}
