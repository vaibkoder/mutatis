import { useState } from 'react';
import { Check, Copy, Download } from 'lucide-react';
import type { RedlineContext } from '../ai/context.js';
import { contextToJson, contextToMarkdown, downloadText } from './export-for-ai.js';

const BTN =
  'flex shrink-0 items-center gap-1.5 rounded border border-border px-2.5 py-1 text-[11px] font-medium text-text transition-colors hover:border-border-strong';

/** Copy or save the full change set as text for an AI chat. Local only. */
export function ExportForAiButton({
  context,
  baseName,
}: {
  context: RedlineContext;
  baseName: string;
}) {
  const [copied, setCopied] = useState<boolean | 'failed'>(false);
  return (
    <div className="flex items-center gap-1.5" data-ai-export>
      <button
        type="button"
        className={BTN}
        title="Copy every change as text to paste into an AI chat"
        onClick={() => {
          void (async () => {
            try {
              await navigator.clipboard.writeText(contextToMarkdown(context));
              setCopied(true);
            } catch {
              setCopied('failed');
            }
            setTimeout(() => setCopied(false), 2000);
          })();
        }}
      >
        {copied === true ? <Check size={12} aria-hidden /> : <Copy size={12} aria-hidden />}
        {copied === true ? 'Copied' : copied === 'failed' ? 'Copy blocked — use Save' : 'Copy for AI'}
      </button>
      <button
        type="button"
        className={BTN}
        title="Save the change set as a Markdown file"
        onClick={() =>
          downloadText(contextToMarkdown(context), 'text/markdown', `${baseName} (changes).md`)
        }
      >
        <Download size={12} aria-hidden />
        .md
      </button>
      <button
        type="button"
        className={BTN}
        title="Save the change set as a JSON file"
        onClick={() =>
          downloadText(contextToJson(context), 'application/json', `${baseName} (changes).json`)
        }
      >
        <Download size={12} aria-hidden />
        .json
      </button>
    </div>
  );
}
