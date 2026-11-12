import { AlertTriangle, Bot, Check, FileText, Loader2 } from 'lucide-react';

/**
 * The two cards an editing turn puts in the chat.
 *
 * The MCQ card is the one that matters most: when the user has asked
 * for an edit without saying WHICH document, the assistant must ask with a card,
 * not a text question. A text question invites "the second one", which is a
 * guess waiting to happen; two buttons naming the actual files cannot be
 * misread.
 */

export interface DocumentNames {
  readonly original: string;
  readonly modified: string;
}

export interface ToolPart {
  readonly type: string;
  readonly toolCallId: string;
  readonly state: string;
  readonly input?: unknown;
  readonly output?: unknown;
}

export function ChooseTargetCard({
  part,
  names,
  onChoose,
}: {
  part: ToolPart;
  names: DocumentNames | null;
  onChoose: (toolCallId: string, target: 'original' | 'modified') => void;
}) {
  const question =
    typeof part.input === 'object' && part.input !== null
      ? String((part.input as Record<string, unknown>)['question'] ?? '')
      : '';
  const answered = part.state === 'output-available';
  const chosen = answered ? String((part.output as Record<string, unknown>)?.['target'] ?? '') : '';

  if (!names) {
    return (
      <Card tone="warn">
        <AlertTriangle size={12} className="mt-0.5 shrink-0 text-danger" aria-hidden />
        <p>No comparison is open, so there is no document to edit.</p>
      </Card>
    );
  }

  return (
    <Card>
      <Bot size={12} className="mt-0.5 shrink-0 text-accent" aria-hidden />
      <div className="min-w-0 flex-1">
        <p className="mb-1.5 text-[11px] leading-relaxed text-text">
          {question || 'Which document should I edit?'}
        </p>
        {answered ? (
          <p className="flex items-center gap-1 text-[11px] text-text-muted">
            <Check size={11} className="text-accent" aria-hidden />
            Editing {chosen === 'original' ? names.original : names.modified}
          </p>
        ) : (
          <div className="flex flex-col gap-1">
            <TargetButton
              label={names.original}
              hint="the original"
              onClick={() => onChoose(part.toolCallId, 'original')}
            />
            <TargetButton
              label={names.modified}
              hint="the modified"
              onClick={() => onChoose(part.toolCallId, 'modified')}
            />
          </div>
        )}
      </div>
    </Card>
  );
}

function TargetButton({
  label,
  hint,
  onClick,
}: {
  label: string;
  hint: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="flex items-center gap-1.5 rounded border border-border bg-surface px-2 py-1.5 text-left text-[11px] text-text transition-colors hover:border-accent"
    >
      <FileText size={11} className="shrink-0 text-text-faint" aria-hidden />
      <span className="min-w-0 truncate">{label}</span>
      <span className="ml-auto shrink-0 text-[10px] text-text-faint">{hint}</span>
    </button>
  );
}

export function ProposedEditsCard({ part }: { part: ToolPart }) {
  const input = (part.input ?? {}) as { target?: string; edits?: { find: string; reason?: string }[] };
  const output = part.output as
    | { applied?: number; refused?: { find: string; refusal: string }[]; error?: string }
    | undefined;
  const count = input.edits?.length ?? 0;

  if (part.state === 'input-streaming' || part.state === 'input-available') {
    return (
      <Card>
        <Loader2 size={12} className="mt-0.5 shrink-0 animate-spin text-accent" aria-hidden />
        <p>
          Applying {count > 0 ? `${count} edit${count === 1 ? '' : 's'}` : 'edits'} to the{' '}
          {input.target ?? 'document'} and redlining the result…
        </p>
      </Card>
    );
  }

  if (output?.error) {
    return (
      <Card tone="warn">
        <AlertTriangle size={12} className="mt-0.5 shrink-0 text-danger" aria-hidden />
        <p>{output.error}</p>
      </Card>
    );
  }

  const refused = output?.refused ?? [];

  return (
    <Card>
      <Bot size={12} className="mt-0.5 shrink-0 text-accent" aria-hidden />
      <div className="min-w-0 flex-1">
        <p className="text-[11px] leading-relaxed text-text">
          {output?.applied ?? 0} edit{(output?.applied ?? 0) === 1 ? '' : 's'} applied to the{' '}
          {input.target ?? 'document'}. Open the assistant&rsquo;s tab on the right to review them
          as tracked changes.
        </p>
        {refused.length > 0 && (
          <p className="mt-1 text-[11px] leading-relaxed text-danger">
            {refused.length} refused — the quoted text could not be matched exactly.
          </p>
        )}
      </div>
    </Card>
  );
}

function Card({
  children,
  tone = 'normal',
}: {
  children: React.ReactNode;
  tone?: 'normal' | 'warn';
}) {
  return (
    <div
      className={
        'flex items-start gap-1.5 rounded-md border px-2.5 py-2 text-[11px] leading-relaxed ' +
        (tone === 'warn'
          ? 'border-danger/40 bg-danger/10 text-danger'
          : 'border-border bg-surface-sunken text-text-muted')
      }
    >
      {children}
    </div>
  );
}
