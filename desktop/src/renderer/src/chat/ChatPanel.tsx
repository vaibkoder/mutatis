import { useCallback, useEffect, useRef, useState } from 'react';
import type { UIMessage } from 'ai';
import { AlertTriangle, ChevronDown, ChevronUp, Plus, Table2 } from 'lucide-react';
import type {
  KeyStatus,
  ModelChoice,
  ProvidersSnapshot,
  StoredMessage,
  StoredThread,
} from '@common/ipc.js';
import { useSettings } from '../settings/SettingsProvider.js';
import {
  SUGGESTED_PROMPTS,
  buildSystemPrompt,
  narrationPrompt,
  type ContextEntry,
} from '../ai/prompt.js';
import type { DeviationMatrix } from '../ai/deviation-matrix.js';
import type { DocumentNames } from '../edits/EditToolCards.js';
import type { ProposedEdit } from '../edits/apply-edits.js';
import { ApiKeyForm } from './ApiKeyForm.js';
import { Composer } from './Composer.js';
import { MessageList } from './MessageList.js';
import { ThreadList } from './ThreadList.js';
import {
  UNTITLED,
  nextAfterRemoval,
  unreadableCount,
  unreadableNotice,
  withRemoved,
  withRenamed,
  withUpserted,
} from './thread-store.js';
import { useChatThread } from './useChatThread.js';

/**
 * The chat surface.
 *
 * It owns the thread list and which thread is open; the turn itself belongs to
 * `ChatThread`, which is keyed by thread id so that switching threads gives
 * `useChat` a clean mount rather than a stale one holding another
 * conversation's stream.
 */
export interface ChatPanelProps {
  /** Every finished comparison, in tab order. */
  readonly contexts: readonly ContextEntry[];
  /** The comparison the user is looking at, carried to the model in full. */
  readonly activeContext: ContextEntry | null;
  /**
   * Compute the deviation matrix and open it as a tab. Returns what it computed,
   * so the on-rails turn can hand the model the finished table.
   */
  readonly onDeviationMatrix: () => DeviationMatrix | null;
  /** The two documents behind the open comparison, for the MCQ card's labels. */
  readonly names: DocumentNames | null;
  /**
   * Apply the assistant's edits to one side of the open comparison. Resolves to
   * the summary that goes back to the model, so the next turn knows what was
   * refused rather than assuming it all worked.
   */
  readonly onApplyEdits: (
    target: 'original' | 'modified',
    edits: readonly ProposedEdit[],
  ) => Promise<{ applied: number; refused: { find: string; refusal: string }[]; error?: string }>;
}

export function ChatPanel({
  contexts,
  activeContext,
  onDeviationMatrix,
  names,
  onApplyEdits,
}: ChatPanelProps) {
  // The provider registry, the key status and the chosen model all live in
  // settings, shared with the Settings dialog. Fetching them again here would
  // mean the gate could still be showing after a key was saved two panels away.
  const { snapshot, keys, loadFailed, refreshKeys, model, setModel } = useSettings();
  const [threads, setThreads] = useState<StoredThread[]>([]);
  const [openThread, setOpenThread] = useState<OpenThread | null>(null);
  const [listOpen, setListOpen] = useState(false);
  /** A thread operation is in flight; the list disables its rows while it is. */
  const [pending, setPending] = useState(false);

  /**
   * The thread the user last asked for.
   *
   * Loads are IPC round trips and can finish out of order, so a click on one
   * thread followed quickly by a click on another can see the FIRST response
   * land last and win. Every resolution checks itself against this ref and drops
   * itself if it is no longer what the user is waiting for. Starting a blank
   * thread sets it too, so an in-flight load cannot overwrite a fresh one.
   */
  const wantedId = useRef<string | null>(null);

  /**
   * A one-shot system prompt for the next turn only.
   *
   * The on-rails matrix turn must reach the model with the computed table and
   * narration-only instructions, and the turn AFTER it must go back to the
   * ordinary prompt. Consuming the ref on read is what guarantees the rails do
   * not leak into the rest of the conversation.
   */
  const railsPrompt = useRef<string | null>(null);

  const others = contexts.filter((entry) => entry !== activeContext);
  const systemFor = useCallback(() => {
    const rails = railsPrompt.current;
    railsPrompt.current = null;
    return rails ?? buildSystemPrompt(activeContext, others);
  }, [activeContext, others]);

  const openBlank = useCallback(() => {
    const fresh = blankThread();
    wantedId.current = fresh.id;
    setOpenThread(fresh);
    setListOpen(false);
  }, []);

  useEffect(() => {
    let live = true;
    void (async () => {
      const list = await window.mutatis.threads.list();
      if (!live) return;
      setThreads(list);
      openBlank();
    })();
    return () => {
      live = false;
    };
  }, [openBlank]);

  const openExisting = useCallback(async (id: string) => {
    wantedId.current = id;
    setListOpen(false);
    setPending(true);
    try {
      const loaded = await window.mutatis.threads.load(id);
      // Superseded: the user has since asked for another thread, or started a
      // blank one. Landing this history now would show the wrong conversation.
      if (wantedId.current !== id) return;
      const usable = loaded.messages.map(toUiMessage).filter(isUiMessage);
      setOpenThread({
        id,
        messages: usable,
        // Rows main could not parse, PLUS rows it parsed but that are not
        // well-formed messages. Both are reported; neither is quietly dropped.
        skipped: unreadableCount(loaded, usable.length),
      });
    } finally {
      // A superseded load must not clear the flag: the load that replaced it is
      // still running and still owns it.
      if (wantedId.current === id) setPending(false);
    }
  }, []);

  const onTitled = useCallback((id: string, title: string) => {
    setThreads((current) => withUpserted(current, id, title, Date.now()));
  }, []);

  const renameThread = useCallback(async (id: string, title: string) => {
    const ok = await window.mutatis.threads.rename(id, title);
    if (!ok) {
      // Main refused the rename, so the title the renderer would show is not on
      // disk. Take the list back from SQLite rather than keep a name that would
      // vanish on restart.
      setThreads(await window.mutatis.threads.list());
      return;
    }
    // `withRenamed` re-sorts because `renameThread` bumps `updated_at`: leaving
    // the row in place would show one order now and another after a reload.
    setThreads((current) => withRenamed(current, id, title, Date.now()));
  }, []);

  const removeThread = useCallback(
    async (id: string) => {
      // The confirm happens inside ThreadList, so by here the user has agreed.
      const wasOpen = openThread?.id === id;
      setPending(true);
      try {
        await window.mutatis.threads.remove(id);
        // `nextAfterRemoval` reads the list as it was BEFORE the removal: the
        // neighbour is chosen by the deleted row's POSITION, which the trimmed
        // list no longer knows. So ask first, then trim.
        const next = nextAfterRemoval(threads, id);
        setThreads(withRemoved(threads, id));
        // Deleting a conversation the user is not in must leave their open one
        // exactly as it was.
        if (!wasOpen) return;
        if (next === null) openBlank();
        else await openExisting(next);
      } finally {
        setPending(false);
      }
    },
    [openBlank, openExisting, openThread, threads],
  );

  if (loadFailed) {
    return <Notice>The assistant could not start. Restart Mutatis and try again.</Notice>;
  }
  if (!snapshot || !keys || !openThread) {
    return <Notice>Loading…</Notice>;
  }

  const provider = snapshot.providers.find((entry) => entry.id === model.provider);
  if (!provider) {
    return <Notice>No provider is configured for this build.</Notice>;
  }

  // The gate appears when the SELECTED provider has no key, not when no provider
  // has one — picking Qwen with only a DeepSeek key must ask for a Qwen key
  // rather than silently answering from the other model.
  if (!keys[model.provider].configured) {
    return (
      <ApiKeyForm
        provider={provider}
        encryptionAvailable={snapshot.encryptionAvailable}
        onSaved={() => void refreshKeys()}
      />
    );
  }

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <ThreadBar
        threads={threads}
        openId={openThread.id}
        snapshot={snapshot}
        keys={keys}
        model={model}
        listOpen={listOpen}
        onToggleList={() => setListOpen((open) => !open)}
        onNew={openBlank}
        onModel={setModel}
      />
      {listOpen && (
        // Bounded and scrolled in its own box: an unbounded list would push the
        // composer off the bottom of the rail once a few conversations exist.
        <div className="max-h-64 shrink-0 overflow-y-auto border-b border-border">
          <ThreadList
            threads={threads}
            openId={openThread.id}
            unsaved={!threads.some((thread) => thread.id === openThread.id)}
            onOpen={(id) => void openExisting(id)}
            onNew={openBlank}
            onRename={(id, title) => void renameThread(id, title)}
            onDelete={(id) => void removeThread(id)}
            // Only a thread operation freezes the rows. Switching DURING a turn
            // is deliberately allowed, and discards the half-streamed answer:
            // a deliberate product decision. `ChatThread` cancels the turn as it
            // unmounts so the drop is a cancellation, not a leak.
            busy={pending}
          />
        </div>
      )}
      <ChatThread
        key={openThread.id}
        threadId={openThread.id}
        initialMessages={openThread.messages}
        skipped={openThread.skipped}
        model={model}
        system={systemFor}
        hasComparison={activeContext !== null}
        comparisons={contexts.length}
        onRails={(id) => {
          if (id !== 'matrix') return true;
          const matrix = onDeviationMatrix();
          if (!matrix) return false;
          railsPrompt.current = narrationPrompt(matrix);
          return true;
        }}
        names={names}
        onApplyEdits={onApplyEdits}
        onTitled={onTitled}
      />
    </div>
  );
}

interface OpenThread {
  readonly id: string;
  readonly messages: UIMessage[];
  readonly skipped: number;
}

function blankThread(): OpenThread {
  return { id: crypto.randomUUID(), messages: [], skipped: 0 };
}

function ChatThread({
  threadId,
  initialMessages,
  skipped,
  model,
  system,
  hasComparison,
  comparisons,
  onRails,
  names,
  onApplyEdits,
  onTitled,
}: {
  threadId: string;
  initialMessages: UIMessage[];
  skipped: number;
  model: ModelChoice;
  system: () => string;
  hasComparison: boolean;
  comparisons: number;
  /** Prepare an on-rails turn. Returns false when it cannot run. */
  onRails: (id: string) => boolean;
  names: DocumentNames | null;
  onApplyEdits: ChatPanelProps['onApplyEdits'];
  onTitled: (id: string, title: string) => void;
}) {
  const { messages, send, stop, error, clearError, addToolResult, busy } = useChatThread({
    threadId,
    initialMessages,
    model,
    system,
    onToolCall: (call) => {
      // chooseEditTarget is deliberately NOT answered here. It is a question for
      // the user, and answering it automatically would be the app guessing which
      // document to edit — the exact thing the MCQ card exists to prevent.
      if (call.toolName !== 'proposeTrackedEdits') return;

      const input = call.input as { target?: unknown; edits?: unknown };
      const target = input.target === 'original' || input.target === 'modified' ? input.target : null;
      const edits = Array.isArray(input.edits) ? (input.edits as ProposedEdit[]) : [];

      if (!target || edits.length === 0) {
        void addToolResult({
          tool: 'proposeTrackedEdits',
          toolCallId: call.toolCallId,
          output: {
            applied: 0,
            refused: [],
            error: 'The tool call was malformed: it needs a target and at least one edit.',
          },
        });
        return;
      }

      void onApplyEdits(target, edits).then((summary: Awaited<ReturnType<ChatPanelProps['onApplyEdits']>>) =>
        addToolResult({
          tool: 'proposeTrackedEdits',
          toolCallId: call.toolCallId,
          output: summary,
        }),
      );
    },
    onTitled,
  });

  /**
   * Cancel the turn when this thread goes away.
   *
   * Switching conversations mid-stream is allowed and discards the partial
   * answer — but `useChat` does NOT abort on unmount (its only effect is for
   * `resume`), so without this the MessagePort would stay open and main would
   * keep streaming the provider's reply to completion with nothing reading it.
   * The dropped answer is a choice; a provider request still billing after the
   * reader is gone is just a leak. `stop()` closes the port, which is what
   * actually aborts upstream — see transport.ts.
   *
   * This also fires when the rail is collapsed, since that unmounts the panel.
   * That already discarded the visible turn before this existed, so cancelling
   * is the honest version of what was happening anyway.
   */
  const stopRef = useRef(stop);
  stopRef.current = stop;
  useEffect(
    () => () => {
      // `stop` resolves once the SDK has settled the stream; a cleanup cannot
      // return a promise, and there is nothing left here to await it anyway.
      void stopRef.current();
    },
    [],
  );

  const chooseTarget = (toolCallId: string, target: 'original' | 'modified') => {
    void addToolResult({ tool: 'chooseEditTarget', toolCallId, output: { target } });
  };

  const runSuggestion = (promptId: string, text: string) => {
    const suggestion = SUGGESTED_PROMPTS.find((entry) => entry.id === promptId);
    if (suggestion?.onRails && !onRails(promptId)) return;
    send(text);
  };

  // A project invariant: what could not be read is always said out loud. The
  // sentence comes from `unreadableNotice` so the rail and any other surface
  // that has to disclose this cannot word it two different ways.
  const notice = unreadableNotice(skipped);

  return (
    <>
      {notice !== null && (
        <p className="flex items-start gap-1.5 border-b border-border bg-danger/10 px-3 py-1.5 text-[11px] leading-relaxed text-danger">
          <AlertTriangle size={12} className="mt-0.5 shrink-0" aria-hidden />
          {notice}
        </p>
      )}

      {messages.length === 0 ? (
        <div className="flex flex-1 flex-col justify-center gap-2 px-4">
          <p className="text-center text-[12px] text-text-muted">
            {hasComparison ? 'Ask about the open comparison.' : 'Compare two documents to begin.'}
          </p>
          <p className="text-center text-[11px] leading-relaxed text-text-faint">
            Sending a message shares this comparison&rsquo;s file names and the before and after
            wording of its changed paragraphs with your chosen AI provider, using your key.
            Nothing is sent until you do. Every edit the assistant proposes lands as a tracked
            change you review.
          </p>
          {hasComparison && (
            <div className="mt-2 flex flex-col gap-1.5">
              {SUGGESTED_PROMPTS.map((suggestion) => {
                // The matrix needs something to compare; one comparison is a
                // table with one column, which tells the user nothing.
                const ready = suggestion.onRails ? comparisons >= 2 : true;
                return (
                  <button
                    key={suggestion.id}
                    type="button"
                    disabled={!ready || busy}
                    onClick={() => runSuggestion(suggestion.id, suggestion.text)}
                    title={
                      ready
                        ? suggestion.text
                        : 'Open at least two comparisons to build a deviation matrix.'
                    }
                    className="flex items-center gap-1.5 rounded border border-border px-2.5 py-1.5 text-left text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text disabled:opacity-40"
                  >
                    {suggestion.onRails && (
                      <Table2 size={11} className="shrink-0 text-accent" aria-hidden />
                    )}
                    {suggestion.label}
                    {suggestion.onRails && (
                      <span className="ml-auto shrink-0 text-[10px] text-text-faint">computed</span>
                    )}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      ) : (
        <MessageList
          messages={messages}
          busy={busy}
          names={names}
          onChooseTarget={chooseTarget}
        />
      )}

      {error && (
        <button
          type="button"
          onClick={clearError}
          title="Dismiss"
          className="flex items-start gap-1.5 border-t border-border bg-danger/10 px-3 py-1.5 text-left text-[11px] leading-relaxed text-danger"
        >
          <AlertTriangle size={12} className="mt-0.5 shrink-0" aria-hidden />
          {error.message}
        </button>
      )}

      {/*
        The on-rails action stays reachable after the first message. The other
        suggestions do not: once a conversation is under way, a row of canned
        prompts is clutter, but the matrix is a computed artefact the user may
        want at any point.
      */}
      {messages.length > 0 && comparisons >= 2 && (
        <div className="shrink-0 border-t border-border px-2 pt-1.5">
          <button
            type="button"
            disabled={busy}
            onClick={() => {
              const matrix = SUGGESTED_PROMPTS.find((entry) => entry.id === 'matrix');
              if (matrix) runSuggestion(matrix.id, matrix.text);
            }}
            className="flex items-center gap-1.5 rounded border border-border px-2 py-1 text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text disabled:opacity-40"
          >
            <Table2 size={11} className="text-accent" aria-hidden />
            Deviation matrix
            <span className="text-[10px] text-text-faint">computed</span>
          </button>
        </div>
      )}

      <Composer busy={busy} disabled={false} onSend={send} onStop={stop} />
    </>
  );
}

function ThreadBar({
  threads,
  openId,
  snapshot,
  keys,
  model,
  listOpen,
  onToggleList,
  onNew,
  onModel,
}: {
  threads: readonly StoredThread[];
  openId: string;
  snapshot: ProvidersSnapshot;
  keys: KeyStatus;
  model: ModelChoice;
  listOpen: boolean;
  onToggleList: () => void;
  onNew: () => void;
  onModel: (choice: ModelChoice) => void;
}) {
  // A thread the list has never heard of is one with no messages yet, which is
  // what main would call it once it has any.
  const title = threads.find((thread) => thread.id === openId)?.title ?? UNTITLED;

  return (
    <div className="flex shrink-0 items-center gap-1.5 border-b border-border px-2 py-1.5">
      {/*
        Never disabled, even with no conversations stored: the empty state lives
        inside the list, and a toggle that cannot be pressed hides it.
      */}
      <button
        type="button"
        onClick={onToggleList}
        aria-expanded={listOpen}
        aria-label={`Conversations: ${title}`}
        title="Conversations"
        className="flex min-w-0 flex-1 items-center gap-1.5 rounded border border-border bg-surface px-1.5 py-1 text-left text-[11px] text-text transition-colors hover:border-border-strong"
      >
        <span className="truncate">{title}</span>
        {listOpen ? (
          <ChevronUp size={12} className="ml-auto shrink-0 text-text-muted" aria-hidden />
        ) : (
          <ChevronDown size={12} className="ml-auto shrink-0 text-text-muted" aria-hidden />
        )}
      </button>

      <select
        value={`${model.provider}/${model.model}`}
        onChange={(event) => {
          const [providerId, ...rest] = event.target.value.split('/');
          if (providerId === 'deepseek' || providerId === 'alibaba') {
            onModel({ provider: providerId, model: rest.join('/') });
          }
        }}
        aria-label="Model"
        className="min-w-0 max-w-[9rem] rounded border border-border bg-surface px-1.5 py-1 text-[11px] text-text outline-none focus:border-focus"
      >
        {snapshot.providers.map((entry) => (
          <optgroup
            key={entry.id}
            label={keys[entry.id].configured ? entry.label : `${entry.label} — no key`}
          >
            {entry.models.map((option) => (
              <option key={option.id} value={`${entry.id}/${option.id}`}>
                {option.label}
              </option>
            ))}
          </optgroup>
        ))}
      </select>

      <IconButton label="New conversation" onClick={onNew}>
        <Plus size={13} aria-hidden />
      </IconButton>
    </div>
  );
}

function IconButton({
  label,
  onClick,
  children,
}: {
  label: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-label={label}
      title={label}
      className="flex h-6 w-6 shrink-0 items-center justify-center rounded text-text-muted transition-colors hover:bg-border hover:text-text"
    >
      {children}
    </button>
  );
}

function Notice({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex flex-1 items-center justify-center px-6 text-center">
      <p className="text-[11px] leading-relaxed text-text-muted">{children}</p>
    </div>
  );
}

/**
 * A stored row is untrusted input like any other: it was written by an earlier
 * version of this app, and it may have been edited on disk. Anything that is not
 * a well-formed message is rejected here and counted, never coerced into one.
 */
function toUiMessage(row: StoredMessage): UIMessage | null {
  if (!Array.isArray(row.parts)) return null;
  return {
    id: row.id,
    role: row.role,
    parts: row.parts,
  } as UIMessage;
}

function isUiMessage(message: UIMessage | null): message is UIMessage {
  return message !== null;
}
