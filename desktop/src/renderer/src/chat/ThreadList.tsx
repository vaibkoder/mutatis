import { useEffect, useRef, useState } from 'react';
import { Check, MessageSquare, Pencil, Plus, Trash2, X } from 'lucide-react';
import type { StoredThread } from '@common/ipc.js';
import { sanitiseTitle } from './thread-store.js';

/**
 * The saved-conversation list for the chat rail.
 *
 * Presentation only: props in, callbacks out, no IPC and no knowledge of the
 * bridge. Everything durable — the list itself, which thread is open, what
 * reaches SQLite — belongs to the caller. What lives here is the handful of
 * states that are pure interaction and would be noise anywhere else: which row
 * is being renamed, the half-typed title, and which row has asked "are you
 * sure?". None of those survive a remount, and none of them should.
 */
export interface ThreadListProps {
  /** Saved conversations, already sorted most-recent-first by the caller. */
  readonly threads: readonly StoredThread[];
  /** The conversation currently open. May not be present in `threads`. */
  readonly openId: string;
  /** True when the open conversation has no messages yet and so is not saved. */
  readonly unsaved: boolean;
  readonly onOpen: (id: string) => void;
  readonly onNew: () => void;
  /** Commit a rename. Only ever called with an already-sanitised, non-empty title. */
  readonly onRename: (id: string, title: string) => void;
  readonly onDelete: (id: string) => void;
  /** A turn is streaming. Switching and deleting are disabled while it is true. */
  readonly busy?: boolean;
}

/**
 * The row being renamed.
 *
 * `original` is carried rather than looked up again at commit time, so the
 * "did it actually change?" test compares the draft against the title the user
 * started editing — not against a list that may have been re-sorted or
 * re-fetched underneath us mid-edit.
 */
interface Rename {
  readonly id: string;
  readonly original: string;
  readonly draft: string;
}

export function ThreadList({
  threads,
  openId,
  unsaved,
  onOpen,
  onNew,
  onRename,
  onDelete,
  busy = false,
}: ThreadListProps): React.ReactElement {
  const [rename, setRename] = useState<Rename | null>(null);
  const [confirming, setConfirming] = useState<string | null>(null);

  /**
   * Enter and blur both commit, and Enter unmounts the input — so the same edit
   * can be asked to finish twice within one event. This latch makes the second
   * ask a no-op; without it a rename could fire `onRename` twice, or commit
   * after an Escape had already cancelled it.
   */
  const settled = useRef(false);
  const inputRef = useRef<HTMLInputElement | null>(null);

  const renamingId = rename?.id ?? null;
  useEffect(() => {
    if (renamingId === null) return;
    const input = inputRef.current;
    if (!input) return;
    input.focus();
    input.select();
  }, [renamingId]);

  // Renaming and confirming a delete are mutually exclusive: one row asking a
  // question while another is mid-edit reads as two half-finished actions.
  function startRename(thread: StoredThread) {
    settled.current = false;
    setConfirming(null);
    setRename({ id: thread.id, original: thread.title, draft: thread.title });
  }

  function finishRename(commit: boolean) {
    if (settled.current || !rename) return;
    settled.current = true;
    const title = commit ? sanitiseTitle(rename.draft) : null;
    // A blank draft and an untouched one are both silent cancels. `onRename` is
    // documented to receive only a title worth storing, and a thread with an
    // empty name is a row the user can no longer tell from any other.
    if (title !== null && title !== rename.original) onRename(rename.id, title);
    setRename(null);
  }

  function startConfirm(id: string) {
    setRename(null);
    settled.current = true;
    setConfirming(id);
  }

  const now = Date.now();

  // Deliberately NOT role="listbox" with role="option" rows: an option may not
  // contain interactive children, and every row here carries a rename and a
  // delete control. A plain list of buttons with aria-current is valid and
  // navigates the same way.
  return (
    <div className="flex min-h-0 flex-col gap-1.5 p-2">
      <button
        type="button"
        onClick={onNew}
        aria-label="New conversation"
        title="New conversation"
        className="flex shrink-0 items-center gap-1.5 rounded border border-border px-2 py-1 text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text"
      >
        <Plus size={12} className="shrink-0" aria-hidden />
        New conversation
      </button>

      {threads.length === 0 ? (
        <p className="px-0.5 py-1 text-[11px] leading-relaxed text-text-faint">
          No saved conversations yet. The one on screen is saved as soon as you send the first
          message.
        </p>
      ) : (
        <ul className="base-ui-disable-scrollbar m-0 flex min-h-0 list-none flex-col gap-0.5 overflow-y-auto p-0">
          {/*
            The not-yet-saved conversation, so the list is never showing a set of
            rows with nothing selected while the user is plainly inside a
            conversation. It is not a button: there is nowhere to navigate to,
            and nothing on disk to rename or delete yet.
          */}
          {unsaved && (
            <li
              aria-current="true"
              className="flex items-center gap-1.5 rounded border border-accent/40 bg-accent-soft px-2 py-1 text-[11px] text-text"
              title="This conversation is saved once you send the first message"
            >
              <MessageSquare size={11} className="shrink-0 text-accent" aria-hidden />
              <span className="truncate">New conversation</span>
              <span className="ml-auto shrink-0 text-[10px] text-text-faint">unsaved</span>
            </li>
          )}

          {threads.map((thread) => {
            const open = thread.id === openId;
            const editing = rename?.id === thread.id;
            const asking = confirming === thread.id;

            return (
              <li key={thread.id} className="flex flex-col">
                <div
                  className={`flex items-center gap-1 rounded border px-1.5 py-1 ${
                    open ? 'border-accent/40 bg-accent-soft' : 'border-transparent'
                  }`}
                >
                  {editing && rename ? (
                    <input
                      ref={inputRef}
                      value={rename.draft}
                      onChange={(event) => setRename({ ...rename, draft: event.target.value })}
                      onBlur={() => finishRename(true)}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
                          event.preventDefault();
                          finishRename(true);
                        } else if (event.key === 'Escape') {
                          event.preventDefault();
                          finishRename(false);
                        }
                      }}
                      aria-label="Conversation title"
                      title="Conversation title"
                      autoComplete="off"
                      spellCheck={false}
                      className="min-w-0 flex-1 rounded border border-border bg-surface px-1.5 py-0.5 text-[11px] text-text outline-none focus:border-focus"
                    />
                  ) : (
                    <>
                      <button
                        type="button"
                        disabled={busy}
                        aria-current={open ? 'true' : undefined}
                        // Reopening the open thread would remount the turn for
                        // nothing — and throw away a stream in progress.
                        onClick={() => {
                          if (!open) onOpen(thread.id);
                        }}
                        title={`${thread.title} — last used ${new Date(thread.updatedAt).toLocaleString()}`}
                        className={`flex min-w-0 flex-1 items-center gap-1.5 rounded px-0.5 py-0.5 text-left text-[11px] transition-colors disabled:opacity-40 ${
                          open ? 'text-text' : 'text-text-muted hover:text-text'
                        }`}
                      >
                        <MessageSquare
                          size={11}
                          className={`shrink-0 ${open ? 'text-accent' : 'text-text-faint'}`}
                          aria-hidden
                        />
                        <span className="truncate">{thread.title}</span>
                        <span className="ml-auto shrink-0 pl-1 text-[10px] text-text-faint">
                          {timeHint(thread.updatedAt, now)}
                        </span>
                      </button>

                      {!asking && (
                        <>
                          <RowButton
                            label={`Rename “${thread.title}”`}
                            onClick={() => startRename(thread)}
                          >
                            <Pencil size={11} aria-hidden />
                          </RowButton>
                          <RowButton
                            label={`Delete “${thread.title}”`}
                            disabled={busy}
                            onClick={() => startConfirm(thread.id)}
                          >
                            <Trash2 size={11} aria-hidden />
                          </RowButton>
                        </>
                      )}
                    </>
                  )}
                </div>

                {/*
                  Two steps, not one, and not `window.confirm`: a conversation is
                  the only copy of its own history, a misclick is unrecoverable,
                  and a blocking native dialog cannot be driven by a test.
                */}
                {asking && (
                  <div
                    onKeyDown={(event) => {
                      if (event.key === 'Escape') {
                        event.preventDefault();
                        setConfirming(null);
                      }
                    }}
                    className="mt-0.5 mb-0.5 flex items-center gap-1 rounded border border-border bg-surface px-1.5 py-1"
                  >
                    <span className="min-w-0 flex-1 truncate text-[11px] text-text-muted">
                      Delete this conversation?
                    </span>
                    <button
                      type="button"
                      disabled={busy}
                      onClick={() => {
                        setConfirming(null);
                        onDelete(thread.id);
                      }}
                      aria-label={`Delete “${thread.title}” permanently`}
                      title={`Delete “${thread.title}” permanently`}
                      className="flex shrink-0 items-center gap-1 rounded px-1 py-0.5 text-[11px] text-danger transition-colors hover:bg-danger/10 disabled:opacity-40"
                    >
                      <Check size={11} aria-hidden />
                      Delete
                    </button>
                    <button
                      type="button"
                      onClick={() => setConfirming(null)}
                      aria-label="Keep this conversation"
                      title="Keep this conversation"
                      className="flex shrink-0 items-center gap-1 rounded px-1 py-0.5 text-[11px] text-text-muted transition-colors hover:bg-border hover:text-text"
                    >
                      <X size={11} aria-hidden />
                      Cancel
                    </button>
                  </div>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}

function RowButton({
  label,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  disabled?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onClick}
      aria-label={label}
      title={label}
      className="flex h-5 w-5 shrink-0 items-center justify-center rounded text-text-faint transition-colors hover:bg-border hover:text-text disabled:opacity-40"
    >
      {children}
    </button>
  );
}

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;

/**
 * A terse age for the row's right edge.
 *
 * The rail is narrow, so this is an initial, not a phrase; the full timestamp
 * is in the row's `title`. `now` is passed in so the whole list agrees on one
 * reading of the clock, and a stored time in the future — clock skew, or a file
 * copied between machines — reads as "now" rather than as a negative age.
 */
function timeHint(updatedAt: number, now: number): string {
  const age = now - updatedAt;
  if (age < MINUTE) return 'now';
  if (age < HOUR) return `${Math.floor(age / MINUTE)}m`;
  if (age < DAY) return `${Math.floor(age / HOUR)}h`;
  if (age < 7 * DAY) return `${Math.floor(age / DAY)}d`;
  return new Date(updatedAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}
