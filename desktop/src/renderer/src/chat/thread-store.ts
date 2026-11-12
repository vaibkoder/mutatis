import type { LoadedThread, StoredThread } from '@common/ipc.js';

/**
 * The thread list as pure data.
 *
 * Every mutation the rail can perform — open, name, rename, delete — is a
 * function from a list to a list here, with no React and no IPC. Two reasons
 * that seam is worth the file:
 *
 *   1. The renderer keeps its own copy of the list so a rename shows instantly
 *      rather than after a round trip. That copy can DRIFT from SQLite, and the
 *      only defence is that the rules for reordering live in one place and are
 *      tested against what `src/main/threads.ts` actually does.
 *   2. Ordering is the subtle half. Main sorts `ORDER BY updated_at DESC` and
 *      bumps `updated_at` on rename as well as on save, so a renderer that
 *      renames without re-sorting shows one order now and a different one after
 *      a restart.
 */

/** Titles are stored as typed but shown in a narrow rail; longer is noise. */
export const MAX_TITLE_LENGTH = 60;

/** What main writes for a thread that has no first message yet. */
export const UNTITLED = 'New conversation';

/** Most recently updated first — the order `listThreads` returns. */
export function sortThreads(threads: readonly StoredThread[]): StoredThread[] {
  return [...threads].sort((a, b) => b.updatedAt - a.updatedAt);
}

/**
 * Clean up a title the user typed.
 *
 * Returns null when there is nothing worth saving, which the caller treats as
 * "cancel", not as "store an empty name". A thread with a blank title is a row
 * the user can no longer tell apart from any other.
 */
export function sanitiseTitle(raw: string): string | null {
  const collapsed = raw.replace(/\s+/gu, ' ').trim();
  if (collapsed === '') return null;
  return collapsed.length > MAX_TITLE_LENGTH
    ? `${collapsed.slice(0, MAX_TITLE_LENGTH - 1)}…`
    : collapsed;
}

/**
 * Name a thread after its first message.
 *
 * Same cap and same ellipsis as a hand-typed rename, so an auto-named thread and
 * a renamed one cannot be told apart by their shape.
 */
export function titleFromText(text: string): string {
  return sanitiseTitle(text) ?? UNTITLED;
}

/**
 * Add a thread, or retitle one already present.
 *
 * `updatedAt` is bumped either way: this fires when a thread gets its first
 * message, which is exactly when main bumps too.
 */
export function withUpserted(
  threads: readonly StoredThread[],
  id: string,
  title: string,
  now: number,
): StoredThread[] {
  const existing = threads.find((thread) => thread.id === id);
  if (existing) {
    return sortThreads(
      threads.map((thread) => (thread.id === id ? { ...thread, title, updatedAt: now } : thread)),
    );
  }
  return sortThreads([...threads, { id, title, createdAt: now, updatedAt: now }]);
}

/**
 * Rename a thread that already exists.
 *
 * Unknown ids are returned unchanged rather than inserted — a rename is not a
 * way to conjure a thread that SQLite has never heard of, and silently creating
 * one here would show a row that vanishes on restart.
 */
export function withRenamed(
  threads: readonly StoredThread[],
  id: string,
  title: string,
  now: number,
): StoredThread[] {
  if (!threads.some((thread) => thread.id === id)) return [...threads];
  return sortThreads(
    threads.map((thread) => (thread.id === id ? { ...thread, title, updatedAt: now } : thread)),
  );
}

export function withRemoved(threads: readonly StoredThread[], id: string): StoredThread[] {
  return threads.filter((thread) => thread.id !== id);
}

/**
 * Which thread to open once `removedId` is gone.
 *
 * `threads` is the list BEFORE the removal, so the neighbour can be chosen by
 * position: the row below the deleted one, or the one above it if it was last.
 * Returns null when nothing is left, which the caller answers with a blank
 * thread rather than an empty pane.
 */
export function nextAfterRemoval(
  threads: readonly StoredThread[],
  removedId: string,
): string | null {
  const remaining = withRemoved(threads, removedId);
  if (remaining.length === 0) return null;
  const index = threads.findIndex((thread) => thread.id === removedId);
  if (index === -1) return null;
  const neighbour = remaining[Math.min(index, remaining.length - 1)];
  return neighbour ? neighbour.id : null;
}

/**
 * How many stored rows could not be shown.
 *
 * Two independent losses, summed because the user cares about the total, not
 * about which layer refused: rows main could not parse (`skipped`), and rows
 * main parsed but that are not well-formed UI messages, which only the renderer
 * can tell. Disclosing this is a project invariant — a thread that quietly shows
 * four of its six messages is worse than one that says so.
 */
export function unreadableCount(loaded: LoadedThread, usable: number): number {
  const rejectedHere = Math.max(0, loaded.messages.length - usable);
  return loaded.skipped + rejectedHere;
}

/** The sentence for `unreadableCount`, or null when there is nothing to disclose. */
export function unreadableNotice(count: number): string | null {
  if (count <= 0) return null;
  return count === 1
    ? '1 earlier message in this conversation could not be read and is not shown.'
    : `${count} earlier messages in this conversation could not be read and are not shown.`;
}
