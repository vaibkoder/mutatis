import { describe, it, expect } from 'vitest';
import type { LoadedThread, StoredMessage, StoredThread } from '../src/common/ipc.js';
import {
  MAX_TITLE_LENGTH,
  UNTITLED,
  nextAfterRemoval,
  sanitiseTitle,
  sortThreads,
  titleFromText,
  unreadableCount,
  unreadableNotice,
  withRemoved,
  withRenamed,
  withUpserted,
} from '../src/renderer/src/chat/thread-store.js';

/**
 * The thread rail keeps its own copy of the list so a rename or a delete shows
 * instantly rather than after an IPC round trip. That copy can DRIFT from
 * SQLite, and the only defence is that every rule for reordering lives in this
 * one module and is tested against what `src/main/threads.ts` actually does.
 *
 * So these tests are not about the happy path. They pin the cases where the
 * renderer could show the user something SQLite will not agree with after a
 * restart: an order that changes on reload, a phantom row a rename conjured, a
 * message count that quietly excludes rows nobody could parse.
 */

function thread(id: string, updatedAt: number, title: string = id): StoredThread {
  return { id, title, createdAt: 1_000, updatedAt };
}

function message(id: string): StoredMessage {
  return { id, threadId: 't1', role: 'user', parts: [{ type: 'text', text: id }], createdAt: 1 };
}

function loadedWith(messageCount: number, skipped: number): LoadedThread {
  return {
    messages: Array.from({ length: messageCount }, (_, i) => message(`m${i}`)),
    skipped,
  };
}

/** The ids of a list, which is all these assertions ever care about. */
function ids(threads: readonly StoredThread[]): string[] {
  return threads.map((t) => t.id);
}

describe('sortThreads', () => {
  it('puts the most recently updated first, matching ORDER BY updated_at DESC', () => {
    const sorted = sortThreads([thread('a', 100), thread('b', 300), thread('c', 200)]);
    expect(ids(sorted)).toEqual(['b', 'c', 'a']);
  });

  // The rail renders from React state. Sorting in place would mutate the array
  // React is holding, so the reorder could land without a re-render — the list
  // would be right in memory and wrong on screen until something else changed.
  it('does not mutate its input', () => {
    const input = [thread('a', 100), thread('b', 300)];
    const sorted = sortThreads(input);
    expect(ids(input)).toEqual(['a', 'b']);
    expect(sorted).not.toBe(input);
  });
});

describe('sanitiseTitle', () => {
  it('trims and collapses runs of whitespace', () => {
    expect(sanitiseTitle('  Master   Services\n\nAgreement  ')).toBe('Master Services Agreement');
    expect(sanitiseTitle('a\t\tb')).toBe('a b');
  });

  it('leaves a short, already-clean title untouched', () => {
    expect(sanitiseTitle('Lease, clause 4.1')).toBe('Lease, clause 4.1');
  });

  // Null is the caller's signal for "cancel the rename". A thread stored with a
  // blank title is a row the user can no longer tell apart from any other, and
  // there is no way back from the rail once its name is gone.
  it('returns null for anything that carries no name', () => {
    expect(sanitiseTitle('')).toBeNull();
    expect(sanitiseTitle('   ')).toBeNull();
    expect(sanitiseTitle('\t\n')).toBeNull();
    expect(sanitiseTitle('\n \t \r\n')).toBeNull();
  });

  it('caps at MAX_TITLE_LENGTH with an ellipsis, so the result is never longer', () => {
    const long = 'a'.repeat(200);
    const capped = sanitiseTitle(long);
    expect(capped).not.toBeNull();
    expect(capped?.length).toBe(MAX_TITLE_LENGTH);
    expect(capped?.endsWith('…')).toBe(true);
    expect(capped?.startsWith('a'.repeat(MAX_TITLE_LENGTH - 1))).toBe(true);
  });

  // The cap has to apply AFTER collapsing, or a title padded with whitespace
  // gets truncated on length it does not really have.
  it('measures the cap against the collapsed text, not the raw text', () => {
    const padded = `${' '.repeat(100)}short title${' '.repeat(100)}`;
    expect(sanitiseTitle(padded)).toBe('short title');
  });

  it('keeps a title that is exactly at the cap intact, with no ellipsis', () => {
    const exact = 'b'.repeat(MAX_TITLE_LENGTH);
    expect(sanitiseTitle(exact)).toBe(exact);
  });
});

describe('titleFromText', () => {
  it('names a thread from its first message', () => {
    expect(titleFromText('  What changed in  clause 7?  ')).toBe('What changed in clause 7?');
  });

  // Same cap and same ellipsis as a hand-typed rename: an auto-named thread and
  // a renamed one must not be tellable apart by their shape.
  it('applies the same cap as a hand-typed rename', () => {
    const long = 'c'.repeat(200);
    expect(titleFromText(long)).toBe(sanitiseTitle(long));
    expect(titleFromText(long).length).toBe(MAX_TITLE_LENGTH);
  });

  // A first message that is only whitespace — an attachment-only turn, say —
  // must still produce the name main writes, not an empty rail entry.
  it('falls back to UNTITLED for text that sanitises to null', () => {
    expect(titleFromText('')).toBe(UNTITLED);
    expect(titleFromText('   \n\t ')).toBe(UNTITLED);
  });
});

describe('withUpserted', () => {
  it('inserts a thread it has never seen', () => {
    const next = withUpserted([thread('a', 100)], 'b', 'Second', 300);
    expect(ids(next)).toEqual(['b', 'a']);
    expect(next.find((t) => t.id === 'b')?.title).toBe('Second');
  });

  // This fires when a thread gets its first message and is renamed from
  // "New conversation" to that message's text. An insert here instead of an
  // update would show the same conversation twice until the next reload.
  it('retitles a known thread without duplicating it', () => {
    const next = withUpserted([thread('a', 100, 'New conversation'), thread('b', 50)], 'a', 'Named', 300);
    expect(next).toHaveLength(2);
    expect(next.filter((t) => t.id === 'a')).toHaveLength(1);
    expect(next.find((t) => t.id === 'a')?.title).toBe('Named');
  });

  it('bumps updatedAt and so moves the row to the top, as main does', () => {
    const next = withUpserted([thread('a', 100), thread('b', 300)], 'a', 'Named', 900);
    expect(ids(next)).toEqual(['a', 'b']);
    expect(next[0]?.updatedAt).toBe(900);
  });

  it('preserves createdAt when it retitles, so the row does not look new', () => {
    const next = withUpserted([{ id: 'a', title: 'x', createdAt: 42, updatedAt: 100 }], 'a', 'y', 900);
    expect(next[0]?.createdAt).toBe(42);
  });

  it('never mutates the input array', () => {
    const input = [thread('a', 100), thread('b', 300)];
    withUpserted(input, 'a', 'Named', 900);
    withUpserted(input, 'c', 'New', 900);
    expect(input).toHaveLength(2);
    expect(ids(input)).toEqual(['a', 'b']);
    expect(input[0]?.title).toBe('a');
  });
});

describe('withRenamed', () => {
  // Main bumps `updated_at` on rename as well as on save. A renderer that
  // renames without re-sorting shows one order now and a different one after a
  // restart, which reads as the app having lost track of the list.
  it('renames in place and re-sorts, because main bumps updated_at too', () => {
    const next = withRenamed([thread('a', 100), thread('b', 300)], 'a', 'Renamed', 900);
    expect(ids(next)).toEqual(['a', 'b']);
    expect(next[0]?.title).toBe('Renamed');
    expect(next[0]?.updatedAt).toBe(900);
  });

  // The interesting case. A rename is not a way to conjure a thread SQLite has
  // never heard of: main's UPDATE matches no row, so inserting one here would
  // put a phantom in the rail that vanishes at the next restart.
  it('returns the list unchanged for an unknown id rather than inserting a phantom', () => {
    const input = [thread('a', 100), thread('b', 300)];
    const next = withRenamed(input, 'ghost', 'Phantom', 900);
    expect(next).toHaveLength(2);
    expect(ids(next)).toEqual(['a', 'b']);
    expect(next.some((t) => t.id === 'ghost')).toBe(false);
  });

  it('does not mutate the input array', () => {
    const input = [thread('a', 100), thread('b', 300)];
    withRenamed(input, 'a', 'Renamed', 900);
    expect(input[0]?.title).toBe('a');
    expect(input[0]?.updatedAt).toBe(100);
  });
});

describe('withRemoved', () => {
  it('removes the id', () => {
    const next = withRemoved([thread('a', 300), thread('b', 200), thread('c', 100)], 'b');
    expect(ids(next)).toEqual(['a', 'c']);
  });

  // A delete whose IPC call lost the race, or a double-click on the confirm
  // button, must not drop an unrelated row.
  it('is a no-op for an unknown id', () => {
    const next = withRemoved([thread('a', 300), thread('b', 200)], 'ghost');
    expect(ids(next)).toEqual(['a', 'b']);
  });

  it('does not mutate the input array', () => {
    const input = [thread('a', 300), thread('b', 200)];
    withRemoved(input, 'a');
    expect(ids(input)).toEqual(['a', 'b']);
  });
});

describe('nextAfterRemoval', () => {
  // It takes the list as it was BEFORE the removal, so the neighbour can be
  // chosen by position rather than by guessing at timestamps.
  it('picks the row below the removed one', () => {
    const before = [thread('a', 300), thread('b', 200), thread('c', 100)];
    expect(nextAfterRemoval(before, 'a')).toBe('b');
    expect(nextAfterRemoval(before, 'b')).toBe('c');
  });

  it('picks the row above when the removed one was last', () => {
    const before = [thread('a', 300), thread('b', 200), thread('c', 100)];
    expect(nextAfterRemoval(before, 'c')).toBe('b');
  });

  // Null, not an id and not a throw: the caller answers it with a blank thread
  // rather than an empty pane with nothing to type into.
  it('returns null when it removed the only thread', () => {
    expect(nextAfterRemoval([thread('a', 300)], 'a')).toBeNull();
  });

  it('returns null for an unknown id', () => {
    expect(nextAfterRemoval([thread('a', 300), thread('b', 200)], 'ghost')).toBeNull();
    expect(nextAfterRemoval([], 'ghost')).toBeNull();
  });
});

/**
 * `loadThread`'s unreadable-row count must be SURFACED, never silently dropped.
 * A thread that quietly shows four of its six messages is worse than one that
 * says two could not be read: the user is reading a legal conversation and has
 * no other way to know a turn is missing.
 */
describe('unreadableCount', () => {
  it('reports nothing when every stored row came through', () => {
    expect(unreadableCount(loadedWith(3, 0), 3)).toBe(0);
    expect(unreadableCount(loadedWith(0, 0), 0)).toBe(0);
  });

  it('counts the rows main could not parse', () => {
    expect(unreadableCount(loadedWith(3, 2), 3)).toBe(2);
  });

  // The loss only the renderer can see: main parsed the JSON, but the shape is
  // not a well-formed UI message. A count that read `skipped` alone would let
  // this half go dark, which is the bug this test exists to prevent.
  it('counts the rows main returned that the renderer then rejected', () => {
    expect(unreadableCount(loadedWith(6, 0), 4)).toBe(2);
  });

  it('sums BOTH losses, because the user cares about the total', () => {
    expect(unreadableCount(loadedWith(6, 3), 4)).toBe(5);
  });

  // A `usable` larger than what main returned can only be a bookkeeping slip;
  // it must not subtract from a real `skipped` and hide it.
  it('never lets an over-large usable count cancel out skipped rows', () => {
    expect(unreadableCount(loadedWith(2, 4), 9)).toBe(4);
  });
});

describe('unreadableNotice', () => {
  it('says nothing when there is nothing to disclose', () => {
    expect(unreadableNotice(0)).toBeNull();
    expect(unreadableNotice(-1)).toBeNull();
  });

  it('uses singular grammar for one lost row', () => {
    const notice = unreadableNotice(1);
    expect(notice).toContain('1 earlier message');
    expect(notice).not.toContain('messages');
  });

  it('uses plural grammar from two upwards', () => {
    expect(unreadableNotice(2)).toContain('2 earlier messages');
    expect(unreadableNotice(17)).toContain('17 earlier messages');
  });

  // The sentence has to state the consequence, not just the number. "2 messages
  // skipped" leaves the user wondering whether they are further down the pane.
  it('says the messages are not shown', () => {
    expect(unreadableNotice(1)).toContain('not shown');
    expect(unreadableNotice(2)).toContain('not shown');
    expect(unreadableNotice(2)).toContain('could not be read');
  });
});
