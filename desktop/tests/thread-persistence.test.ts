import { afterAll, describe, expect, it, vi } from 'vitest';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import type { StoredMessage, StoredThread } from '../src/common/ipc.js';

/**
 * `src/main/threads.ts` reaches for Electron's userData directory the first
 * time it opens the database. Point that at a throwaway directory instead.
 *
 * `vi.mock` is hoisted above every import, so TEMP_DIR is still uninitialised
 * when this factory runs. It works anyway because `getPath` is a CLOSURE:
 * threads.ts calls it only inside its lazy `open()`, which no test reaches
 * until the module body below has assigned TEMP_DIR.
 */
vi.mock('electron', () => ({ app: { getPath: () => TEMP_DIR } }));

const TEMP_DIR = mkdtempSync(join(tmpdir(), 'mutatis-threads-'));
const DB_PATH = join(TEMP_DIR, 'threads.db');

import {
  closeThreads,
  deleteThread,
  ensureThread,
  listThreads,
  loadThread,
  renameThread,
  upsertMessage,
} from '../src/main/threads.js';
import { ThreadList, type ThreadListProps } from '../src/renderer/src/chat/ThreadList.js';

afterAll(() => {
  // Close first: threads.ts memoises the handle, and WAL leaves -wal and -shm
  // files that the connection still holds.
  closeThreads();
  rmSync(TEMP_DIR, { recursive: true, force: true });
});

/**
 * The restart. threads.ts caches its handle in a module-level `db`, so closing
 * is the only way to prove the next read comes off disk rather than out of a
 * connection that never let go of the write.
 */
function restart(): void {
  closeThreads();
}

/**
 * Advance the wall clock by at least a millisecond.
 *
 * threads.ts stamps `updated_at` with its own `Date.now()`, which no test can
 * inject. Two writes inside the same millisecond tie, and `ORDER BY updated_at
 * DESC` is then free to return either first — so without this the ordering
 * assertions would pass by luck on a fast machine.
 */
function tick(): void {
  const start = Date.now();
  while (Date.now() - start < 2) {
    // Busy-wait on purpose: a timer would make every ordering test async.
  }
}

function storedMessage(
  threadId: string,
  id: string,
  text: string,
  createdAt: number,
): StoredMessage {
  return { id, threadId, role: 'user', parts: [{ type: 'text', text }], createdAt };
}

function textsOf(threadId: string): string[] {
  return loadThread(threadId).messages.map((m) => {
    const parts = m.parts;
    const first = Array.isArray(parts) ? (parts[0] as { text?: unknown } | undefined) : undefined;
    return String(first?.text);
  });
}

/**
 * The relative order of just these ids in `listThreads`.
 *
 * Every test in this file shares one database file, so an absolute assertion on
 * the whole list would couple unrelated tests to each other.
 */
function orderOf(...ids: readonly string[]): string[] {
  const wanted = new Set(ids);
  return listThreads()
    .filter((t) => wanted.has(t.id))
    .map((t) => t.id);
}

/** A second connection to the same file, for rows threads.ts would never write. */
function rawDb(): DatabaseSync {
  const raw = new DatabaseSync(DB_PATH);
  raw.exec('PRAGMA foreign_keys = ON');
  return raw;
}

/**
 * GROUP 1 — the rail's markup.
 *
 * The test environment is `node`: no DOM, no jsdom, nothing to click. So these
 * assert on the exact HTML the browser would receive, which is the only level
 * at which "a hostile title cannot become an element" is a claim rather than a
 * hope. Static markup is FIRST RENDER only, so nothing here touches the rename
 * input or the delete confirmation — those are local state behind a click.
 */

/** A title that is user-typed AND arrives from a file anything on disk can edit. */
const HOSTILE_TITLE = '<img src=x onerror="alert(1)">';

function listed(id: string, title: string, updatedAt: number): StoredThread {
  return { id, title, createdAt: 1_000, updatedAt };
}

const SAMPLE_THREADS: readonly StoredThread[] = [
  listed('t1', 'Share purchase agreement', 300),
  listed('t2', 'Lease, clause 4.1', 200),
  listed('t3', 'NDA markup', 100),
];

function railProps(
  threads: readonly StoredThread[],
  openId: string,
  unsaved = false,
): ThreadListProps {
  return {
    threads,
    openId,
    unsaved,
    onOpen: () => {},
    onNew: () => {},
    onRename: () => {},
    onDelete: () => {},
  };
}

function render(props: ThreadListProps): string {
  return renderToStaticMarkup(createElement(ThreadList, props));
}

/** The markup of each row, so an attribute can be tied to the title beside it. */
function rowsOf(html: string): string[] {
  return html.split('<li').slice(1);
}

function occurrences(html: string, needle: string): number {
  return html.split(needle).length - 1;
}

/** Everything a user would actually read, with the tags stripped out. */
function visibleText(html: string): string {
  return html
    .replace(/<[^>]*>/gu, ' ')
    .replace(/\s+/gu, ' ')
    .trim();
}

describe('ThreadList — what it renders', () => {
  it('shows every thread it was given', () => {
    const html = render(railProps(SAMPLE_THREADS, 't1'));
    for (const thread of SAMPLE_THREADS) {
      expect(html).toContain(thread.title);
    }
    expect(rowsOf(html)).toHaveLength(SAMPLE_THREADS.length);
  });

  // A title is untrusted twice over: the user types it, and it also comes back
  // out of threads.db — an unencrypted SQLite file in userData that anything on
  // the machine can edit. React escapes text children, and this is the test
  // that says so rather than assuming it.
  it('escapes a hostile title into text instead of injecting an element', () => {
    const html = render(railProps([listed('t1', HOSTILE_TITLE, 300)], 't1'));
    // No element, so no attribute for the browser to fire.
    expect(html).not.toContain('<img');
    expect(html).not.toMatch(/onerror=(?!&quot;)/);
    // It arrived escaped, and stays visible — a user who sees this knows the
    // file has been tampered with.
    expect(html).toContain('&lt;img');
  });

  it('marks the open row with aria-current, and no other row', () => {
    const html = render(railProps(SAMPLE_THREADS, 't2'));
    expect(occurrences(html, 'aria-current="true"')).toBe(1);

    const current = rowsOf(html).filter((row) => row.includes('aria-current="true"'));
    expect(current).toHaveLength(1);
    expect(current[0]).toContain('Lease, clause 4.1');
    expect(current[0]).not.toContain('Share purchase agreement');
  });

  // Follows from "it marks the open row": if openId names no row, no row is the
  // open one. A rail that defaulted to marking the first row would tell the user
  // they are reading a conversation they are not.
  it('marks nothing when openId matches no stored thread and nothing is unsaved', () => {
    const html = render(railProps(SAMPLE_THREADS, 'not-a-thread'));
    expect(occurrences(html, 'aria-current="true"')).toBe(0);
  });

  // A <ul> with no <li> is a list the screen reader announces as empty and the
  // user sees as a blank strip. The empty state has to be prose instead.
  it('renders an empty state rather than an empty list', () => {
    const html = render(railProps([], ''));
    expect(html).not.toContain('<ul');
    expect(html).not.toContain('<li');
    expect(html).toContain('<p');
    // Longer than the one control that is always on screen, which is how this
    // fails if somebody drops the explanatory sentence and leaves the button.
    expect(visibleText(html).length).toBeGreaterThan('New conversation'.length);
  });

  // An empty rail must not be a dead end: the one control that can produce a
  // thread has to survive the branch that removes the list.
  it('still offers the new-conversation control when the list is empty', () => {
    expect(occurrences(render(railProps([], '')), 'aria-label="New conversation"')).toBe(1);
  });
});

describe('ThreadList — the unsaved conversation', () => {
  /**
   * A brand-new conversation exists in the renderer before SQLite has heard of
   * it: main writes the thread row on the first message. Without an entry for
   * it the rail shows nothing selected while the user is typing into it, which
   * reads as the message going nowhere.
   */
  it('adds one entry for the open-but-unsaved conversation', () => {
    const saved = render(railProps(SAMPLE_THREADS, 't1'));
    const draft = render(railProps(SAMPLE_THREADS, 'draft-1', true));

    expect(rowsOf(saved)).toHaveLength(SAMPLE_THREADS.length);
    expect(rowsOf(draft)).toHaveLength(SAMPLE_THREADS.length + 1);
  });

  it('treats that entry as the current one, so exactly one row is marked', () => {
    const html = render(railProps(SAMPLE_THREADS, 'draft-1', true));
    expect(occurrences(html, 'aria-current="true"')).toBe(1);
  });

  // The entry is not deletable, and this is the one way first-render markup can
  // say so: a delete control per STORED thread and not one more. There is
  // nothing on disk to delete yet, and a delete that silently did nothing would
  // teach the user that the button is unreliable.
  it('gives the unsaved entry no delete control', () => {
    const deletes = (html: string) => occurrences(html, 'aria-label="Delete ');
    expect(deletes(render(railProps(SAMPLE_THREADS, 't1')))).toBe(SAMPLE_THREADS.length);
    expect(deletes(render(railProps(SAMPLE_THREADS, 'draft-1', true)))).toBe(
      SAMPLE_THREADS.length,
    );
  });

  /**
   * The two rules I was handed conflict when BOTH hold — nothing saved and the
   * open conversation unsaved. "Empty state instead of a <ul>" wins, and that
   * is the right way round: on a fresh install the prose already says the
   * conversation on screen is saved as soon as the first message goes out, so
   * an "unsaved" row beside that sentence would be the same fact twice, in a
   * list of one. Pinned deliberately so the precedence is a decision rather
   * than an accident of which branch was written first.
   */
  it('lets the empty state stand in for the unsaved entry when nothing is saved', () => {
    const html = render(railProps([], 'draft-1', true));
    expect(html).not.toContain('<ul');
    expect(html).not.toContain('<li');
    expect(html).toContain('<p');
    expect(occurrences(html, 'aria-current="true"')).toBe(0);
  });
});

describe('ThreadList — busy', () => {
  /**
   * Switching thread mid-stream is how a half-written reply ends up attributed
   * to the wrong conversation. While a turn is in flight the controls that
   * would move the user off this thread must be inert, and the real `disabled`
   * attribute is the only form of inert that also reaches the keyboard and the
   * screen reader.
   *
   * Matched as `disabled=""` — the attribute React actually emits — rather than
   * as the bare word, because a Tailwind `disabled:` variant class sits in
   * every row's `class` attribute whether the rail is busy or not.
   */
  const DISABLED_ATTR = /\sdisabled=""/u;

  it('disables the switching controls while a turn is in flight', () => {
    const html = render({ ...railProps(SAMPLE_THREADS, 't1'), busy: true });
    expect(html).toMatch(DISABLED_ATTR);
  });

  it('leaves them enabled when it is not busy', () => {
    expect(render(railProps(SAMPLE_THREADS, 't1'))).not.toMatch(DISABLED_ATTR);
    expect(render({ ...railProps(SAMPLE_THREADS, 't1'), busy: false })).not.toMatch(DISABLED_ATTR);
  });
});

describe('threads.ts — the restart round trip', () => {
  // The whole point of the file. Everything the rail shows after a relaunch
  // comes back through this path, and nothing else in the suite closes the
  // handle, so nothing else proves the write actually reached disk.
  it('returns saved messages in created_at order after a close and reopen', () => {
    ensureThread('t-restart', 'Lease review');
    upsertMessage(storedMessage('t-restart', 'm3', 'third', 3_000));
    upsertMessage(storedMessage('t-restart', 'm1', 'first', 1_000));
    upsertMessage(storedMessage('t-restart', 'm2', 'second', 2_000));

    restart();

    const loaded = loadThread('t-restart');
    expect(loaded.skipped).toBe(0);
    expect(loaded.messages.map((m) => m.id)).toEqual(['m1', 'm2', 'm3']);
    expect(textsOf('t-restart')).toEqual(['first', 'second', 'third']);
  });

  it('round-trips the parts payload, not just the row', () => {
    ensureThread('t-parts', 'Parts');
    upsertMessage({
      id: 'p1',
      threadId: 't-parts',
      role: 'assistant',
      parts: [{ type: 'text', text: 'clause 4.1 was reworded' }, { type: 'text', text: 'and 7.2' }],
      createdAt: 5_000,
    });

    restart();

    const loaded = loadThread('t-parts');
    expect(loaded.messages[0]?.role).toBe('assistant');
    expect(loaded.messages[0]?.parts).toEqual([
      { type: 'text', text: 'clause 4.1 was reworded' },
      { type: 'text', text: 'and 7.2' },
    ]);
  });

  it('survives a rename across a restart', () => {
    ensureThread('t-rename', 'Original name');
    restart();
    renameThread('t-rename', 'Share purchase agreement');
    restart();

    expect(listThreads().find((t) => t.id === 't-rename')?.title).toBe(
      'Share purchase agreement',
    );
  });

  // Main's UPDATE matches no row for an unknown id, and it must stay that way:
  // an INSERT here would let the renderer's optimistic rename create a thread
  // that the user never opened.
  it('does not create a thread when renaming an unknown id', () => {
    renameThread('t-ghost', 'Phantom');
    restart();

    expect(listThreads().some((t) => t.id === 't-ghost')).toBe(false);
    expect(loadThread('t-ghost').messages).toEqual([]);
  });
});

describe('threads.ts — upsertMessage is an upsert', () => {
  // This is what keeps a half-streamed answer. An assistant message is written
  // when the stream starts and again as it completes; insert-only would either
  // throw on the primary key or keep the first, empty version forever.
  it('keeps the later parts for a message id saved twice, in one row', () => {
    ensureThread('t-upsert', 'Streaming');
    upsertMessage(storedMessage('t-upsert', 'stream-1', 'The clause', 9_000));
    upsertMessage(storedMessage('t-upsert', 'stream-1', 'The clause was reworded in full.', 9_000));

    restart();

    const loaded = loadThread('t-upsert');
    expect(loaded.messages).toHaveLength(1);
    expect(textsOf('t-upsert')).toEqual(['The clause was reworded in full.']);
  });

  it('creates the thread row when a message arrives before it', () => {
    upsertMessage(storedMessage('t-implicit', 'i1', 'no ensureThread first', 1_000));
    restart();

    expect(listThreads().some((t) => t.id === 't-implicit')).toBe(true);
    expect(loadThread('t-implicit').messages).toHaveLength(1);
  });
});

describe('threads.ts — deleteThread cascades', () => {
  /**
   * The schema declares ON DELETE CASCADE and `open()` sets
   * `PRAGMA foreign_keys = ON` — which SQLite defaults to OFF, per connection.
   * If either were missing the DELETE would leave orphaned message rows behind,
   * and because thread ids can be recycled those orphans would resurface inside
   * somebody else's conversation.
   */
  it('removes the thread and its messages, leaving no orphan rows', () => {
    ensureThread('t-doomed', 'Doomed');
    upsertMessage(storedMessage('t-doomed', 'd1', 'goodbye', 1_000));
    upsertMessage(storedMessage('t-doomed', 'd2', 'goodbye again', 2_000));

    deleteThread('t-doomed');
    restart();

    expect(loadThread('t-doomed').messages).toEqual([]);
    expect(listThreads().some((t) => t.id === 't-doomed')).toBe(false);

    // loadThread filters by thread_id, so it cannot tell a cascade from an
    // orphan. Count the raw rows to tell them apart.
    const raw = rawDb();
    try {
      expect(raw.prepare('SELECT id FROM messages WHERE thread_id = ?').all('t-doomed')).toEqual(
        [],
      );
    } finally {
      raw.close();
    }
  });

  it('leaves another thread’s messages intact', () => {
    ensureThread('t-keep', 'Keep');
    upsertMessage(storedMessage('t-keep', 'k1', 'still here', 1_000));
    ensureThread('t-drop', 'Drop');
    upsertMessage(storedMessage('t-drop', 'x1', 'gone', 1_000));

    deleteThread('t-drop');
    restart();

    expect(textsOf('t-keep')).toEqual(['still here']);
    expect(loadThread('t-drop').messages).toEqual([]);
  });
});

describe('threads.ts — listThreads ordering', () => {
  it('returns the most recently updated first', () => {
    ensureThread('t-order-old', 'Older');
    tick();
    ensureThread('t-order-new', 'Newer');
    restart();

    expect(orderOf('t-order-old', 't-order-new')).toEqual(['t-order-new', 't-order-old']);
  });

  // Saving a message bumps its thread's updated_at. The rail relies on this:
  // the conversation you just typed into has to be the one at the top.
  it('bumps a thread to the top when a message is saved to it', () => {
    ensureThread('t-bump-a', 'A');
    tick();
    ensureThread('t-bump-b', 'B');
    expect(orderOf('t-bump-a', 't-bump-b')).toEqual(['t-bump-b', 't-bump-a']);

    tick();
    upsertMessage(storedMessage('t-bump-a', 'bump1', 'typed here', 1_000));
    restart();

    expect(orderOf('t-bump-a', 't-bump-b')).toEqual(['t-bump-a', 't-bump-b']);
  });
});

/**
 * The invariant, end to end: a row that satisfies the column types but carries
 * the wrong SHAPE is accepted on write and only fails when history is read
 * back. It must cost ONE message, never the whole thread, and it must be
 * COUNTED so the renderer can tell the user something is missing.
 *
 * Rows like these cannot be produced through threads.ts — that is the point, so
 * they are written with a second connection to the same file, which is also how
 * they would really arrive: threads.db is an unencrypted file in userData that
 * anything on the machine can edit.
 */
describe('threads.ts — unreadable rows are disclosed, not dropped', () => {
  it('reports one skipped row per bad row and still returns the good messages', () => {
    ensureThread('t-corrupt', 'Corrupt');
    upsertMessage(storedMessage('t-corrupt', 'good-1', 'readable one', 1_000));
    upsertMessage(storedMessage('t-corrupt', 'good-2', 'readable two', 4_000));

    // Close before writing behind threads.ts's back, so the reopen below is a
    // genuine read of the file rather than of a cached handle.
    restart();

    const raw = rawDb();
    try {
      const insert = raw.prepare(
        'INSERT INTO messages (id, thread_id, role, parts, created_at) VALUES (?, ?, ?, ?, ?)',
      );
      // parts is NOT valid JSON: JSON.parse throws on the way out.
      insert.run('bad-json', 't-corrupt', 'assistant', '{not json', 2_000);
      // role is not one of user/assistant/system, so it can never become a
      // StoredMessage even though the JSON parses.
      insert.run('bad-role', 't-corrupt', 'tool', '[{"type":"text","text":"x"}]', 3_000);
    } finally {
      raw.close();
    }

    const loaded = loadThread('t-corrupt');
    expect(loaded.skipped).toBe(2);
    expect(loaded.messages.map((m) => m.id)).toEqual(['good-1', 'good-2']);
    expect(textsOf('t-corrupt')).toEqual(['readable one', 'readable two']);
  });

  // JSON that parses but is not an array is the quietest of the three: nothing
  // throws, so only the explicit Array.isArray check catches it.
  it('skips parts that parse to something other than an array', () => {
    ensureThread('t-shape', 'Shape');
    upsertMessage(storedMessage('t-shape', 'shape-good', 'fine', 1_000));
    restart();

    const raw = rawDb();
    try {
      raw
        .prepare(
          'INSERT INTO messages (id, thread_id, role, parts, created_at) VALUES (?, ?, ?, ?, ?)',
        )
        .run('shape-bad', 't-shape', 'user', '{"type":"text"}', 2_000);
    } finally {
      raw.close();
    }

    const loaded = loadThread('t-shape');
    expect(loaded.skipped).toBe(1);
    expect(loaded.messages.map((m) => m.id)).toEqual(['shape-good']);
  });

  it('survives a thread whose every row is unreadable, rather than throwing', () => {
    ensureThread('t-all-bad', 'All bad');
    restart();

    const raw = rawDb();
    try {
      raw
        .prepare(
          'INSERT INTO messages (id, thread_id, role, parts, created_at) VALUES (?, ?, ?, ?, ?)',
        )
        .run('only-bad', 't-all-bad', 'user', 'not json either', 1_000);
    } finally {
      raw.close();
    }

    const loaded = loadThread('t-all-bad');
    expect(loaded.skipped).toBe(1);
    expect(loaded.messages).toEqual([]);
  });
});
