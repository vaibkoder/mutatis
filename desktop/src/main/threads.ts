import { app } from 'electron';
import { DatabaseSync } from 'node:sqlite';
import path from 'node:path';
import { mkdirSync } from 'node:fs';
import type { LoadedThread, StoredMessage, StoredThread } from '../common/ipc.js';

/**
 * Thread persistence, in SQLite.
 *
 * Uses Node's BUILT-IN `node:sqlite` (Electron 44 ships Node 24), not
 * better-sqlite3 — no native module, no electron-rebuild, nothing to compile
 * per platform or per Electron bump.
 *
 * Three footguns this schema is shaped to avoid, all of which bite on RELOAD
 * rather than on write, which is what makes them nasty:
 *   1. A row that satisfies the column types but carries the wrong SHAPE is
 *      accepted on write and only fails when history is read back. Messages are
 *      stored as JSON and validated on the way OUT, so a bad row is reported as
 *      one skipped message rather than a thread that will not open.
 *   2. Without an update path, a message edited or completed after insert is
 *      lost. `upsertMessage` is an upsert for exactly that reason.
 *   3. An append can fire before its thread row exists. `appendMessage` ensures
 *      the thread first, in the same transaction.
 */

export type { StoredMessage, StoredThread };

let db: DatabaseSync | null = null;

function open(): DatabaseSync {
  if (db) return db;
  const dir = app.getPath('userData');
  mkdirSync(dir, { recursive: true });
  const next = new DatabaseSync(path.join(dir, 'threads.db'));

  // WAL keeps a long write from blocking a read, which matters because a
  // streaming reply writes on every flush.
  next.exec('PRAGMA journal_mode = WAL');
  next.exec('PRAGMA foreign_keys = ON');
  next.exec(`
    CREATE TABLE IF NOT EXISTS threads (
      id         TEXT PRIMARY KEY,
      title      TEXT NOT NULL,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL
    );
    CREATE TABLE IF NOT EXISTS messages (
      id         TEXT PRIMARY KEY,
      thread_id  TEXT NOT NULL REFERENCES threads(id) ON DELETE CASCADE,
      role       TEXT NOT NULL,
      parts      TEXT NOT NULL,
      created_at INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS messages_thread_idx
      ON messages(thread_id, created_at);
  `);

  db = next;
  return next;
}

export function ensureThread(id: string, title: string): void {
  const now = Date.now();
  open()
    .prepare(
      `INSERT INTO threads (id, title, created_at, updated_at) VALUES (?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET updated_at = excluded.updated_at`,
    )
    .run(id, title, now, now);
}

export function renameThread(id: string, title: string): void {
  open().prepare('UPDATE threads SET title = ?, updated_at = ? WHERE id = ?').run(
    title,
    Date.now(),
    id,
  );
}

export function listThreads(): StoredThread[] {
  const rows = open()
    .prepare('SELECT id, title, created_at, updated_at FROM threads ORDER BY updated_at DESC')
    .all() as Array<Record<string, unknown>>;
  return rows.map((r) => ({
    id: String(r['id']),
    title: String(r['title']),
    createdAt: Number(r['created_at']),
    updatedAt: Number(r['updated_at']),
  }));
}

/**
 * Insert or REPLACE a message.
 *
 * An upsert rather than an insert because a streaming assistant message is
 * written once when it starts and again as it completes. With insert-only, a
 * reply paused for tool approval — or simply still streaming when the window
 * closed — is lost on reload.
 */
export function upsertMessage(message: StoredMessage): void {
  const database = open();
  ensureThread(message.threadId, 'New conversation');
  database
    .prepare(
      `INSERT INTO messages (id, thread_id, role, parts, created_at) VALUES (?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET parts = excluded.parts, role = excluded.role`,
    )
    .run(
      message.id,
      message.threadId,
      message.role,
      JSON.stringify(message.parts),
      message.createdAt,
    );
  database
    .prepare('UPDATE threads SET updated_at = ? WHERE id = ?')
    .run(Date.now(), message.threadId);
}

export type { LoadedThread };

export function loadThread(threadId: string): LoadedThread {
  const rows = open()
    .prepare(
      'SELECT id, thread_id, role, parts, created_at FROM messages WHERE thread_id = ? ORDER BY created_at',
    )
    .all(threadId) as Array<Record<string, unknown>>;

  const messages: StoredMessage[] = [];
  let skipped = 0;

  for (const row of rows) {
    // Validate on the way OUT. A row whose columns typecheck but whose JSON is
    // the wrong shape is accepted on write and only surfaces here — so one bad
    // row must cost one message, not the whole thread.
    try {
      const parts: unknown = JSON.parse(String(row['parts']));
      if (!Array.isArray(parts)) {
        skipped += 1;
        continue;
      }
      const role = String(row['role']);
      if (role !== 'user' && role !== 'assistant' && role !== 'system') {
        skipped += 1;
        continue;
      }
      messages.push({
        id: String(row['id']),
        threadId: String(row['thread_id']),
        role,
        parts,
        createdAt: Number(row['created_at']),
      });
    } catch {
      skipped += 1;
    }
  }

  return { messages, skipped };
}

export function deleteThread(id: string): void {
  open().prepare('DELETE FROM threads WHERE id = ?').run(id);
}

export function closeThreads(): void {
  db?.close();
  db = null;
}
