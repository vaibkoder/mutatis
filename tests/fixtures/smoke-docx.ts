/**
 * The optional `smoke-test/` redline pair (`original.docx` / `revised.docx`) —
 * local binary fixtures that are **not in the repository**.
 *
 * `.gitignore` carries `smoke-test/`. Tests that read the pair must be wrapped
 * in `describe.skipIf(!smokeFixturesAvailable())`: absence is a visible SKIP,
 * never a throw.
 */
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

/** Repository root, from this file's location. */
const ROOT = resolve(__dirname, '../..');

/** The two packages the smoke-test compare is driven over. */
const FILES = ['original.docx', 'revised.docx'] as const;

/** Absolute path to one smoke fixture. */
export function smokeDocxPath(name: string): string {
  return resolve(ROOT, 'smoke-test', name);
}

/**
 * True when BOTH packages are on disk. Wrap every dependent block in
 * `describe.skipIf(!smokeFixturesAvailable())` — a skip is a visible absence,
 * a throw is a broken suite, and a silently-passing empty block is neither.
 */
export function smokeFixturesAvailable(): boolean {
  return FILES.every((f) => existsSync(smokeDocxPath(f)));
}

/** Read one smoke fixture. Only call inside a `skipIf`-guarded block. */
export function smokeDocxBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(smokeDocxPath(name)));
}
