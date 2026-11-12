/**
 * JSON formatter (feature 11 — rule F7: canonical machine format).
 *
 * The envelope is already a named-field projection (F9 — no rsids, ids, or style
 * plumbing), so serialisation is a straight stringify with 2-space indent.
 */
import type { AiExportEnvelope } from '../types.js';

export function toJson(envelope: AiExportEnvelope): string {
  return JSON.stringify(envelope, null, 2);
}
