/**
 * AI-readable export — public API (feature 11 — ADR-0021, spec §7).
 *
 * A pure, client-side serialisation layer. `buildExport`/`buildBriefingPack`
 * project in-memory checker data into an `AiExportEnvelope`; `formatEnvelope`
 * renders JSON or Markdown; `copyExport`/`downloadExport` write to the clipboard
 * or a local file. Zero network calls — the user decides what to do next.
 */
import type { AiExportEnvelope, CheckId, ExportFormat } from './types.js';
import type { SectionIndex } from '@shared/section-index/index.js';
import {
  buildSectionEnvelope, buildBriefingEnvelope,
  type ExportContext, type BriefingData,
  type DefinedTermsData, type ChangeSetData, type DealFactsData, type CastListData,
} from './assemble/briefing-pack.js';
import { toJson } from './formats/json.js';
import { toMarkdown } from './formats/markdown.js';

export type {
  AiExportEnvelope, CheckId, ExportFormat, ExportSectionKey,
} from './types.js';
export type {
  ExportContext, BriefingData,
  DefinedTermsData, ChangeSetData, DealFactsData, CastListData,
} from './assemble/briefing-pack.js';
export { partLabel } from './assemble/briefing-pack.js';

// UI
export { CopyButton } from './ui/CopyButton.js';
export { BriefingPackButton } from './ui/BriefingPackButton.js';
export { AwarenessNotice } from './ui/AwarenessNotice.js';
export { exportFormat, setExportFormat } from './ui/prefs.js';

/** Build a single-panel export envelope (#1–#5). */
export function buildExport(
  checkId: CheckId,
  data: DefinedTermsData | SectionIndex | ChangeSetData | DealFactsData | CastListData,
  ctx: ExportContext,
): AiExportEnvelope {
  return buildSectionEnvelope(checkId, data, ctx);
}

/** Build the composite briefing-pack envelope (#14). */
export function buildBriefingPack(data: BriefingData, ctx: ExportContext): AiExportEnvelope {
  return buildBriefingEnvelope(data, ctx);
}

/** Render an envelope to the chosen format (F7 — two serialisations, one schema). */
export function formatEnvelope(envelope: AiExportEnvelope, format: ExportFormat): string {
  return format === 'json' ? toJson(envelope) : toMarkdown(envelope);
}

/** Copy a rendered export to the clipboard (primary action). Local-only. */
export async function copyExport(envelope: AiExportEnvelope, format: ExportFormat): Promise<void> {
  await navigator.clipboard.writeText(formatEnvelope(envelope, format));
}

/** Download a rendered export as a local file (secondary action). Local-only. */
export function downloadExport(
  envelope: AiExportEnvelope,
  format: ExportFormat,
  label: string,
): void {
  const text = formatEnvelope(envelope, format);
  const ext = format === 'json' ? 'json' : 'md';
  const mime = format === 'json' ? 'application/json' : 'text/markdown';
  const isoDate = (envelope.generated.split('T')[0]) || 'export';
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `mutatis-export-${label}-${isoDate}.${ext}`;
  a.click();
  URL.revokeObjectURL(url);
}
