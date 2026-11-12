import type { RedlineContext } from '../ai/context.js';

/**
 * The change set as plain text for pasting into an AI chat or saving as a file.
 * Built from the same RedlineContext the assistant reads, so it carries the full
 * change set, the counts and the blind-spot list. Nothing leaves the machine:
 * the user decides where it goes next.
 */
export function contextToMarkdown(ctx: RedlineContext): string {
  const lines = [
    '# Mutatis change set',
    '',
    `**Baseline:** ${ctx.baseline}  `,
    `**Revised:** ${ctx.revised}  `,
    `**Jurisdiction:** ${ctx.jurisdiction}  `,
    `**Generated:** ${ctx.generated}  `,
    `**Total changes:** ${ctx.totals.total}` +
      Object.entries(ctx.totals.byCategory)
        .map(([k, v]) => ` · ${k} ${v}`)
        .join(''),
  ];
  if (ctx.omitted > 0) {
    lines.push('', `> ⚠ ${ctx.omitted} change(s) were found but are NOT included below.`);
  }
  if (ctx.blindSpots.length > 0) {
    lines.push('', '## Not covered by this export', ...ctx.blindSpots.map((b) => `> ⚠ ${b}`));
  }
  lines.push('', '## Changes');
  if (ctx.changes.length === 0) lines.push('', '_No changes._');
  for (const c of ctx.changes) {
    const target = c.move_target ? ` → ${c.move_target}` : '';
    const conf = c.low_confidence ? ' _(low confidence)_' : '';
    lines.push('', `### ${c.type} — ${c.clause || '_unlocated_'} [${c.classification}]${target}${conf}`);
    if (c.original !== undefined) lines.push(`- **Original:** ${c.original}`);
    if (c.revised !== undefined) lines.push(`- **Revised:** ${c.revised}`);
  }
  return lines.join('\n') + '\n';
}

export function contextToJson(ctx: RedlineContext): string {
  return JSON.stringify(ctx, null, 2) + '\n';
}

export function downloadText(text: string, mime: string, fileName: string): void {
  const url = URL.createObjectURL(new Blob([text], { type: mime }));
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
