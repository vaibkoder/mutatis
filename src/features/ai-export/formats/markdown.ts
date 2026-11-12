/**
 * Markdown formatter (feature 11 — rule F7: paste-into-chat rendering).
 *
 * Same envelope as the JSON formatter (F7), rendered as headings + tables an LLM
 * reads natively. Blind spots are a `> ⚠` blockquote (F6). Citations are human
 * labels only — never paraIds (F3/F9).
 */
import type {
  AiExportEnvelope, DefinedTermsExport, SectionOutlineExport, ClauseNode,
  ChangeSetExport, DealFactsExport, CastListExport, BriefingPreamble,
} from '../types.js';

export function toMarkdown(env: AiExportEnvelope): string {
  const blocks: string[] = [];

  blocks.push('# Mutatis AI Export');
  blocks.push(provenance(env));
  blocks.push(blindSpots(env.blind_spots));

  if (env.preamble) blocks.push(preamble(env.preamble));
  if (env.defined_terms) blocks.push(definedTerms(env.defined_terms));
  if (env.section_outline) blocks.push(sectionOutline(env.section_outline));
  if (env.change_set) blocks.push(changeSet(env.change_set));
  if (env.deal_facts) blocks.push(dealFacts(env.deal_facts));
  if (env.cast_list) blocks.push(castList(env.cast_list));

  return blocks.filter(b => b.length > 0).join('\n\n') + '\n';
}

// --- shared helpers --------------------------------------------------------

/** Escape pipe and newline so a cell never breaks the table grid. */
function cell(s: string | number | boolean | null | undefined): string {
  if (s === null || s === undefined) return '';
  return String(s).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}

function table(headers: string[], rows: Array<Array<string | number | boolean | null | undefined>>): string {
  const head = `| ${headers.join(' | ')} |`;
  const sep = `| ${headers.map(() => '---').join(' | ')} |`;
  const body = rows.map(r => `| ${r.map(cell).join(' | ')} |`).join('\n');
  return body ? `${head}\n${sep}\n${body}` : `${head}\n${sep}`;
}

// --- sections --------------------------------------------------------------

function provenance(env: AiExportEnvelope): string {
  return [
    `**Document:** ${env.document}`,
    `**Jurisdiction:** ${env.jurisdiction}`,
    `**Generated:** ${env.generated}`,
    `**Checks run:** ${env.checks_run.length ? env.checks_run.join(', ') : 'none'}`,
  ].join('  \n');
}

function blindSpots(spots: string[]): string {
  if (spots.length === 0) return '';
  return ['> ⚠ **Blind spots — not covered by this export:**', ...spots.map(s => `> - ${s}`)].join('\n');
}

function preamble(p: BriefingPreamble): string {
  const lines = ['## Briefing', p.summary];
  if (p.open_items.length > 0) {
    lines.push('', '### Open items');
    lines.push(table(['Type', 'Item', 'Clause'], p.open_items.map(o => [o.type, o.raw, o.clause])));
  }
  return lines.join('\n');
}

function definedTerms(dt: DefinedTermsExport): string {
  const lines = ['## Defined terms'];
  if (dt.glossary.length === 0) lines.push('_None detected._');
  for (const t of dt.glossary) {
    lines.push('', `### ${t.term}`);
    lines.push(`- **Definition:** ${t.definition || '_(text unavailable)_'}`);
    lines.push(`- **Defined at:** ${t.defined_at || '_unknown_'}`);
    lines.push(`- **Source:** ${t.source}`);
    if (t.defined_multiple_times) lines.push(`- **Defined multiple times:** yes`);
    if (t.variants && t.variants.length > 0) lines.push(`- **Variants:** ${t.variants.join(', ')}`);
  }
  return lines.join('\n');
}

function sectionOutline(so: SectionOutlineExport): string {
  const lines = ['## Section outline'];
  if (so.clauses.length === 0) lines.push('_No structure detected._');
  const walk = (nodes: ClauseNode[], depth: number): void => {
    for (const n of nodes) {
      const indent = '  '.repeat(depth);
      const heading = n.heading ? ` — ${n.heading}` : '';
      lines.push(`${indent}- **${n.number}** (${n.kind})${heading}`);
      if (n.children && n.children.length > 0) walk(n.children, depth + 1);
    }
  };
  walk(so.clauses, 0);
  return lines.join('\n');
}

function changeSet(cs: ChangeSetExport): string {
  const lines = ['## Change set', `**Baseline:** ${cs.baseline}  \n**Revised:** ${cs.revised}`];
  if (cs.changes.length === 0) lines.push('', '_No changes._');
  for (const c of cs.changes) {
    const conf = c.low_confidence ? ' _(low confidence)_' : '';
    const target = c.move_target ? ` → ${c.move_target}` : '';
    lines.push('', `### ${c.type} — ${c.clause || '_unlocated_'} [${c.classification}]${target}${conf}`);
    if (c.original !== undefined) lines.push(`- **Original:** ${c.original}`);
    if (c.revised !== undefined) lines.push(`- **Revised:** ${c.revised}`);
  }
  return lines.join('\n');
}

function dealFacts(df: DealFactsExport): string {
  const lines = ['## Deal facts'];

  lines.push('', '### Dates');
  lines.push(df.dates.length
    ? table(['Raw', 'ISO', 'Day', 'Role', 'Weekend', 'Clause'],
        df.dates.map(d => [d.raw, d.iso, d.day_of_week, d.role, d.flagged_weekend ? 'yes' : '', d.clause]))
    : '_None._');

  lines.push('', '### Amounts');
  lines.push(df.amounts.length
    ? table(['Raw', 'Value', 'Currency', 'Words', 'Mismatch', 'Clause'],
        df.amounts.map(a => [a.raw, a.value, a.currency, a.words_raw ?? '', a.mismatch ?? '', a.clause]))
    : '_None._');

  lines.push('', '### Placeholders');
  lines.push(df.placeholders.length
    ? table(['Raw', 'Category', 'Clause'], df.placeholders.map(p => [p.raw, p.category, p.clause]))
    : '_None._');

  return lines.join('\n');
}

function castList(cl: CastListExport): string {
  const lines = ['## Cast list'];
  lines.push(cl.parties.length
    ? table(['Role', 'Entities', 'Introduced at', 'Body uses', 'Finding'],
        cl.parties.map(p => [p.role, p.entities.join('; '), p.introduced_at ?? '', p.body_uses, p.finding ?? '']))
    : '_None._');
  return lines.join('\n');
}
