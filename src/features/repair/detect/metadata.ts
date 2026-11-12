import type { RepairFinding } from '../types.js';

// Known DMS XML namespaces to identify and tag
const DMS_NAMESPACES = [
  'imanage',      // iManage Work / FileSite
  'netdocuments', // NetDocuments
  'opentext',     // OpenText eDOCS
] as const;

/** Human-readable label for a detected DMS namespace substring. */
function dmsLabel(ns: string): string {
  if (ns === 'imanage') return 'iManage';
  if (ns === 'netdocuments') return 'NetDocuments';
  if (ns === 'opentext') return 'OpenText';
  return ns;
}

/**
 * Detect confidentiality-relevant metadata in the document.
 *
 * Must be called INSIDE an existing `Word.run` context — the caller provides `ctx`.
 * Issues multiple `ctx.sync()` calls internally.
 */
export async function detectMetadata(ctx: Word.RequestContext): Promise<RepairFinding[]> {
  const findings: RepairFinding[] = [];

  // ── 1. Built-in document properties ─────────────────────────────────────
  const props = ctx.document.properties;
  props.load('author,lastAuthor,subject,title,keywords,company,manager,comments,category,revisionNumber');
  await ctx.sync();

  const builtinFields: Array<[string, unknown]> = [
    ['author', props.author],
    ['lastAuthor', props.lastAuthor],
    ['subject', props.subject],
    ['title', props.title],
    ['keywords', props.keywords],
    ['company', props.company],
    ['manager', props.manager],
    ['comments', props.comments],
    ['category', props.category],
    ['revisionNumber', props.revisionNumber],
  ];
  const nonEmpty = builtinFields.filter(([, v]) => v !== '' && v !== null && v !== undefined && v !== 0);
  if (nonEmpty.length > 0) {
    findings.push({
      id: 'metadata.builtinProperties',
      category: 'metadata',
      severity: 'confidentiality',
      label: 'Document properties contain author/identity information',
      detail: nonEmpty.map(([k]) => k).join(', '),
      count: nonEmpty.length,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'clearBuiltinProps',
    });
  }

  // ── 2. Custom document properties ───────────────────────────────────────
  const customProps = ctx.document.properties.customProperties;
  customProps.load('items');
  await ctx.sync();

  if (customProps.items.length > 0) {
    findings.push({
      id: 'metadata.customProperties',
      category: 'metadata',
      severity: 'confidentiality',
      label: 'Custom document properties present',
      detail: `${customProps.items.length} custom propert${customProps.items.length === 1 ? 'y' : 'ies'}`,
      count: customProps.items.length,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'deleteCustomProps',
    });
  }

  // ── 3. Custom XML parts (DMS artefacts) ─────────────────────────────────
  const xmlParts = ctx.document.customXmlParts;
  xmlParts.load('items');
  await ctx.sync();

  if (xmlParts.items.length > 0) {
    // Collect ClientResult handles for each part's XML, then sync once.
    const xmlResults: OfficeExtension.ClientResult<string>[] = xmlParts.items.map(part => part.getXml());
    await ctx.sync();

    // Inspect each XML string for known DMS namespace substrings.
    const detectedLabels = new Set<string>();
    for (const result of xmlResults) {
      const xml = result.value.toLowerCase();
      for (const ns of DMS_NAMESPACES) {
        if (xml.includes(ns)) {
          detectedLabels.add(dmsLabel(ns));
        }
      }
    }

    const partCount = xmlParts.items.length;
    const dmsDetail =
      detectedLabels.size > 0
        ? `${partCount} custom XML part${partCount === 1 ? '' : 's'} — detected: ${[...detectedLabels].join(', ')}`
        : `${partCount} custom XML part${partCount === 1 ? '' : 's'} — generic custom XML (possible DMS client/matter identifiers)`;

    findings.push({
      id: 'metadata.customXml',
      category: 'customXml',
      severity: 'confidentiality',
      label: 'Custom XML parts present (possible DMS metadata)',
      detail: dmsDetail,
      count: partCount,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'deleteCustomXml',
    });
  }

  // ── 4. Comments ──────────────────────────────────────────────────────────
  const body = ctx.document.body;
  const comments = body.getComments();
  comments.load('items');
  await ctx.sync();

  if (comments.items.length > 0) {
    findings.push({
      id: 'metadata.comments',
      category: 'comments',
      severity: 'confidentiality',
      label: 'Comments present',
      detail: `${comments.items.length} comment${comments.items.length === 1 ? '' : 's'}`,
      count: comments.items.length,
      destructive: false,
      requiresConfirm: false,
      selectedByDefault: true,
      action: 'deleteComments',
    });
  }

  return findings;
}
