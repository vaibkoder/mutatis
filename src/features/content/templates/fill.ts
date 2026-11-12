/**
 * Template fill logic — C3.
 *
 * Scan → reconcile → render questionnaire → write values back via insertText.
 *
 * Office JS notes (ARCHITECTURE.md §Office JS Constraints Reference):
 *   - Batch: contentControls.load("tag,title,text,appearance") in ONE sync.
 *   - Write-back: processInChunks — no per-item ctx.sync() inside a loop.
 *   - Duplicate tags: getByTag returns a collection; scalar vars enforce
 *     uniqueness at authoring time; use .getFirstOrNullObject() defensively.
 */

import { wordRun, processInChunks } from '@shared/office/index.js';
import { formatValueForInsert } from './format.js';
import type { TemplateVariable, TemplateSchema, AnswersMap } from './types.js';
import { loadSchema } from './author.js';

// ---------------------------------------------------------------------------
// Reconciliation types
// ---------------------------------------------------------------------------

/** A tag present in the document but absent from the schema. */
export interface TagNotInSchema {
  tag: string;
  title: string;
  currentText: string;
}

/** A tag present in the schema but absent from the document. */
export interface TagNotInDocument {
  tag: string;
  label: string;
}

export interface ReconciliationResult {
  /** Variables in both schema and document (the happy path). */
  matched: TemplateVariable[];
  /** Tags found in the document but not in the schema (warn user). */
  inDocNotSchema: TagNotInSchema[];
  /** Tags in the schema but not found as content controls in the document. */
  inSchemaNotDoc: TagNotInDocument[];
  /** All content-control data loaded from the document. */
  documentControls: Array<{ tag: string; title: string; text: string; appearance: string }>;
}

// ---------------------------------------------------------------------------
// C3.1 — Scan and reconcile
// ---------------------------------------------------------------------------

/**
 * Load all content controls from the document, compare against the schema,
 * and return a reconciliation result.  Warns both directions.
 */
export async function scanAndReconcile(): Promise<ReconciliationResult> {
  const schema = await loadSchema();

  const documentControls = await wordRun(async ctx => {
    const controls = ctx.document.contentControls;
    controls.load('tag,title,text,appearance');
    await ctx.sync();
    return controls.items.map(cc => ({
      tag: cc.tag,
      title: cc.title,
      text: cc.text,
      appearance: cc.appearance,
    }));
  });

  const docTagSet = new Set(documentControls.map(c => c.tag).filter(Boolean));
  const schemaTagSet = new Set(schema.variables.map(v => v.tag));

  // Matched: in both schema and document.
  const matched = schema.variables.filter(v => docTagSet.has(v.tag));

  // In document but not schema.
  const inDocNotSchema: TagNotInSchema[] = documentControls
    .filter(c => c.tag && !schemaTagSet.has(c.tag))
    .map(c => ({ tag: c.tag, title: c.title, currentText: c.text }));

  // In schema but not document.
  const inSchemaNotDoc: TagNotInDocument[] = schema.variables
    .filter(v => !docTagSet.has(v.tag))
    .map(v => ({ tag: v.tag, label: v.label }));

  return { matched, inDocNotSchema, inSchemaNotDoc, documentControls };
}

// ---------------------------------------------------------------------------
// C3.2 — Write-back via processInChunks
// ---------------------------------------------------------------------------

/** A single fill operation: tag → formatted value. */
export interface FillItem {
  tag: string;
  type: string;
  rawValue: string | boolean | number;
}

export interface FillResult {
  /** Tags successfully written. */
  written: string[];
  /** Tags that were not found in the document as content controls. */
  notFound: string[];
}

/**
 * Write answer values into the document content controls via insertText.
 * Uses processInChunks to avoid per-item ctx.sync().
 * Chunk size 50 — bounded sync count on large templates.
 */
export async function fillQuestionnaire(
  items: FillItem[],
): Promise<FillResult> {
  if (items.length === 0) return { written: [], notFound: [] };

  const CHUNK_SIZE = 50;
  const written: string[] = [];
  const notFound: string[] = [];

  await processInChunks(
    items,
    CHUNK_SIZE,
    async (chunk, ctx) => {
      // Load all tags in one go per chunk.
      const controls = ctx.document.contentControls;
      controls.load('tag');
      await ctx.sync();

      const tagToControl = new Map<string, Word.ContentControl>();
      for (const cc of controls.items) {
        if (cc.tag) tagToControl.set(cc.tag, cc);
      }

      for (const item of chunk) {
        const cc = tagToControl.get(item.tag);
        if (!cc) {
          notFound.push(item.tag);
          continue;
        }
        const formatted = formatValueForInsert(item.type, item.rawValue);
        cc.insertText(formatted, 'Replace');
        written.push(item.tag);
      }

      // One sync per chunk — no per-item sync.
      await ctx.sync();
      return [];
    },
  );

  return { written, notFound };
}

// ---------------------------------------------------------------------------
// Answers map helpers
// ---------------------------------------------------------------------------

/**
 * Build an initial answers map from the schema's defaults and current
 * document control text values (from a prior reconciliation scan).
 */
export function buildInitialAnswers(
  schema: TemplateSchema,
  documentControls: Array<{ tag: string; text: string }>,
): AnswersMap {
  const docTextByTag = new Map(documentControls.map(c => [c.tag, c.text]));
  const answers: AnswersMap = {};

  for (const v of schema.variables) {
    const currentText = docTextByTag.get(v.tag) ?? '';
    // Use the document's current value if it looks like it has been filled,
    // otherwise fall back to the schema default.
    answers[v.tag] = currentText || (v.default ?? '');
  }

  return answers;
}
