// src/features/repair/apply/metadata.ts
import type { RepairFinding } from '../types.js';

export async function applyMetadata(
  ctx: Word.RequestContext,
  findings: RepairFinding[]
): Promise<RepairFinding[]> {
  const results: RepairFinding[] = [];

  for (const finding of findings) {
    switch (finding.action) {
      case 'clearBuiltinProps': {
        try {
          const p = ctx.document.properties;
          p.author = '';
          p.subject = '';
          p.title = '';
          p.keywords = '';
          p.company = '';
          p.manager = '';
          p.comments = '';
          p.category = '';
          // lastAuthor and revisionNumber are Word-managed; we clear only writable props
          await ctx.sync();
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      case 'deleteCustomProps': {
        try {
          const customProps = ctx.document.properties.customProperties;
          customProps.load('items');
          await ctx.sync();
          for (const item of customProps.items) {
            item.delete();
          }
          await ctx.sync();
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      case 'deleteCustomXml': {
        try {
          const xmlParts = ctx.document.customXmlParts;
          xmlParts.load('items');
          await ctx.sync();
          for (const part of xmlParts.items) {
            part.delete();
          }
          await ctx.sync();
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      case 'deleteComments': {
        try {
          const commentsCollection = ctx.document.body.getComments();
          commentsCollection.load('items');
          await ctx.sync();
          for (const comment of commentsCollection.items) {
            comment.delete();
          }
          await ctx.sync();
          results.push({ ...finding, result: 'done' });
        } catch (err) {
          if (isGeneralException(err)) {
            results.push({ ...finding, result: 'error' });
          } else {
            throw err;
          }
        }
        break;
      }

      default: {
        results.push({ ...finding, result: 'skipped' });
        break;
      }
    }
  }

  return results;
}

function isGeneralException(err: unknown): boolean {
  if (err instanceof Error && err.name === 'GeneralException') return true;
  if (
    typeof err === 'object' &&
    err !== null &&
    (err as Record<string, unknown>)['code'] === 'GeneralException'
  ) {
    return true;
  }
  return false;
}
