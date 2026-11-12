import { jsonSchema, tool, type ToolSet } from 'ai';

/**
 * The assistant's editing tools. MAIN PROCESS — these are the definitions the
 * provider is told about.
 *
 * None of them has an `execute`. That is deliberate and load-bearing: a tool
 * without `execute` is a CLIENT-side tool, so the call is streamed to the
 * renderer and the renderer decides what happens. Main never touches a document.
 * It holds the API key; the renderer holds the .docx. Neither needs what the
 * other has.
 *
 * JSON Schema rather than Zod, so the tool contract is data that can be read
 * without running anything, and so `zod` does not become a direct dependency of
 * a legal tool for the sake of two shapes.
 */

export const EDIT_TOOLS: ToolSet = {
  chooseEditTarget: tool({
    description:
      'Ask the user WHICH document to edit. Call this FIRST whenever the user has asked for an ' +
      'edit without saying whether they mean the original or the modified document. Do not ask ' +
      'in prose — this tool renders the question as a card with buttons, which is faster for ' +
      'the user and unambiguous for you.',
    inputSchema: jsonSchema<{ question: string }>({
      type: 'object',
      properties: {
        question: {
          type: 'string',
          description:
            'One short sentence saying what you are about to change, so the user knows what ' +
            'they are choosing a target for.',
        },
      },
      required: ['question'],
      additionalProperties: false,
    }),
  }),

  proposeTrackedEdits: tool({
    description:
      'Propose edits to one of the two documents. Every edit is applied and then redlined ' +
      'against the document it changed, so the user reviews it as a tracked change — you are ' +
      'never editing anything silently. ' +
      'The `find` text must be copied EXACTLY from the document, including punctuation and ' +
      'spacing, and must be long enough to occur only once. An edit whose text cannot be found, ' +
      'or that matches more than one place, is REFUSED and reported back to you — it is never ' +
      'guessed at. Never call this without knowing which document you are editing.',
    inputSchema: jsonSchema<{
      target: 'original' | 'modified';
      edits: { find: string; replace: string; reason?: string }[];
    }>({
      type: 'object',
      properties: {
        target: {
          type: 'string',
          enum: ['original', 'modified'],
          description: 'Which of the two compared documents to edit.',
        },
        edits: {
          type: 'array',
          minItems: 1,
          items: {
            type: 'object',
            properties: {
              find: {
                type: 'string',
                description:
                  'Exact text from the document to replace. Copy it verbatim; include enough ' +
                  'surrounding words that it appears only once.',
              },
              replace: {
                type: 'string',
                description: 'The replacement text. An empty string deletes the found text.',
              },
              reason: {
                type: 'string',
                description: 'One short sentence on why. Shown to the user, never inserted.',
              },
            },
            required: ['find', 'replace'],
            additionalProperties: false,
          },
        },
      },
      required: ['target', 'edits'],
      additionalProperties: false,
    }),
  }),
};
