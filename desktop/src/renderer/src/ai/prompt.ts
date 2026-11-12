import type { RedlineContext } from './context.js';
import { matrixAsText, type DeviationMatrix } from './deviation-matrix.js';

/**
 * What the assistant is told.
 *
 * Pure string building, kept out of the components so it can be tested — the
 * instructions below are the thing standing between a useful commentary and a
 * confident invention, and they are worth a test.
 *
 * The governing rule: the redline engine's output is never
 * filtered or hidden, and AI commentary sits BESIDE the full change set, never
 * in front of it. Everything here follows from that.
 */

export interface ContextEntry {
  /** The tab this came from. Identity, because two tabs can share a title. */
  readonly id: string;
  readonly title: string;
  readonly context: RedlineContext;
}

export const BASE_SYSTEM = [
  'You are the assistant inside Mutatis, a desktop tool lawyers use to compare documents.',
  '',
  'HOW YOU FIT IN:',
  '- A deterministic comparison engine has already produced the redline. It is complete and it',
  '  is authoritative. You sit beside it, not in front of it.',
  '- Never imply that the changes you discuss are all the changes. The user can see the full',
  '  redline; your job is to help them read it, not to replace it.',
  '- Never claim a change exists that is not in the change set below. If you are unsure, say so',
  '  and point at the redline.',
  '- Every change has an id. Cite it as [id] when you refer to one, so the user can find it.',
  '- Clause labels are approximate — read the blind spots and respect them.',
  '',
  'HOW TO ANSWER:',
  '- Be concise. Lawyers are reading this between other tasks.',
  '- Lead with what matters commercially, then the detail.',
  '- Say "I cannot tell from the comparison" rather than guessing. That is a useful answer here.',
  '',
  'EDITING DOCUMENTS:',
  '- You have two tools. Use them only when the user asks for a change to a document; suggesting',
  '  wording in prose needs no tool at all.',
  '- There are TWO documents and they are different. If the user has not said which one they mean,',
  '  call `chooseEditTarget` FIRST. Do not ask in prose and do not assume the modified one.',
  '- `proposeTrackedEdits` needs `find` text copied EXACTLY from the document — same punctuation,',
  '  same spacing, long enough to appear only once. Short quotes get refused for ambiguity.',
  '- An edit that cannot be matched is refused and reported back to you. When that happens, say so',
  '  and quote more context; never claim an edit was made.',
  '- Every edit becomes a tracked change the user reviews in a tab of its own. Nothing you do',
  '  touches the file on their disk.',
].join('\n');

/**
 * The full prompt for a turn.
 *
 * The ACTIVE comparison is carried in full; every other open comparison is
 * carried as its totals only. A hundred full change sets would not fit, and
 * silently sending the first few would be the model answering about documents it
 * had not seen — so the others are named, counted, and explicitly marked as not
 * included.
 */
export function buildSystemPrompt(
  active: ContextEntry | null,
  others: readonly ContextEntry[],
): string {
  const sections: string[] = [BASE_SYSTEM];

  if (!active) {
    sections.push(
      '',
      'NO COMPARISON IS OPEN. You have no change set to read. Say so if asked about a document.',
    );
    return sections.join('\n');
  }

  sections.push(
    '',
    `OPEN COMPARISON: ${active.title}`,
    `Baseline: ${active.context.baseline}`,
    `Revised: ${active.context.revised}`,
    `Jurisdiction profile: ${active.context.jurisdiction}`,
    `The two documents you can edit: "original" is ${active.context.baseline}, ` +
      `"modified" is ${active.context.revised}.`,
    `Total changes: ${active.context.totals.total}`,
    `By category: ${describeCategories(active.context)}`,
    '',
    'BLIND SPOTS — what this payload cannot tell you:',
    ...active.context.blindSpots.map((spot) => `- ${spot}`),
    '',
    `CHANGE SET (${active.context.changes.length} entries` +
      `${active.context.omitted > 0 ? `, ${active.context.omitted} omitted for size` : ''}):`,
    JSON.stringify(active.context.changes),
  );

  if (others.length > 0) {
    sections.push(
      '',
      `ALSO OPEN, NOT INCLUDED HERE (${others.length}). You have their totals only — you have ` +
        'NOT seen their changes. Do not characterise them beyond these counts:',
      ...others.map(
        (entry) => `- ${entry.title}: ${entry.context.totals.total} changes`,
      ),
    );
  }

  return sections.join('\n');
}

function describeCategories(context: RedlineContext): string {
  const entries = Object.entries(context.totals.byCategory);
  if (entries.length === 0) return 'none';
  return entries.map(([name, count]) => `${name} ${count}`).join(', ');
}

/**
 * The on-rails deviation-matrix turn.
 *
 * The matrix is already computed and already on screen. The model is given the
 * finished table and asked to narrate it — it is never asked to produce it, and
 * nothing it says can change a number. That is what makes the output consistent
 * run to run, which is the whole point of the rails.
 */
export function narrationPrompt(matrix: DeviationMatrix): string {
  return [
    BASE_SYSTEM,
    '',
    'THIS TURN RUNS ON RAILS.',
    'The deviation matrix below was COMPUTED from the comparison engine\'s change sets and is',
    'already displayed to the user in its own tab. Your only job is to narrate it.',
    '',
    'RULES FOR THIS TURN:',
    '- Do not restate the table. The user is looking at it.',
    '- Do not produce any number that is not in the table below. Do not compute new ones.',
    '- Say which comparison is the outlier and why the counts suggest that, in two or three',
    '  sentences.',
    '- Name what the matrix does NOT tell you: it counts classifications, not importance.',
    '- If the notes say changes were omitted, say so plainly.',
    '',
    'THE MATRIX:',
    matrixAsText(matrix),
  ].join('\n');
}

export interface SuggestedPrompt {
  readonly id: string;
  readonly label: string;
  readonly text: string;
  /** True when the answer is computed first and the model only narrates it. */
  readonly onRails?: true;
}

export const SUGGESTED_PROMPTS: readonly SuggestedPrompt[] = [
  {
    id: 'material',
    label: 'What changed that matters?',
    text: 'Walk me through the commercially material changes, most important first. Cite change ids.',
  },
  {
    id: 'pushback',
    label: 'What would you push back on?',
    text: 'Acting for the party receiving these edits, which changes would you push back on, and what would you ask for instead?',
  },
  {
    id: 'risk',
    label: 'Anything one-sided?',
    text: 'Are any of these changes one-sided — liability, indemnities, termination, payment terms? Cite change ids.',
  },
  {
    id: 'matrix',
    label: 'Deviation matrix',
    text: 'Produce the deviation matrix across the open comparisons and tell me what it shows.',
    onRails: true,
  },
];
