import type { SelfRedlineOutcome } from '../compare/protocol.js';

/**
 * Save the edited document.
 *
 * A Blob download rather than anything through IPC: the bytes are already in the
 * renderer, the user picks the location through the browser's own save flow, and
 * main never needs a "write this file anywhere" handler — which is a capability
 * worth not having in the process that holds the API keys.
 *
 * The saved file is CLEAN, not redlined. The tracked changes live in the tab the
 * user is looking at; a file full of revision markup is not what someone wants
 * to send to the other side. If they want the redline as a document, that is the
 * .docx export of the comparison, which is a different thing on purpose.
 */
export function saveEditedDocument(outcome: SelfRedlineOutcome, targetName: string): void {
  const blob = new Blob([outcome.editedBytes], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `${targetName.replace(/\.docx$/i, '')} (edited).docx`;
  anchor.click();
  URL.revokeObjectURL(url);
}
