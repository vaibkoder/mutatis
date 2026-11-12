import { useEffect, useRef } from 'react';
import type { UIMessage } from 'ai';
import { Markdown } from './Markdown.js';
import {
  ChooseTargetCard,
  ProposedEditsCard,
  type DocumentNames,
  type ToolPart,
} from '../edits/EditToolCards.js';

interface MessageListProps {
  messages: UIMessage[];
  busy: boolean;
  /** The two documents the assistant could edit, for the MCQ card's labels. */
  names: DocumentNames | null;
  onChooseTarget: (toolCallId: string, target: 'original' | 'modified') => void;
}

/**
 * Assistant output is markdown; the user's own text is not.
 *
 * The asymmetry is deliberate. The user typed what they typed — reinterpreting
 * it would mangle a pasted clause that happens to contain `*` or `#`, and they
 * gain nothing from seeing their own words restyled. The model, on the other
 * hand, emits markdown whether or not anyone renders it, so the choice is
 * between reading it and reading its punctuation. See Markdown.tsx for what
 * that renderer refuses to do with untrusted output.
 */
export function MessageList({ messages, busy, names, onChooseTarget }: MessageListProps) {
  const endRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const pinned = useRef(true);

  // Follow the stream, but stop following the moment the user scrolls up to
  // re-read something. Yanking them back to the bottom mid-sentence is the most
  // reliably irritating behaviour a chat panel can have.
  useEffect(() => {
    if (pinned.current) endRef.current?.scrollIntoView({ block: 'end' });
  }, [messages]);

  function onScroll() {
    const element = scrollRef.current;
    if (!element) return;
    const distance = element.scrollHeight - element.scrollTop - element.clientHeight;
    pinned.current = distance < 40;
  }

  return (
    <div
      ref={scrollRef}
      onScroll={onScroll}
      className="base-ui-disable-scrollbar flex-1 overflow-y-auto px-3 py-3"
    >
      <div className="flex flex-col gap-3">
        {messages.map((message) => (
          <Bubble
            key={message.id}
            message={message}
            names={names}
            onChooseTarget={onChooseTarget}
          />
        ))}
        {busy && lastIsUser(messages) && (
          <p className="text-[11px] text-text-faint" aria-live="polite">
            Thinking…
          </p>
        )}
      </div>
      <div ref={endRef} />
    </div>
  );
}

function lastIsUser(messages: UIMessage[]): boolean {
  return messages[messages.length - 1]?.role === 'user';
}

function Bubble({
  message,
  names,
  onChooseTarget,
}: {
  message: UIMessage;
  names: DocumentNames | null;
  onChooseTarget: (toolCallId: string, target: 'original' | 'modified') => void;
}) {
  const text = message.parts
    .filter((part): part is { type: 'text'; text: string } => part.type === 'text')
    .map((part) => part.text)
    .join('');

  if (message.role === 'user') {
    return (
      <div className="self-end max-w-[92%] rounded-md bg-accent-soft px-2.5 py-1.5">
        <p className="whitespace-pre-wrap text-[12px] leading-relaxed text-text">{text}</p>
      </div>
    );
  }

  // Tool parts are rendered in part order, so an MCQ card appears where the
  // assistant asked the question rather than being collected at the end.
  const tools = message.parts
    .filter((part) => part.type.startsWith('tool-'))
    .map((part) => part as unknown as ToolPart);

  return (
    <div className="flex max-w-full flex-col gap-2">
      {text !== '' && <Markdown>{text}</Markdown>}
      {tools.map((part) =>
        part.type === 'tool-chooseEditTarget' ? (
          <ChooseTargetCard
            key={part.toolCallId}
            part={part}
            names={names}
            onChoose={onChooseTarget}
          />
        ) : part.type === 'tool-proposeTrackedEdits' ? (
          <ProposedEditsCard key={part.toolCallId} part={part} />
        ) : null,
      )}
    </div>
  );
}
