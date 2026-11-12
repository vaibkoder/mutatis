import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useChat } from '@ai-sdk/react';
import { lastAssistantMessageIsCompleteWithToolCalls, type UIMessage } from 'ai';
import type { ModelChoice, StoredMessage } from '@common/ipc.js';
import { IpcChatTransport } from './transport.js';
import { titleFromText } from './thread-store.js';

/**
 * `useChat`, plus the two things the SDK deliberately leaves to the app:
 * persistence and a system prompt.
 *
 * Persistence writes on both edges of a turn — the user message when it is sent,
 * the assistant message when it finishes — rather than snapshotting the whole
 * thread. A snapshot would be simpler and would also be the thing that loses a
 * half-streamed answer when the app is closed mid-turn, which is precisely when
 * a user most wants the text back.
 */

/** The message's prose, with tool calls and other non-text parts left out. */
function textOf(message: UIMessage): string {
  return message.parts
    .filter((part): part is { type: 'text'; text: string } => part.type === 'text')
    .map((part) => part.text)
    .join(' ');
}

function toStored(message: UIMessage, threadId: string): StoredMessage {
  return {
    id: message.id,
    threadId,
    role: message.role === 'assistant' ? 'assistant' : message.role === 'system' ? 'system' : 'user',
    parts: message.parts,
    createdAt: Date.now(),
  };
}

export interface ChatThreadOptions {
  readonly threadId: string;
  readonly initialMessages: UIMessage[];
  readonly model: ModelChoice;
  /**
   * The system prompt for the NEXT turn, read at send time.
   *
   * A function rather than a string because the prompt carries the open
   * comparison's change set, which changes as the user switches tabs — and
   * because an on-rails turn substitutes a different prompt entirely for one
   * message. Reading it at send time is what makes both work without rebuilding
   * the transport mid-stream.
   */
  readonly system: () => string;
  /**
   * Perform a client-side tool call. The editing tools have no `execute` in
   * main, so the work happens here, in the process that actually holds the
   * documents.
   */
  readonly onToolCall?: (call: { toolName: string; toolCallId: string; input: unknown }) => void;
  /** Called once when a thread first gets a message, so the sidebar can name it. */
  readonly onTitled?: (threadId: string, title: string) => void;
}

export function useChatThread({
  threadId,
  initialMessages,
  model,
  system,
  onToolCall,
  onTitled,
}: ChatThreadOptions) {
  // The model is read at send time through this ref, so switching models does
  // not rebuild the transport and tear down an in-flight turn.
  const modelRef = useRef(model);
  modelRef.current = model;
  const systemRef = useRef(system);
  systemRef.current = system;

  const transport = useMemo(
    () =>
      new IpcChatTransport({
        model: () => modelRef.current,
        system: () => systemRef.current(),
      }),
    [],
  );

  const [named, setNamed] = useState(initialMessages.length > 0);

  const chat = useChat({
    id: threadId,
    messages: initialMessages,
    transport,
    // A tool result has to go back to the model for the turn to continue. The
    // MCQ card is the exception that makes this worth stating: it deliberately
    // does NOT answer in `onToolCall`, so nothing is sent until the user picks —
    // this condition then fires on their click.
    sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls,
    ...(onToolCall
      ? {
          onToolCall: ({ toolCall }: { toolCall: { toolName: string; toolCallId: string; input: unknown } }) => {
            onToolCall(toolCall);
          },
        }
      : {}),
    onFinish: ({ message }) => {
      void window.mutatis.threads.save(toStored(message, threadId));
    },
  });

  const { messages, sendMessage, status, stop, error, clearError, addToolResult } = chat;

  // Persist the user's message as soon as it lands in state. Doing it here
  // rather than in the submit handler means a regenerate or an edit is captured
  // by the same path.
  // Seeded once, from the history this mount was handed. That is still right
  // because `ChatPanel` keys `ChatThread` by thread id: switching conversations
  // remounts this hook, so the set is re-seeded with the new thread's rows
  // rather than carrying the previous thread's ids over and skipping saves.
  const savedIds = useRef(new Set(initialMessages.map((message) => message.id)));
  useEffect(() => {
    for (const message of messages) {
      if (message.role !== 'user' || savedIds.current.has(message.id)) continue;
      savedIds.current.add(message.id);

      if (!named) {
        setNamed(true);
        const title = titleFromText(textOf(message));
        void window.mutatis.threads.ensure(threadId, title);
        onTitled?.(threadId, title);
      }
      void window.mutatis.threads.save(toStored(message, threadId));
    }
  }, [messages, named, onTitled, threadId]);

  const send = useCallback(
    (text: string) => {
      const trimmed = text.trim();
      if (trimmed === '') return;
      void sendMessage({ text: trimmed });
    },
    [sendMessage],
  );

  return {
    messages,
    send,
    stop,
    status,
    error,
    clearError,
    addToolResult,
    busy: status === 'submitted' || status === 'streaming',
  };
}
