import { useRef, useState } from 'react';
import { ArrowUp, Square } from 'lucide-react';

interface ComposerProps {
  busy: boolean;
  disabled: boolean;
  onSend: (text: string) => void;
  onStop: () => void;
}

export function Composer({ busy, disabled, onSend, onStop }: ComposerProps) {
  const [value, setValue] = useState('');
  const areaRef = useRef<HTMLTextAreaElement>(null);

  function grow() {
    const area = areaRef.current;
    if (!area) return;
    area.style.height = 'auto';
    area.style.height = `${Math.min(area.scrollHeight, 160)}px`;
  }

  function submit() {
    if (busy || disabled || value.trim() === '') return;
    onSend(value);
    setValue('');
    requestAnimationFrame(grow);
  }

  return (
    <div className="shrink-0 border-t border-border p-2">
      <div className="flex items-end gap-1.5 rounded-md border border-border bg-surface px-2 py-1.5 focus-within:border-focus">
        <textarea
          ref={areaRef}
          value={value}
          rows={1}
          disabled={disabled}
          aria-label="Message the assistant"
          placeholder={disabled ? 'Add an API key to start' : 'Ask about this comparison…'}
          onChange={(event) => {
            setValue(event.target.value);
            grow();
          }}
          onKeyDown={(event) => {
            // Enter sends, Shift+Enter is a newline. IME composition must never
            // be interrupted — during composition Enter is committing a
            // candidate, not sending a message.
            if (event.key === 'Enter' && !event.shiftKey && !event.nativeEvent.isComposing) {
              event.preventDefault();
              submit();
            }
          }}
          className="base-ui-disable-scrollbar max-h-40 flex-1 resize-none bg-transparent text-[12px] leading-relaxed text-text outline-none placeholder:text-text-faint disabled:opacity-50"
        />
        {busy ? (
          <button
            type="button"
            onClick={onStop}
            aria-label="Stop generating"
            title="Stop generating"
            className="flex h-6 w-6 shrink-0 items-center justify-center rounded bg-border text-text transition-colors hover:bg-border-strong"
          >
            <Square size={11} aria-hidden />
          </button>
        ) : (
          <button
            type="button"
            onClick={submit}
            disabled={disabled || value.trim() === ''}
            aria-label="Send message"
            title="Send message"
            className="flex h-6 w-6 shrink-0 items-center justify-center rounded bg-accent text-accent-text transition-opacity hover:opacity-90 disabled:opacity-30"
          >
            <ArrowUp size={13} aria-hidden />
          </button>
        )}
      </div>
    </div>
  );
}
