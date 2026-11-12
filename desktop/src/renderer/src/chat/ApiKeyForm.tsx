import { useState } from 'react';
import { KeyRound, ExternalLink } from 'lucide-react';
import type { ProviderInfo, SetKeyResult } from '@common/ipc.js';

interface ApiKeyFormProps {
  provider: ProviderInfo;
  encryptionAvailable: boolean;
  onSaved: () => void;
}

/**
 * The key gate.
 *
 * The key is written straight through to the main process and encrypted with
 * the OS keychain. It is never held in React state longer than the keystroke
 * that typed it, never logged, and never readable back — there is no IPC call
 * that returns one, only one that reports whether it exists.
 *
 * When the keychain is unavailable we refuse to store anything rather than
 * falling back to a plaintext file on disk. A key silently written in the clear
 * is worse than no key, because the user believes it is protected.
 */
export function ApiKeyForm({ provider, encryptionAvailable, onSaved }: ApiKeyFormProps) {
  const [value, setValue] = useState('');
  const [result, setResult] = useState<SetKeyResult | null>(null);
  const [saving, setSaving] = useState(false);

  async function save(event: React.FormEvent) {
    event.preventDefault();
    if (value.trim() === '' || saving) return;
    setSaving(true);
    const outcome = await window.mutatis.secrets.set(provider.id, value.trim());
    setSaving(false);
    setValue('');
    setResult(outcome);
    if (outcome.ok) onSaved();
  }

  return (
    <div className="flex flex-1 flex-col justify-center gap-3 px-5">
      <div className="flex items-center gap-2 text-text">
        <KeyRound size={14} className="text-accent" aria-hidden />
        <h2 className="text-[12px] font-semibold">Connect {provider.label}</h2>
      </div>

      <p className="text-[11px] leading-relaxed text-text-muted">
        Your key is encrypted with the system keychain and is never shown to this window. When
        you send a chat message, Mutatis sends the file names and the before and after wording
        of the changed paragraphs (not the whole document) to {provider.label}, using your key.
        Nothing is sent until you do.
      </p>

      {!encryptionAvailable && (
        <p className="rounded border border-danger/40 bg-danger/10 px-2 py-1.5 text-[11px] leading-relaxed text-danger">
          The system keychain is unavailable, so a key cannot be stored securely. Mutatis
          will not write one to disk in plain text. On Linux this usually means no
          gnome-keyring or kwallet is running.
        </p>
      )}

      <form onSubmit={save} className="flex flex-col gap-2">
        <input
          type="password"
          value={value}
          disabled={!encryptionAvailable || saving}
          onChange={(event) => setValue(event.target.value)}
          placeholder="sk-…"
          aria-label={`${provider.label} API key`}
          autoComplete="off"
          spellCheck={false}
          className="rounded border border-border bg-surface px-2 py-1.5 font-mono text-[11px] text-text outline-none placeholder:text-text-faint focus:border-focus disabled:opacity-50"
        />
        <button
          type="submit"
          disabled={!encryptionAvailable || saving || value.trim() === ''}
          className="rounded bg-accent px-2 py-1.5 text-[11px] font-medium text-accent-text transition-opacity hover:opacity-90 disabled:opacity-40"
        >
          {saving ? 'Saving…' : 'Save key'}
        </button>
      </form>

      {result?.ok === false && (
        <p className="text-[11px] leading-relaxed text-danger">{result.error}</p>
      )}

      <button
        type="button"
        onClick={() => void window.mutatis.openKeyPage(provider.keyUrl)}
        className="flex items-center gap-1 self-start text-[11px] text-text-muted underline-offset-2 hover:text-accent hover:underline"
      >
        Get a key from {provider.label}
        <ExternalLink size={11} aria-hidden />
      </button>
    </div>
  );
}
