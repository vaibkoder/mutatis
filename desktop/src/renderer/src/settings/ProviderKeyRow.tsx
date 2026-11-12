import { useState } from 'react';
import { Check, ExternalLink, Trash2 } from 'lucide-react';
import type { ProviderId, ProviderInfo } from '@common/ipc.js';

interface ProviderKeyRowProps {
  provider: ProviderInfo;
  configured: boolean;
  mask: string;
  encryptionAvailable: boolean;
  onSave: (provider: ProviderId, key: string) => Promise<{ ok: boolean; error?: string }>;
}

/**
 * One provider's key, in Settings.
 *
 * The stored key is shown as its mask and nothing more. There is no "reveal"
 * button, because there is no IPC call that could implement one — main has no
 * handler that returns a key, by design. Removing a key is an explicit button,
 * not an empty-field save, so it cannot happen by accident.
 */
export function ProviderKeyRow({
  provider,
  configured,
  mask,
  encryptionAvailable,
  onSave,
}: ProviderKeyRowProps) {
  const [value, setValue] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [saved, setSaved] = useState(false);

  async function submit(event: React.FormEvent) {
    event.preventDefault();
    if (value.trim() === '' || busy) return;
    setBusy(true);
    setError(null);
    const result = await onSave(provider.id, value.trim());
    setBusy(false);
    if (result.ok) {
      setValue('');
      setSaved(true);
      window.setTimeout(() => setSaved(false), 2000);
    } else {
      setError(result.error ?? 'The key could not be saved.');
    }
  }

  async function remove() {
    setBusy(true);
    setError(null);
    const result = await onSave(provider.id, '');
    setBusy(false);
    if (!result.ok) setError(result.error ?? 'The key could not be removed.');
  }

  return (
    <div className="flex flex-col gap-1.5 rounded border border-border p-2.5">
      <div className="flex items-center gap-2">
        <span className="text-[12px] font-medium text-text">{provider.label}</span>
        {configured ? (
          <span className="rounded bg-accent-soft px-1.5 py-0.5 font-mono text-[10px] text-text-muted">
            {mask}
          </span>
        ) : (
          <span className="text-[11px] text-text-faint">No key</span>
        )}
        <button
          type="button"
          onClick={() => void window.mutatis.openKeyPage(provider.keyUrl)}
          className="ml-auto flex items-center gap-1 text-[11px] text-text-muted underline-offset-2 hover:text-accent hover:underline"
        >
          Get a key
          <ExternalLink size={10} aria-hidden />
        </button>
        {configured && (
          <button
            type="button"
            onClick={() => void remove()}
            disabled={busy}
            aria-label={`Remove the ${provider.label} key`}
            title={`Remove the ${provider.label} key`}
            className="flex h-6 w-6 items-center justify-center rounded text-text-muted transition-colors hover:bg-border hover:text-danger disabled:opacity-40"
          >
            <Trash2 size={12} aria-hidden />
          </button>
        )}
      </div>

      <form onSubmit={submit} className="flex items-center gap-1.5">
        <input
          type="password"
          value={value}
          disabled={!encryptionAvailable || busy}
          onChange={(event) => setValue(event.target.value)}
          placeholder={configured ? 'Replace the stored key…' : 'Paste a key…'}
          aria-label={`${provider.label} API key`}
          autoComplete="off"
          spellCheck={false}
          className="min-w-0 flex-1 rounded border border-border bg-surface px-2 py-1 font-mono text-[11px] text-text outline-none placeholder:text-text-faint focus:border-focus disabled:opacity-50"
        />
        <button
          type="submit"
          disabled={!encryptionAvailable || busy || value.trim() === ''}
          className="flex items-center gap-1 rounded bg-accent px-2 py-1 text-[11px] font-medium text-accent-text transition-opacity hover:opacity-90 disabled:opacity-40"
        >
          {saved ? <Check size={11} aria-hidden /> : null}
          {saved ? 'Saved' : 'Save'}
        </button>
      </form>

      {error && <p className="text-[11px] leading-relaxed text-danger">{error}</p>}
    </div>
  );
}
