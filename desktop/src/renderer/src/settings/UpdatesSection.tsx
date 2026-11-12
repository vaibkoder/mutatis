import { useCallback, useEffect, useState } from 'react';
import type { UpdaterState } from '@common/ipc.js';

/**
 * The Updates panel.
 *
 * It is a button, not a banner, and that is the whole design. Nothing here runs
 * unless the user opens Settings and presses something — see
 * src/main/updater.ts. The two-step Check → Download is also deliberate: a
 * version number is cheap, a package is not.
 *
 * When the build cannot update itself the panel says WHY in full sentences
 * rather than hiding or greying out. On an unsigned copy that reason is
 * "macOS will not replace an unsigned application", which is a fact about this
 * download the user is entitled to know before they rely on it.
 */
export function UpdatesSection() {
  const [state, setState] = useState<UpdaterState | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let live = true;
    void window.mutatis.updates.state().then((next) => {
      if (live) setState(next);
    });
    const off = window.mutatis.updates.onState((next) => setState(next));
    return () => {
      live = false;
      off();
    };
  }, []);

  const run = useCallback(async (action: () => Promise<UpdaterState | boolean | null>) => {
    setBusy(true);
    try {
      const next = await action();
      if (next && typeof next === 'object') setState(next);
    } finally {
      setBusy(false);
    }
  }, []);

  if (!state) return <p className="text-[11px] text-text-muted">Loading…</p>;

  const blocked = state.blockers.length > 0;

  return (
    <div className="flex flex-col gap-2" data-updates-section>
      <p className="text-[11px] text-text-muted">
        Version <span className="font-mono" data-app-version>{state.version}</span>
      </p>

      {blocked ? (
        <ul className="flex list-disc flex-col gap-1 pl-4 text-[11px] leading-relaxed text-text-faint">
          {state.blockers.map((reason) => (
            <li key={reason}>{reason}</li>
          ))}
        </ul>
      ) : (
        <div className="flex items-center gap-2">
          <button
            type="button"
            data-check-updates
            disabled={busy || state.phase === 'checking' || state.phase === 'downloading'}
            onClick={() => void run(() => window.mutatis.updates.check())}
            className="rounded border border-border px-2.5 py-1 text-[11px] text-text-muted transition-colors hover:border-border-strong hover:text-text disabled:opacity-50"
          >
            Check for updates
          </button>

          {state.phase === 'available' && (
            <button
              type="button"
              onClick={() => void run(() => window.mutatis.updates.download())}
              className="rounded border border-accent bg-accent-soft px-2.5 py-1 text-[11px] font-medium text-text"
            >
              Download {state.newVersion}
            </button>
          )}

          {state.phase === 'downloaded' && (
            <button
              type="button"
              onClick={() => void run(() => window.mutatis.updates.install())}
              className="rounded border border-accent bg-accent-soft px-2.5 py-1 text-[11px] font-medium text-text"
            >
              Restart and install {state.newVersion}
            </button>
          )}
        </div>
      )}

      <p className="text-[11px] leading-relaxed text-text-faint" data-updates-status>
        {statusLine(state)}
      </p>
    </div>
  );
}

function statusLine(state: UpdaterState): string {
  switch (state.phase) {
    case 'unsupported':
      return 'This copy will not check for updates, and does not contact anything on startup.';
    case 'checking':
      return 'Checking…';
    case 'up-to-date':
      return 'You are on the latest version.';
    case 'available':
      return `Version ${state.newVersion ?? ''} is available. Nothing is downloaded until you say so.`;
    case 'downloading':
      return `Downloading… ${state.percent ?? 0}%`;
    case 'downloaded':
      return `Version ${state.newVersion ?? ''} is ready. It installs when you restart.`;
    case 'error':
      return state.message ?? 'The update check failed.';
    default:
      return 'Updates are checked only when you press the button.';
  }
}
