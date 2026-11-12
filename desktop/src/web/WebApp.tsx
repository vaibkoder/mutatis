import { useEffect } from 'react';
import { ThemeProvider } from '../renderer/src/theme/ThemeProvider.js';
import { SettingsProvider } from './WebSettingsProvider.js';
import { WebShell } from './WebShell.js';

/**
 * The single-file build's root: App.tsx without the Electron settings provider.
 * There is no window.mutatis here, so nothing may reach for it.
 */
export function WebApp() {
  // A dropped file outside a drop zone makes the browser navigate to it, which
  // would throw away the loaded documents. Drop zones preventDefault themselves.
  useEffect(() => {
    const swallow = (e: DragEvent) => e.preventDefault();
    window.addEventListener('dragover', swallow);
    window.addEventListener('drop', swallow);
    return () => {
      window.removeEventListener('dragover', swallow);
      window.removeEventListener('drop', swallow);
    };
  }, []);

  return (
    <ThemeProvider>
      <SettingsProvider>
        <WebShell />
      </SettingsProvider>
    </ThemeProvider>
  );
}
