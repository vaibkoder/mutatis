import { useEffect } from 'react';
import { SettingsProvider } from './settings/SettingsProvider.js';
import { ThemeProvider } from './theme/ThemeProvider.js';
import { AppShell } from './shell/AppShell.js';

export function App() {
  // Dropping a file anywhere OUTSIDE a drop zone makes Chromium navigate the
  // window to that file — which in Electron means the app is simply gone, with
  // no way back. The drop zones call preventDefault themselves; this catches
  // everywhere else.
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
        <AppShell />
      </SettingsProvider>
    </ThemeProvider>
  );
}
