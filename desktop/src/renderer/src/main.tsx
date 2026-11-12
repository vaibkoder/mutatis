import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { CSPProvider } from '@base-ui/react/csp-provider';
import { App } from './App.js';
import './index.css';

const container = document.getElementById('root');
if (!container) throw new Error('#root missing from index.html');

createRoot(container).render(
  <StrictMode>
    {/*
      disableStyleElements is what lets the app run under `style-src 'self'` with
      no 'unsafe-inline'. The scrollbar rules it suppresses are re-supplied in
      index.css. This is the single decisive reason we are on Base UI and not
      Radix.
    */}
    <CSPProvider disableStyleElements>
      <App />
    </CSPProvider>
  </StrictMode>,
);
