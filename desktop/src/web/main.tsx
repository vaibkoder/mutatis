import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { CSPProvider } from '@base-ui/react/csp-provider';
import { WebApp } from './WebApp.js';
import './web.css';

const container = document.getElementById('root');
if (!container) throw new Error('#root missing from index.html');

createRoot(container).render(
  <StrictMode>
    <CSPProvider disableStyleElements>
      <WebApp />
    </CSPProvider>
  </StrictMode>,
);
