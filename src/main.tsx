import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'

if (typeof window !== 'undefined') {
  window.addEventListener('error', (e) => {
    console.error("GLOBAL_ERROR_CAUGHT", e.message, e.error?.stack);
    const div = document.createElement('div');
    div.style.color = 'red';
    div.style.background = '#ffe4e6';
    div.style.border = '2px solid #f43f5e';
    div.style.padding = '15px';
    div.style.margin = '15px';
    div.style.zIndex = '99999';
    div.style.position = 'relative';
    div.style.fontSize = '14px';
    div.innerText = `UNHANDLED ERROR: ${e.message}\nStack: ${e.error?.stack || 'no stack'}`;
    document.body.prepend(div);
  });

  window.addEventListener('unhandledrejection', (e) => {
    console.error("GLOBAL_REJECTION_CAUGHT", e.reason);
    const div = document.createElement('div');
    div.style.color = 'red';
    div.style.background = '#ffe4e6';
    div.style.border = '2px solid #f43f5e';
    div.style.padding = '15px';
    div.style.margin = '15px';
    div.style.zIndex = '99999';
    div.style.position = 'relative';
    div.style.fontSize = '14px';
    div.innerText = `UNHANDLED REJECTION: ${e.reason}`;
    document.body.prepend(div);
  });
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
