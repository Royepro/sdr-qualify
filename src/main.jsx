import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'

// ── Storage polyfill: localStorage with in-memory fallback ──
const _mem = {};
window.storage = {
  async get(key) {
    try {
      const raw = localStorage.getItem('sdrq_' + key);
      if (raw) {
        console.log(`[Storage] GET "${key}" — ${raw.length} chars`);
        return { key, value: raw };
      }
      console.log(`[Storage] GET "${key}" — empty`);
      return null;
    } catch (e) {
      console.warn('[Storage] localStorage.getItem failed, using memory:', e.message);
      const v = _mem[key];
      return v ? { key, value: v } : null;
    }
  },
  async set(key, value) {
    _mem[key] = value; // always write to memory too
    try {
      localStorage.setItem('sdrq_' + key, value);
      console.log(`[Storage] SET "${key}" — ${value.length} chars ✓`);
      return { key, value };
    } catch (e) {
      console.warn('[Storage] localStorage.setItem failed (quota?):', e.message);
      return { key, value }; // memory-only fallback
    }
  },
  async delete(key) {
    delete _mem[key];
    try { localStorage.removeItem('sdrq_' + key); } catch {}
    return { key, deleted: true };
  }
};

console.log('[SDR.qualify] Starting on', window.location.hostname);

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
)
