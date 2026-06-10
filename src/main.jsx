import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'

// LocalStorage-backed storage polyfill
const _mem = {};
window.storage = {
  async get(key) {
    try {
      const v = localStorage.getItem('sdrq_' + key);
      return v ? { key, value: v } : null;
    } catch { return _mem[key] ? { key, value: _mem[key] } : null; }
  },
  async set(key, value) {
    _mem[key] = value;
    try { localStorage.setItem('sdrq_' + key, value); } catch {}
    return { key, value };
  },
  async delete(key) {
    delete _mem[key];
    try { localStorage.removeItem('sdrq_' + key); } catch {}
    return { key, deleted: true };
  }
};

console.log('[SDR.qualify v3] host:', window.location.hostname);
ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
);
