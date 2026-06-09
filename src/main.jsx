import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'

// Persistent storage — uses localStorage on Vercel
window.storage = {
  async get(key) {
    try {
      const val = localStorage.getItem('sdrq_' + key)
      return val ? { key, value: val } : null
    } catch { return null }
  },
  async set(key, value) {
    try {
      localStorage.setItem('sdrq_' + key, value)
      return { key, value }
    } catch { return null }
  },
  async delete(key) {
    try {
      localStorage.removeItem('sdrq_' + key)
      return { key, deleted: true }
    } catch { return null }
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
