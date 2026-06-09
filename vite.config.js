import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/pipedrive': {
        target: 'https://api.pipedrive.com',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/pipedrive/, '/v1')
      }
    }
  }
})
