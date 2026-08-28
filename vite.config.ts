import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './',
  server: {
    proxy: {
      // Forward /api/* to the local Express proxy (npm run server) in dev mode.
      '/api': 'http://localhost:3001',
    },
  },
})
