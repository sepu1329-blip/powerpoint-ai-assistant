import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import mkcert from 'vite-plugin-mkcert'

// https://vite.dev/config/
export default defineConfig({
  base: '/powerpoint-ai-assistant/',
  plugins: [
    react(),
    mkcert(),
  ],
  server: {
    port: 3000,
    https: {},
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: './index.html',
      },
    },
  },
})
