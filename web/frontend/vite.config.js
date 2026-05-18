import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// base: './' -> byggda tillgangar far relativa sokvagar sa FastAPI kan
// servera dist/ direkt fran roten.
export default defineConfig({
  plugins: [react()],
  base: './',
  server: {
    port: 5173,
    proxy: {
      '/api': 'http://127.0.0.1:8765',
    },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  },
})
