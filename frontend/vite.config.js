import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/upload': { target: 'http://localhost:3001', changeOrigin: true },
      '/generate-report': { target: 'http://localhost:3001', changeOrigin: true },
      '/chat': {
        target: 'http://localhost:3001',
        changeOrigin: true,
        // SSE requires no response buffering
        configure: (proxy) => {
          proxy.on('proxyReq', (proxyReq) => {
            proxyReq.setHeader('Accept', 'text/event-stream');
          });
        },
      },
      '/slides': { target: 'http://localhost:3001', changeOrigin: true },
      '/report': { target: 'http://localhost:3001', changeOrigin: true },
      '/health': { target: 'http://localhost:3001', changeOrigin: true },
    },
  },
})
