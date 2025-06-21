import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
// import basicSsl from '@vitejs/plugin-basic-ssl'

export default defineConfig({
  plugins: [react()],
  server: {
    https: false,
    host: 'localhost',
    port: 5173,
  },
})
