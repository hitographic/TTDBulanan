import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  // Path relatif agar jalan di GitHub Pages (subpath /TTDBulanan/) maupun domain sendiri
  base: './',
})
