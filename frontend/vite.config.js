import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import path from 'path'

export default defineConfig({
  plugins: [vue()],
  server: {
    port: 5174
  },
  build: {
    outDir: path.resolve(__dirname, 'dist'),
    assetsDir: 'assets',
    sourcemap: false,
    minify: 'esbuild'
  }
})
