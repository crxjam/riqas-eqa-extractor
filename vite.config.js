import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

const isGhPages = process.env.GH_PAGES === '1'

export default defineConfig({
  plugins: [react()],
  base: isGhPages ? '/riqas-eqa-extractor/' : '/',
})
