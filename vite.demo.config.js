import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Nur für die lokale Demo (npm run demo). Wird nicht publiziert.
export default defineConfig({
  root: 'demo',
  plugins: [react()],
})
