import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
import { libInjectCss } from 'vite-plugin-lib-inject-css'

// Library-Modus: baut die Komponente als ES- und CJS-Modul nach dist/.
// react und react-dom werden NICHT mitgebündelt (external), das Zielprojekt liefert sie.
export default defineConfig({
  plugins: [react(), libInjectCss()],
  build: {
    lib: {
      entry: fileURLToPath(new URL('./src/index.js', import.meta.url)),
      name: 'MyBargraph',
      formats: ['es', 'cjs'],
      fileName: (format) => `mybargraph.${format}.js`,
    },
    rollupOptions: {
      external: ['react', 'react-dom', 'react/jsx-runtime', 'prop-types'],
      output: {
        // Default- und Named-Export zusammen, ohne Warnung
        exports: 'named',
        globals: {
          react: 'React',
          'react-dom': 'ReactDOM',
          'react/jsx-runtime': 'jsxRuntime',
        },
      },
    },
  },
})
