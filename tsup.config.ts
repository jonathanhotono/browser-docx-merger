import { defineConfig } from 'tsup'

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm', 'iife'],
  outDir: 'dist',
  dts: true,
  minify: true,
  globalName: 'DocxMerger',
  // Ensure the IIFE properly exposes the global variable
  platform: 'browser',
  // Add explicit banner to ensure global assignment
  banner: {
    js: 'var DocxMerger;'
  },
  // Explicitly set the global export for IIFE
  esbuildOptions(options) {
    if (options.format === 'iife') {
      options.footer = {
        js: 'if (typeof window !== "undefined") { window.DocxMerger = DocxMerger; }'
      }
    }
  }
})