import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// npm hoists some @tiptap packages to the repo root (shared with docs at a
// different version) and nests others under this app — dedupe forces every
// import onto this app's single copy so the bundle never carries two cores.
const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]

// @genoffice/i18n and @genoffice/electron-utils ship as TS source — must be
// bundled. @genoffice/odp-engine and its own @genoffice/pptx-engine dependency
// (odp-export.ts's zip/archive work) are the same: raw TS source, no compiled
// dist a plain Node `require` in the packaged app could load.
const MAIN_BUNDLED_DEPS = [
  '@genoffice/i18n',
  '@genoffice/electron-utils',
  '@genoffice/odp-engine',
  '@genoffice/pptx-engine',
]

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin({ exclude: MAIN_BUNDLED_DEPS })],
  },
  preload: {
    // same bundling requirement as main (see comment above)
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  renderer: {
    plugins: [react()],
    resolve: { dedupe: TIPTAP_DEDUPE },
    server: {
      port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
      strictPort: Boolean(process.env.MARKDOWN_DEV_PORT),
    },
  },
})
