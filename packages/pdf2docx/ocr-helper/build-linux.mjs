/**
 * Compile the Linux system-OCR helper (the linux-ocr crate -> target/release/
 * linux-ocr) via cargo — no manual staleness check needed, unlike
 * build-win.mjs's csc.exe invocation: cargo's own incremental build already
 * makes a no-op rebuild fast.
 *
 * Used by apps/shell/electron-builder.cjs (packaging preflight) and
 * smoke-linux.mjs (manual/CI check).
 */
import { execFileSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

if (process.platform !== 'linux') {
  console.error('build-linux.mjs only runs on Linux')
  process.exit(1)
}

execFileSync(
  'cargo',
  ['build', '--release', '--manifest-path', join(here, 'linux-ocr', 'Cargo.toml')],
  { stdio: 'inherit' },
)
console.log('built', join(here, 'linux-ocr', 'target', 'release', 'linux-ocr'))
