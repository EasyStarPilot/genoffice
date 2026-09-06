/**
 * Node-side OcrEngines backed by the platform system-OCR helper binaries in
 * ocr-helper/ (macOS Vision: vision-ocr.swift; Windows.Media.Ocr: win-ocr.cs;
 * Linux — no built-in OS OCR API, so this shells out to the system
 * `tesseract` binary instead: linux-ocr/src/main.rs). All three speak the
 * same protocol — PNG on stdin, JSON lines with normalized bottom-left boxes
 * and a paper-tone share on stdout — so one spawn wrapper serves every
 * platform.
 *
 * Kept OUT of src/index.ts on purpose: it pulls in node:child_process, which
 * browser/renderer bundles must not see — consumers that run in Node
 * (Electron main, tests, eval scripts) import this module directly and pass
 * the engine via ConvertOptions.ocr.
 */
import { spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { Rect } from './geometry'
import type { OcrEngine, OcrRecognition } from './ocr'

/** helper output caps: a page render is a few MB, JSON a few hundred KB */
const HELPER_MAX_BUFFER = 64 * 1024 * 1024
const HELPER_TIMEOUT_MS = 30_000

interface HelperChar {
  t: string
  b: [number, number, number, number]
}
interface HelperLine {
  t: string
  c: number
  b: [number, number, number, number]
  chars?: HelperChar[]
}

const toRect = (b: [number, number, number, number]): Rect => ({
  x0: b[0],
  y0: b[1],
  x1: b[2],
  y1: b[3],
})

/** spawn wrapper shared by the platform engines (protocol is identical) */
function createHelperOcrEngine(helperPath: string, languages?: string[]): OcrEngine {
  return (png: Uint8Array): OcrRecognition | null => {
    const args = languages && languages.length > 0 ? [languages.join(',')] : []
    const res = spawnSync(helperPath, args, {
      input: png,
      maxBuffer: HELPER_MAX_BUFFER,
      timeout: HELPER_TIMEOUT_MS,
    })
    if (res.status !== 0 || res.stdout == null) return null
    let parsed: { lines?: HelperLine[]; paper?: number }
    try {
      // defense in depth: .NET Framework consoles can prepend a UTF-8 BOM
      const text = res.stdout.toString('utf8').replace(/^\uFEFF/, '')
      parsed = JSON.parse(text) as { lines?: HelperLine[]; paper?: number }
    } catch {
      return null
    }
    if (!Array.isArray(parsed.lines)) return null
    return {
      lines: parsed.lines.map((l) => ({
        text: l.t,
        confidence: l.c,
        box: toRect(l.b),
        ...(l.chars ? { chars: l.chars.map((c) => ({ text: c.t, box: toRect(c.b) })) } : {}),
      })),
      ...(typeof parsed.paper === 'number' ? { paperShare: parsed.paper } : {}),
    }
  }
}

/**
 * Build an OcrEngine that shells out to the compiled macOS Vision helper.
 * Returns null when the platform or binary cannot serve (caller converts
 * without OCR — scanned pages keep the bitmap fallback).
 *
 * @param helperPath absolute path to the compiled vision-ocr binary
 * @param languages  optional recognition hints, e.g. ['zh-Hans', 'en-US'];
 *                   omit to let Vision auto-detect
 */
export function createVisionOcrEngine(helperPath: string, languages?: string[]): OcrEngine | null {
  if (process.platform !== 'darwin') return null
  if (!existsSync(helperPath)) return null
  return createHelperOcrEngine(helperPath, languages)
}

/**
 * Build an OcrEngine that shells out to the compiled Windows.Media.Ocr
 * helper (win-ocr.exe). Same null-on-unavailable contract as the Vision
 * engine; machines without any OCR recognizer language make the helper exit
 * non-zero, which the wrapper already maps to "no result" per page.
 *
 * @param helperPath absolute path to the compiled win-ocr.exe
 * @param languages  optional BCP-47 hints tried in order, e.g. ['zh-Hans-CN'];
 *                   omit to use the user-profile languages
 */
export function createWindowsOcrEngine(helperPath: string, languages?: string[]): OcrEngine | null {
  if (process.platform !== 'win32') return null
  if (!existsSync(helperPath)) return null
  return createHelperOcrEngine(helperPath, languages)
}

/**
 * Build an OcrEngine that shells out to the compiled Linux helper
 * (linux-ocr), which in turn shells out to the system `tesseract` binary.
 * Same null-on-unavailable contract as the other two engines — including
 * when the *helper* exists but the machine has no tesseract-ocr installed:
 * the helper's own exit code 4 already maps to "no result" through the
 * shared spawn wrapper below, same as it does for the Windows helper's
 * no-language-available case.
 *
 * @param helperPath absolute path to the compiled linux-ocr binary
 * @param languages  optional recognition hints, e.g. ['zh-Hans', 'en-US'];
 *                   omit to use every tesseract language installed on the machine
 */
export function createLinuxOcrEngine(helperPath: string, languages?: string[]): OcrEngine | null {
  if (process.platform !== 'linux') return null
  if (!existsSync(helperPath)) return null
  return createHelperOcrEngine(helperPath, languages)
}

/**
 * Shared helper-path resolution for every consumer (apps/pdf's viewer bridge,
 * the shell's pdf2docx-local.ts): packaged under Resources/ocr/<helper> via
 * electron-builder extraResources, repo-relative to ocr-helper/ in dev. The
 * Linux binary alone lives at its normal cargo output path in dev (it isn't
 * flattened into ocr-helper/ like the mac/Windows helpers are) — only the
 * *packaged* copy is flat, matching the other two.
 *
 * @param importMetaUrl the calling module's own `import.meta.url` — the dev
 *   candidate is resolved relative to it, so this only works from a module
 *   four directories above its app's `src/` root (apps/*\/src/main/*.ts),
 *   matching every current caller.
 * @param languages optional recognition hints forwarded to the platform engine
 */
export function resolvePlatformOcrEngine(
  importMetaUrl: string,
  languages?: string[],
): OcrEngine | null {
  const config: { helper: string; devRelative: string; create: typeof createVisionOcrEngine } =
    process.platform === 'darwin'
      ? { helper: 'vision-ocr', devRelative: 'vision-ocr', create: createVisionOcrEngine }
      : process.platform === 'win32'
        ? { helper: 'win-ocr.exe', devRelative: 'win-ocr.exe', create: createWindowsOcrEngine }
        : {
            helper: 'linux-ocr',
            devRelative: 'linux-ocr/target/release/linux-ocr',
            create: createLinuxOcrEngine,
          }
  const here = dirname(fileURLToPath(importMetaUrl))
  // Electron-only global; this package also runs in plain Node (tests, eval
  // scripts) so it isn't in scope, unlike the apps/* callers' tsconfigs.
  const resourcesPath = (process as { resourcesPath?: string }).resourcesPath
  const candidates = [
    ...(resourcesPath ? [join(resourcesPath, 'ocr', config.helper)] : []),
    join(here, '../../../../packages/pdf2docx/ocr-helper', config.devRelative),
  ]
  for (const path of candidates) {
    const engine = config.create(path, languages)
    if (engine) return engine
  }
  return null
}
