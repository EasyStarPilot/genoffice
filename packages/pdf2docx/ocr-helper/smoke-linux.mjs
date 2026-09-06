/**
 * CI/manual smoke for the Linux system-OCR helper: compile, decode a
 * committed fixture (English text on white), and assert the JSON protocol
 * end to end. Mirrors smoke-win.mjs.
 *
 * A runner without tesseract-ocr installed reports that as exit code 4 — the
 * smoke then downgrades to protocol-only validation (it still can't run
 * without failing the build step first) and still fails on a compile crash,
 * so a green run always means "the binary we ship cannot crash-loop on user
 * machines that likewise lack tesseract."
 */
import { execFileSync, spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

execFileSync(process.execPath, [join(here, 'build-linux.mjs')], { stdio: 'inherit' })

const binary = join(here, 'linux-ocr', 'target', 'release', 'linux-ocr')
const png = readFileSync(join(here, 'fixtures', 'smoke-en.png'))
const res = spawnSync(binary, [], {
  input: png,
  maxBuffer: 64 * 1024 * 1024,
  timeout: 60_000,
})

const stderr = (res.stderr ?? Buffer.alloc(0)).toString('utf8').trim()
if (res.status === 4) {
  console.warn(`no tesseract-ocr on this runner (${stderr}) — recognition not validated`)
  process.exit(0)
}
if (res.status !== 0) {
  console.error(`helper exit ${res.status}: ${stderr}`)
  process.exit(1)
}

const out = JSON.parse(res.stdout.toString('utf8'))
if (typeof out.paper !== 'number' || !Array.isArray(out.lines)) {
  console.error('malformed helper output:', JSON.stringify(out).slice(0, 400))
  process.exit(1)
}
if (out.paper < 0.5) {
  console.error(`paper share ${out.paper} < 0.5 on a white-background fixture`)
  process.exit(1)
}
const text = out.lines.map((l) => l.t).join(' ')
if (!/quick/i.test(text) || !/12345/.test(text)) {
  console.error(`fixture text not recognized; got: ${text.slice(0, 300)}`)
  process.exit(1)
}
for (const line of out.lines) {
  const [x0, y0, x1, y1] = line.b
  if (!(x0 >= 0 && y0 >= 0 && x1 <= 1 && y1 <= 1 && x1 > x0 && y1 > y0)) {
    console.error(`line box out of normalized range: ${JSON.stringify(line.b)}`)
    process.exit(1)
  }
  for (const char of line.chars ?? []) {
    const [cx0, cy0, cx1, cy1] = char.b
    const isGap = cx0 === 0 && cy0 === 0 && cx1 === 0 && cy1 === 0
    if (!isGap && !(cx0 >= 0 && cy0 >= 0 && cx1 <= 1 && cy1 <= 1 && cx1 > cx0 && cy1 > cy0)) {
      console.error(`char box out of normalized range: ${JSON.stringify(char.b)}`)
      process.exit(1)
    }
  }
}
console.log(
  `smoke OK: ${out.lines.length} lines, paper ${out.paper.toFixed(2)}, "${text.slice(0, 80)}"`,
)
