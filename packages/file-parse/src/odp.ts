import JSZip from 'jszip'
import { collectBlocks, odfTextParser } from './odf-text'

/** Find every draw:page node (in document order) and collect its own text lines. */
function collectSlides(nodes: readonly unknown[], out: string[][]): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'draw:page') {
        const lines: string[] = []
        collectBlocks(value, lines)
        out.push(lines)
      } else {
        collectSlides(value, out)
      }
    }
  }
}

/** extract slide text from an .odp: one "## Slide N" section per slide, a line per paragraph/table row */
export async function odpToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file('content.xml')
  if (!file) throw new Error('Invalid odp: missing content.xml')
  const xml = await file.async('text')
  const slides: string[][] = []
  collectSlides(odfTextParser.parse(xml), slides)
  return slides.map((lines, i) => [`## Slide ${i + 1}`, ...lines].join('\n')).join('\n\n')
}
