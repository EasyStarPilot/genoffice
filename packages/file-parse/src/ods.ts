import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import { tableLines } from './odf-text'

// Needs attributes (table:name) unlike the other ODF extractors; preserveOrder
// still applies so tableLines' shared row/cell walker works unmodified —
// attributes land in a sibling ':@' key, which that walker's Array.isArray
// guard already ignores.
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  preserveOrder: true,
})

interface Section {
  name: string
  lines: string[]
}

function collectSheets(nodes: readonly unknown[], out: Section[]): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    const attrs = (node as Record<string, unknown>)[':@'] as Record<string, unknown> | undefined
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'table:table') {
        const name = String(attrs?.['@_table:name'] ?? '')
        out.push({ name, lines: tableLines(value) })
      } else {
        collectSheets(value, out)
      }
    }
  }
}

/** extract sheet text from an .ods: one "# SheetName" section per sheet, rows " | "-joined */
export async function odsToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file('content.xml')
  if (!file) throw new Error('Invalid ods: missing content.xml')
  const xml = await file.async('text')
  const sheets: Section[] = []
  collectSheets(parser.parse(xml), sheets)
  return sheets.map((s) => [`# ${s.name}`, ...s.lines].join('\n')).join('\n\n')
}
