/**
 * Shared text-walking helpers for the OpenDocument family (.odt/.odp/.ods all
 * share the same content.xml container and text:p/text:span/table:table
 * schema) — used by odt.ts/odp.ts/ods.ts. preserveOrder keeps text:span/
 * text:tab/text:line-break in sequence with bare text within a paragraph;
 * grouped by tag name (the default parser mode) they lose that position.
 */
import { XMLParser } from 'fast-xml-parser'

export const odfTextParser = new XMLParser({
  ignoreAttributes: true,
  trimValues: false,
  parseTagValue: false,
  preserveOrder: true,
})

/** Walk one text:p/text:h's children collecting its flat text (spans/tabs/breaks in order). */
export function collectRunText(nodes: readonly unknown[], out: string[]): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (key === '#text') out.push(String(value))
      else if (key === 'text:tab') out.push('\t')
      else if (key === 'text:line-break') out.push('\n')
      else if (key === 'text:s') out.push(' ')
      else if (Array.isArray(value)) collectRunText(value, out)
    }
  }
}

/** Walk a subtree collecting one output line per text:p/text:h, rendering any nested table:table inline as " | "-joined rows. Recurses through list/frame/shape wrappers looking for either. */
export function collectBlocks(nodes: readonly unknown[], out: string[]): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'text:p' || key === 'text:h') {
        const parts: string[] = []
        collectRunText(value, parts)
        const line = parts.join('')
        if (line.trim()) out.push(line)
      } else if (key === 'table:table') {
        out.push(...tableLines(value))
      } else {
        collectBlocks(value, out)
      }
    }
  }
}

/** table:table's children -> one output line per table:table-row, cells joined with " | ". */
export function tableLines(rows: readonly unknown[]): string[] {
  const lines: string[] = []
  const walkRows = (nodes: readonly unknown[]): void => {
    for (const node of nodes) {
      if (node == null || typeof node !== 'object') continue
      for (const [key, value] of Object.entries(node)) {
        if (!Array.isArray(value)) continue
        if (key === 'table:table-row') {
          const cells: string[] = []
          for (const cellNode of value) {
            if (cellNode == null || typeof cellNode !== 'object') continue
            for (const [cellKey, cellValue] of Object.entries(cellNode)) {
              if (cellKey !== 'table:table-cell' || !Array.isArray(cellValue)) continue
              const parts: string[] = []
              collectBlocks(cellValue, parts)
              cells.push(parts.join(' '))
            }
          }
          // A wholly-empty row (e.g. a repeated trailing blank row) adds nothing readable
          if (cells.some((c) => c.trim())) lines.push(cells.join(' | '))
        } else {
          walkRows(value)
        }
      }
    }
  }
  walkRows(rows)
  return lines
}
