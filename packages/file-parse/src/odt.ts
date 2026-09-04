import JSZip from 'jszip'
import { collectBlocks, odfTextParser } from './odf-text'

/** extract body text from an .odt: one line per paragraph/heading, tables rendered as " | "-joined rows */
export async function odtToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file('content.xml')
  if (!file) throw new Error('Invalid odt: missing content.xml')
  const xml = await file.async('text')
  const lines: string[] = []
  collectBlocks(odfTextParser.parse(xml), lines)
  return lines.join('\n')
}
