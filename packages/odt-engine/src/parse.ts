/**
 * .odt -> docx-engine's Block/Run model, wrapped as a full ParsedDocFull stub
 * so `blocksToPmDoc` and the rest of apps/docs' load path run completely
 * unmodified (every field besides `blocks` mirrors a brand-new blank
 * document's defaults: no headers/footers/comments/notes/protection/theme).
 *
 * Scope (v1): headings, paragraphs, bullet/ordered lists (shared synthesized
 * numbering — see numbering.ts), tables (plain cell text, no merges/styling),
 * runs with bold/italic/underline/strike/color/size/font, paragraph
 * alignment, and images (inline "as-char" -> Run.image, anchored/floating ->
 * a standalone image Block). Everything else recognized-but-unhandled
 * (footnotes, comments, tracked changes, sections/page setup, frames,
 * drawing shapes other than images) is simply absent, not preserved.
 *
 * No byte-fidelity/dirty-flag bookkeeping: saveOdt (generate.ts) always
 * regenerates the whole content.xml from the live editor state.
 */
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import type { Block, ParsedDocFull, ParaAlign, Run, TableCell, TableModel } from '@genoffice/docx-engine'
import type { OdtPageLayout } from './generate'
import { asXmlNode, xmlArray, type XmlNode } from './xml-utils'
import { parseOdfHalfPoints, parseOdfLengthTwips } from './units'
import { odtNumberingDefs, ODT_BULLET_NUM_ID, ODT_ORDERED_NUM_ID } from './numbering'

const ARRAY_TAGS = new Set([
  'text:p',
  'text:h',
  'text:span',
  'text:list',
  'text:list-item',
  'table:table',
  'table:table-row',
  'table:table-cell',
  'style:style',
  'draw:frame',
])

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  // xml:space="preserve" runs carry meaningful leading/trailing spaces (word
  // boundaries between runs) — trimming would silently glue words together.
  trimValues: false,
  isArray: (name) => ARRAY_TAGS.has(name),
})

const ODT_MIME = 'application/vnd.oasis.opendocument.text'

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
}

export async function parseOdt(bytes: Uint8Array): Promise<ParsedDocFull> {
  const zip = await JSZip.loadAsync(bytes)
  const mimeFile = zip.file('mimetype')
  const mimetype = mimeFile ? (await mimeFile.async('string')).trim() : undefined
  if (mimetype?.startsWith('application/vnd.oasis.opendocument') && mimetype !== ODT_MIME) {
    throw new Error(`OpenDocument file (${mimetype}), not a text document — expected .odt`)
  }
  const contentFile = zip.file('content.xml')
  if (!contentFile) throw new Error('odt: missing content.xml')
  const contentXml = await contentFile.async('string')
  const content = asXmlNode(xmlParser.parse(contentXml))
  const root = asXmlNode(content['office:document-content'])
  const styles = collectStyles(root)
  const listKinds = collectListStyleKinds(root)

  // Named (non-automatic) paragraph/text/list styles commonly live in
  // styles.xml's office:styles, not content.xml's automatic-styles — merge
  // them in (automatic-styles wins on a name collision, the more specific one).
  const stylesFile = zip.file('styles.xml')
  const stylesXml = stylesFile ? (await stylesFile.async('string')) : null
  const pageLayout = parsePageLayout(stylesXml)
  if (stylesFile) {
    const stylesXmlDoc = asXmlNode(xmlParser.parse(stylesXml!))
    const stylesRoot = asXmlNode(stylesXmlDoc['office:document-styles'])
    const officeStyles = asXmlNode(stylesRoot['office:styles'])
    for (const style of xmlArray(officeStyles['style:style'])) {
      const name = style['@_style:name'] as string | undefined
      if (name && !styles.has(name)) styles.set(name, style)
    }
    for (const st of xmlArray(officeStyles['text:list-style'])) {
      const name = st['@_style:name'] as string | undefined
      if (name && !listKinds.has(name)) {
        listKinds.set(name, st['text:list-level-style-number'] ? 'ordered' : 'bullet')
      }
    }
  }

  const body = asXmlNode(root['office:body'])
  const text = asXmlNode(body['office:text'])

  // Pass 1 (sync): collect every image href referenced anywhere in the body.
  const hrefs = new Set<string>()
  collectImageHrefs([text], hrefs)

  // Resolve every referenced image to a data URL up front (async), so the
  // main block-building walk (pass 2) can stay synchronous.
  const media = new Map<string, string>()
  await Promise.all(
    [...hrefs].map(async (href) => {
      const file = zip.file(href)
      if (!file) return
      const raw = await file.async('uint8array')
      const ext = href.slice(href.lastIndexOf('.') + 1).toLowerCase()
      const mime = IMAGE_MIME[ext] ?? 'application/octet-stream'
      media.set(href, `data:${mime};base64,${Buffer.from(raw).toString('base64')}`)
    }),
  )

  let nextId = 0
  const blocks: Block[] = []
  walkBody([text], styles, listKinds, media, null, blocks, () => `odt-b${nextId++}`)

  return {
    blocks,
    comments: [],
    footnotes: [],
    endnotes: [],
    sources: [],
    inks: [],
    protection: null,
    writeProtection: null,
    removePersonalInfo: false,
    styles: new Map(),
    headingStyleIds: new Map(),
    numbering: odtNumberingDefs(),
    internal: {
      originalBytes: bytes,
      documentXml: '',
      bodyInnerStart: 0,
      bodyInnerEnd: 0,
    },
    extras: { elements: [], chartParts: {} },
  }
}

function collectImageHrefs(nodes: readonly XmlNode[], out: Set<string>): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (key === 'draw:image') {
        for (const img of xmlArray(value)) {
          const href = img['@_xlink:href'] as string | undefined
          if (href) out.add(href)
        }
      } else if (Array.isArray(value)) {
        collectImageHrefs(value as XmlNode[], out)
      } else if (value && typeof value === 'object') {
        collectImageHrefs([value as XmlNode], out)
      }
    }
  }
}

/** Parse the page layout (page size + margins) from an ODT file's styles.xml. */
export async function parseOdtPageLayout(bytes: Uint8Array): Promise<OdtPageLayout | undefined> {
  try {
    const zip = await JSZip.loadAsync(bytes)
    const stylesFile = zip.file('styles.xml')
    if (!stylesFile) return undefined
    return parsePageLayout(await stylesFile.async('string'))
  } catch {
    return undefined
  }
}

// ── automatic-styles: style:name -> raw style node ──

function collectStyles(root: XmlNode): Map<string, XmlNode> {
  const map = new Map<string, XmlNode>()
  const autoStyles = asXmlNode(root['office:automatic-styles'])
  for (const style of xmlArray(autoStyles['style:style'])) {
    const name = style['@_style:name'] as string | undefined
    if (name) map.set(name, style)
  }
  return map
}

function parsePageLayout(stylesXml: string | null): OdtPageLayout | undefined {
  if (!stylesXml) return undefined
  try {
    const doc = asXmlNode(xmlParser.parse(stylesXml))
    const root = asXmlNode(doc['office:document-styles'])
    const autoStyles = asXmlNode(root['office:automatic-styles'])
    const layout = xmlArray(autoStyles['style:page-layout'])[0]
    if (!layout) return undefined
    const props = asXmlNode(layout['style:page-layout-properties'])
    const pw = props['@_fo:page-width'] as string | undefined
    const ph = props['@_fo:page-height'] as string | undefined
    const mt = props['@_fo:margin-top'] as string | undefined
    const mb = props['@_fo:margin-bottom'] as string | undefined
    const ml = props['@_fo:margin-left'] as string | undefined
    const mr = props['@_fo:margin-right'] as string | undefined
    if (!pw && !ph) return undefined
    return {
      pageWidth: pw ?? '21.001cm',
      pageHeight: ph ?? '29.7cm',
      marginTop: mt ?? '2cm',
      marginBottom: mb ?? '2cm',
      marginLeft: ml ?? '2cm',
      marginRight: mr ?? '2cm',
    }
  } catch {
    return undefined
  }
}

function styleOf(styles: Map<string, XmlNode>, name: unknown): XmlNode | undefined {
  return typeof name === 'string' ? styles.get(name) : undefined
}

/** text:list-style name -> 'ordered' if its first level is text:list-level-style-number, else 'bullet'. Best-effort: an unresolvable style-name defaults to bullet. */
function collectListStyleKinds(root: XmlNode): Map<string, 'bullet' | 'ordered'> {
  const map = new Map<string, 'bullet' | 'ordered'>()
  const autoStyles = asXmlNode(root['office:automatic-styles'])
  const styleNodes = [...xmlArray(autoStyles['text:list-style'])]
  for (const st of styleNodes) {
    const name = st['@_style:name'] as string | undefined
    if (!name) continue
    map.set(name, st['text:list-level-style-number'] ? 'ordered' : 'bullet')
  }
  return map
}

// ── run / paragraph formatting (same schema as odp-engine — style:text-properties / style:paragraph-properties) ──

const ALIGN_MAP: Record<string, ParaAlign> = {
  start: 'left',
  left: 'left',
  end: 'right',
  right: 'right',
  center: 'center',
  justify: 'justify',
}

function paragraphAlign(style: XmlNode | undefined): ParaAlign | undefined {
  const align = asXmlNode(style?.['style:paragraph-properties'])['@_fo:text-align'] as
    | string
    | undefined
  return align ? ALIGN_MAP[align] : undefined
}

function runPropsFromStyle(style: XmlNode | undefined): Partial<Run> {
  if (!style) return {}
  const tp = asXmlNode(style['style:text-properties'])
  const out: Partial<Run> = {}
  if ((tp['@_fo:font-weight'] as string | undefined) === 'bold') out.bold = true
  const fstyle = tp['@_fo:font-style'] as string | undefined
  if (fstyle === 'italic' || fstyle === 'oblique') out.italic = true
  const underline = tp['@_style:text-underline-style'] as string | undefined
  if (underline && underline !== 'none') out.underline = true
  const strike = tp['@_style:text-line-through-style'] as string | undefined
  if (strike && strike !== 'none') out.strike = true
  const sizeHalfPoints = parseOdfHalfPoints(tp['@_fo:font-size'] as string | undefined)
  if (sizeHalfPoints !== undefined) out.sizeHalfPoints = sizeHalfPoints
  const color = tp['@_fo:color'] as string | undefined
  if (color && color !== 'transparent') out.color = color.replace(/^#/, '')
  const font =
    (tp['@_style:font-name'] as string | undefined) ?? (tp['@_fo:font-family'] as string | undefined)
  if (font) out.font = font.replace(/^['"]|['"]$/g, '')
  return out
}

// ── inline content: text:span/text:tab/text:line-break/draw:frame(as-char image) -> Run[] ──

/**
 * Known limitation: fast-xml-parser (non-preserveOrder mode) groups every
 * same-named child into one array per tag, which loses the relative document
 * order BETWEEN DIFFERENT tag names at the same level (all text:span content
 * comes out before a sibling draw:frame regardless of which was actually
 * first/interspersed in the XML). A paragraph that is uniformly one kind of
 * content (all text:span, or one lone image) round-trips with correct order;
 * only a paragraph mixing inline text spans with an inline ("as-char") image
 * can have that image's position shift relative to its surrounding text.
 * Every run's own text/formatting is still correct — only their relative
 * sequence in this specific mixed case is not guaranteed.
 */
function collectRuns(
  nodes: readonly XmlNode[],
  styles: Map<string, XmlNode>,
  media: Map<string, string>,
  inherited: Partial<Run>,
  out: Run[],
): void {
  for (const node of nodes) {
    if (node == null) continue
    // fast-xml-parser collapses an attribute-less element with only text
    // content to a bare string (e.g. a <text:span> with no style-name) rather
    // than { '#text': ... } — treat it as literal text, not a node to descend into.
    if (typeof node !== 'object') {
      const t = String(node)
      if (t) out.push({ text: t, ...inherited })
      continue
    }
    for (const [key, value] of Object.entries(node)) {
      if (key === '#text') {
        const t = String(value)
        if (t) out.push({ text: t, ...inherited })
      } else if (key === 'text:tab') {
        out.push({ text: '\t', ...inherited })
      } else if (key === 'text:line-break') {
        out.push({ text: '\n', ...inherited })
      } else if (key === 'text:s') {
        out.push({ text: ' ', ...inherited })
      } else if (key === 'text:span') {
        for (const span of xmlArray(value)) {
          const own = { ...inherited, ...runPropsFromStyle(styleOf(styles, span['@_text:style-name'])) }
          collectRuns([span], styles, media, own, out)
        }
      } else if (key === 'draw:frame') {
        for (const frame of xmlArray(value)) {
          const images = xmlArray(frame['draw:image'])
          const href = images[0]?.['@_xlink:href'] as string | undefined
          const dataUrl = href ? media.get(href) : undefined
          if (dataUrl) {
            out.push({
              text: '',
              // xml has no ODF meaning (there is no OOXML fragment to preserve) but must be
              // non-empty: apps/docs' inlineToRuns treats a falsy xml as "not really an
              // image" and silently drops the run when converting back from ProseMirror.
              image: {
                dataUrl,
                xml: '<odt-image/>',
                widthPx: parseFrameWidthPx(frame),
                heightPx: parseFrameHeightPx(frame),
              },
            })
          }
        }
      } else if (Array.isArray(value)) {
        collectRuns(value as XmlNode[], styles, media, inherited, out)
      }
    }
  }
}

/** A paragraph whose only meaningful content is one non-"as-char" draw:frame — a floating/anchored picture, not one mixed inline with text. */
function soleFloatingImage(
  p: XmlNode,
  media: Map<string, string>,
): { dataUrl: string; widthPx?: number; heightPx?: number } | null {
  const frames = xmlArray(p['draw:frame'])
  if (frames.length !== 1) return null
  for (const [key, value] of Object.entries(p)) {
    if (key === 'draw:frame' || key.startsWith('@_')) continue
    if (key === '#text' && !String(value).trim()) continue
    return null // some other real content shares this paragraph — keep it inline
  }
  const frame = frames[0]!
  if ((frame['@_text:anchor-type'] as string | undefined) === 'as-char') return null
  const href = xmlArray(frame['draw:image'])[0]?.['@_xlink:href'] as string | undefined
  const dataUrl = href ? media.get(href) : undefined
  if (!dataUrl) return null
  return { dataUrl, widthPx: parseFrameWidthPx(frame), heightPx: parseFrameHeightPx(frame) }
}

function parseFrameWidthPx(frame: XmlNode): number | undefined {
  const twips = parseOdfLengthTwips(frame['@_svg:width'] as string | undefined)
  return twips ? Math.round((twips / 1440) * 96) : undefined
}
function parseFrameHeightPx(frame: XmlNode): number | undefined {
  const twips = parseOdfLengthTwips(frame['@_svg:height'] as string | undefined)
  return twips ? Math.round((twips / 1440) * 96) : undefined
}

// ── tables ──

function paragraphPlainText(nodes: readonly XmlNode[]): string {
  const runs: Run[] = []
  collectRuns(nodes, new Map(), new Map(), {}, runs)
  return runs.map((r) => r.text).join('')
}

function parseTable(
  table: XmlNode,
  styles: Map<string, XmlNode>,
  media: Map<string, string>,
): TableModel {
  const rows: TableCell[][] = []
  for (const row of xmlArray(table['table:table-row'])) {
    const cells: TableCell[] = []
    for (const cell of xmlArray(row['table:table-cell'])) {
      const paras = xmlArray(cell['text:p']).map((p) => paragraphPlainText([p]))
      const richParas = xmlArray(cell['text:p']).map((p) => {
        const runs: Run[] = []
        collectRuns([p], styles, media, {}, runs)
        return { runs }
      })
      const colSpan = Number(cell['@_table:number-columns-spanned'] ?? 1) || 1
      cells.push({
        paras: paras.length > 0 ? paras : [''],
        ...(richParas.length > 0 ? { richParas } : {}),
        ...(colSpan > 1 ? { colSpan } : {}),
      })
    }
    if (cells.length > 0) rows.push(cells)
  }
  return { rows: rows.length > 0 ? rows : [[{ paras: [''] }]] }
}

// ── body walk: paragraphs/headings/lists/tables/floating images -> Block[] ──

function walkBody(
  nodes: readonly XmlNode[],
  styles: Map<string, XmlNode>,
  listKinds: Map<string, 'bullet' | 'ordered'>,
  media: Map<string, string>,
  list: { kind: 'bullet' | 'ordered'; ilvl: number } | null,
  out: Block[],
  nextId: () => string,
): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'text:h') {
        for (const h of value) {
          const runs: Run[] = []
          collectRuns([h], styles, media, {}, runs)
          const level = Number(h['@_text:outline-level'] ?? 1) || 1
          const align = paragraphAlign(styleOf(styles, h['@_text:style-name']))
          out.push({
            id: nextId(),
            type: 'heading',
            docxIndex: null,
            originalXml: null,
            level: Math.min(9, Math.max(1, level)),
            runs,
            ...(align ? { format: { align } } : {}),
          })
        }
      } else if (key === 'text:p') {
        for (const p of value) {
          // A paragraph whose only content is one non-inline draw:frame is how
          // a floating/anchored picture (as opposed to one mixed inline with
          // text, anchor-type="as-char") is conventionally represented — model
          // it as its own image block rather than a paragraph with one run.
          const floating = list ? null : soleFloatingImage(p, media)
          if (floating) {
            out.push({
              id: nextId(),
              type: 'image',
              docxIndex: null,
              originalXml: null,
              imageDataUrl: floating.dataUrl,
              ...(floating.widthPx ? { imageWidthPx: floating.widthPx } : {}),
              ...(floating.heightPx ? { imageHeightPx: floating.heightPx } : {}),
            })
            continue
          }
          const runs: Run[] = []
          collectRuns([p], styles, media, {}, runs)
          const align = paragraphAlign(styleOf(styles, p['@_text:style-name']))
          if (list) {
            out.push({
              id: nextId(),
              type: 'listItem',
              docxIndex: null,
              originalXml: null,
              list: {
                kind: list.kind,
                numId: list.kind === 'ordered' ? ODT_ORDERED_NUM_ID : ODT_BULLET_NUM_ID,
                ilvl: Math.min(8, list.ilvl),
              },
              runs,
              ...(align ? { format: { align } } : {}),
            })
          } else {
            out.push({
              id: nextId(),
              type: 'paragraph',
              docxIndex: null,
              originalXml: null,
              runs,
              ...(align ? { format: { align } } : {}),
            })
          }
        }
      } else if (key === 'table:table') {
        for (const table of value) {
          out.push({
            id: nextId(),
            type: 'table',
            docxIndex: null,
            originalXml: null,
            table: parseTable(table, styles, media),
          })
        }
      } else if (key === 'text:list') {
        for (const listNode of value) {
          const kind = listKinds.get(listNode['@_text:style-name'] as string) ?? 'bullet'
          const ilvl = list ? list.ilvl + 1 : 0
          for (const item of xmlArray(listNode['text:list-item'])) {
            walkBody([item], styles, listKinds, media, { kind, ilvl }, out, nextId)
          }
        }
      } else if (key === 'draw:frame') {
        // A frame directly under office:text (not inside a text:p) is a
        // paragraph-/page-anchored floating picture — its own image block.
        for (const frame of value) {
          const images = xmlArray(frame['draw:image'])
          const href = images[0]?.['@_xlink:href'] as string | undefined
          const dataUrl = href ? media.get(href) : undefined
          if (dataUrl) {
            out.push({
              id: nextId(),
              type: 'image',
              docxIndex: null,
              originalXml: null,
              imageDataUrl: dataUrl,
              imageWidthPx: parseFrameWidthPx(frame),
              imageHeightPx: parseFrameHeightPx(frame),
            })
          }
        }
      } else {
        walkBody(value as XmlNode[], styles, listKinds, media, list, out, nextId)
      }
    }
  }
}
