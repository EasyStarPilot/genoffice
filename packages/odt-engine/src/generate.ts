/**
 * Plain block data -> .odt bytes. No byte-fidelity: every save fully
 * regenerates content.xml from whatever the caller hands in (there is no odt
 * equivalent of docx-engine's byte-anchor patch strategy to reuse, and this
 * package never round-trips the original .odt bytes as text).
 *
 * Callers materialize their live editor state into the input shape below
 * themselves (in apps/docs, via docx-engine's own exported
 * pmNodeToGeneratedBlock/pmTableToModel — this package only turns already-
 * plain data into ODF XML, it knows nothing about ProseMirror).
 */
import JSZip from 'jszip'
import type { GeneratedBlock, ParaAlign, Run, TableModel } from '@genoffice/docx-engine'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { halfPointsToOdfPt } from './units'

export type OdtSaveBlock =
  | { kind: 'text'; block: GeneratedBlock }
  | { kind: 'table'; model: TableModel }
  | {
      kind: 'image'
      dataUrl: string
      widthPx?: number
      heightPx?: number
      align?: 'left' | 'center' | 'right'
    }

const ODT_MIME = 'application/vnd.oasis.opendocument.text'

const ALIGN_TO_ODF: Record<ParaAlign, string> = {
  left: 'start',
  right: 'end',
  center: 'center',
  justify: 'justify',
  distribute: 'justify',
}

interface PendingImage {
  token: string
  mime: string
  base64: string
}

/** Accumulates automatic-styles fragments and pending embedded images while walking the blocks; both get resolved once, at the end, in saveOdt. */
class StyleAcc {
  private fragments: string[] = []
  private n = 0
  readonly images: PendingImage[] = []
  private listStyleNames: Partial<Record<'bullet' | 'ordered', string>> = {}
  private id(prefix: string): string {
    return `${prefix}${++this.n}`
  }
  paragraph(align: ParaAlign | undefined): string | undefined {
    if (!align) return undefined
    const name = this.id('P')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties fo:text-align="${ALIGN_TO_ODF[align]}"/></style:style>`,
    )
    return name
  }
  text(run: Run): string | undefined {
    const attrs: string[] = []
    if (run.bold) attrs.push('fo:font-weight="bold"')
    if (run.italic) attrs.push('fo:font-style="italic"')
    if (run.underline) {
      attrs.push(
        'style:text-underline-style="solid"',
        'style:text-underline-width="auto"',
        'style:text-underline-color="font-color"',
      )
    }
    if (run.strike) {
      attrs.push('style:text-line-through-style="solid"', 'style:text-line-through-type="single"')
    }
    if (run.sizeHalfPoints !== undefined) {
      attrs.push(`fo:font-size="${halfPointsToOdfPt(run.sizeHalfPoints)}"`)
    }
    if (run.color) attrs.push(`fo:color="#${run.color.replace(/^#/, '')}"`)
    if (run.font) attrs.push(`style:font-name="${escapeXmlAttr(run.font)}"`)
    if (attrs.length === 0) return undefined
    const name = this.id('T')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="text"><style:text-properties ${attrs.join(' ')}/></style:style>`,
    )
    return name
  }
  /** graphic-properties style for a draw:frame's size */
  frame(widthPx: number | undefined, heightPx: number | undefined): string {
    const name = this.id('fr')
    const w = widthPx ? `${((widthPx / 96) * 2.54).toFixed(4)}cm` : '5cm'
    const h = heightPx ? `${((heightPx / 96) * 2.54).toFixed(4)}cm` : '5cm'
    this.fragments.push(
      `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties style:wrap="none" svg:width="${w}" svg:height="${h}"/></style:style>`,
    )
    return name
  }
  /** One shared list-style per kind (bullet/ordered), lazily defined on first use — parseOdt's collectListStyleKinds resolves it back via the same text:list-level-style-number/-bullet distinction. */
  listStyle(kind: 'bullet' | 'ordered'): string {
    const existing = this.listStyleNames[kind]
    if (existing) return existing
    const name = this.id(kind === 'ordered' ? 'LO' : 'LB')
    const levels = Array.from({ length: 9 }, (_, i) => {
      const level = i + 1
      const space = `text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:margin-left="${(0.5 * level).toFixed(2)}cm" fo:text-indent="-0.5cm"/>`
      return kind === 'ordered'
        ? `<text:list-level-style-number text:level="${level}" style:num-format="1" style:num-suffix="."><style:list-level-properties ${space}</style:list-level-properties></text:list-level-style-number>`
        : `<text:list-level-style-bullet text:level="${level}" text:bullet-char="•"><style:list-level-properties ${space}</style:list-level-properties></text:list-level-style-bullet>`
    }).join('')
    this.fragments.push(`<text:list-style style:name="${name}">${levels}</text:list-style>`)
    this.listStyleNames[kind] = name
    return name
  }
  /** Registers an embedded image (data: URL) and returns a placeholder href token, resolved into a real Pictures/imageN.ext path once in saveOdt (ODF has no way to inline image bytes directly in a draw:image href). */
  registerImage(dataUrl: string): string | null {
    const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl)
    if (!m) return null
    const token = `\0ODT-IMG-${this.images.length}\0`
    this.images.push({ token, mime: m[1]!, base64: m[2]! })
    return token
  }
  toXml(): string {
    return this.fragments.join('')
  }
}

function drawImageXml(dataUrl: string, styleName: string, anchorType: 'as-char' | 'paragraph', acc: StyleAcc): string {
  const token = acc.registerImage(dataUrl)
  if (!token) return ''
  return (
    `<draw:frame draw:style-name="${styleName}" text:anchor-type="${anchorType}">` +
    `<draw:image xlink:href="${token}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    '</draw:frame>'
  )
}

function runXml(run: Run, acc: StyleAcc): string {
  if (run.image) {
    const styleName = acc.frame(run.image.widthPx, run.image.heightPx)
    return drawImageXml(run.image.dataUrl, styleName, 'as-char', acc)
  }
  if (!run.text) return ''
  const styleName = acc.text(run)
  const attr = styleName ? ` text:style-name="${styleName}"` : ''
  // \t / \n round-trip as tab / line-break; ODF has no run-internal escape for either
  const parts = run.text.split(/(\t|\n)/).map((seg) => {
    if (seg === '\t') return '<text:tab/>'
    if (seg === '\n') return '<text:line-break/>'
    return escapeXmlText(seg)
  })
  return `<text:span${attr}>${parts.join('')}</text:span>`
}

function paragraphLikeXml(tag: 'text:p' | 'text:h', block: GeneratedBlock, acc: StyleAcc): string {
  const styleName = acc.paragraph(block.format?.align)
  const attrs = styleName ? ` text:style-name="${styleName}"` : ''
  const levelAttr = tag === 'text:h' ? ` text:outline-level="${block.level ?? 1}"` : ''
  const body = block.runs.map((r) => runXml(r, acc)).join('')
  return `<${tag}${attrs}${levelAttr}>${body}</${tag}>`
}

function tableXml(model: TableModel): string {
  const rows = model.rows
    .map((row) => {
      const cells = row
        .map((cell) => {
          const paras =
            cell.paras.length > 0
              ? cell.paras.map((t) => `<text:p>${escapeXmlText(t)}</text:p>`).join('')
              : '<text:p/>'
          const span =
            cell.colSpan && cell.colSpan > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : ''
          return `<table:table-cell office:value-type="string"${span}>${paras}</table:table-cell>`
        })
        .join('')
      return `<table:table-row>${cells}</table:table-row>`
    })
    .join('')
  const cols = model.rows[0]?.length ?? 1
  const gridCols = Array.from({ length: cols }, () => '<table:table-column/>').join('')
  return `<table:table>${gridCols}${rows}</table:table>`
}

interface BlockEntry {
  xml: string
  isListItem: boolean
  ilvl: number
  kind: 'bullet' | 'ordered'
}

function blockXml(block: OdtSaveBlock, acc: StyleAcc): BlockEntry {
  if (block.kind === 'table')
    return { xml: tableXml(block.model), isListItem: false, ilvl: 0, kind: 'bullet' }
  if (block.kind === 'image') {
    const styleName = acc.frame(block.widthPx, block.heightPx)
    const frame = drawImageXml(block.dataUrl, styleName, 'paragraph', acc)
    return { xml: `<text:p>${frame}</text:p>`, isListItem: false, ilvl: 0, kind: 'bullet' }
  }
  const g = block.block
  if (g.type === 'heading')
    return { xml: paragraphLikeXml('text:h', g, acc), isListItem: false, ilvl: 0, kind: 'bullet' }
  if (g.type === 'listItem') {
    // Left bare (not yet wrapped in text:list-item) — wrapLists nests each
    // item's XML inside the tree it builds from consecutive items' ilvls.
    return {
      xml: paragraphLikeXml('text:p', g, acc),
      isListItem: true,
      ilvl: g.list?.ilvl ?? 0,
      kind: g.list?.kind ?? 'bullet',
    }
  }
  return { xml: paragraphLikeXml('text:p', g, acc), isListItem: false, ilvl: 0, kind: 'bullet' }
}

interface ListForestNode {
  itemXml: string
  kind: 'bullet' | 'ordered'
  children: ListForestNode[]
}

/**
 * A deeper item nests *inside* the text:list-item of the last item at the
 * level above it (ODF's content model: text:list's only children are
 * text:list-item/text:list-header, so a nested text:list must itself be a
 * child of one of those, never its sibling) — not a flat sequence of opening/
 * closing text:list tags around consecutive items, which produces a bare
 * text:list sibling to text:list-item and is not schema-valid.
 */
function buildListForest(
  items: Array<{ xml: string; ilvl: number; kind: 'bullet' | 'ordered' }>,
): ListForestNode[] {
  const roots: ListForestNode[] = []
  const stack: ListForestNode[] = [] // stack[i] = the currently-open node at depth i
  for (const item of items) {
    const node: ListForestNode = { itemXml: item.xml, kind: item.kind, children: [] }
    const depth = Math.max(0, item.ilvl)
    if (depth === 0 || stack.length === 0) {
      roots.push(node)
      stack.length = 0
      stack.push(node)
    } else {
      // A depth jump deeper than the current stack (e.g. 0 -> 2, skipping 1)
      // clamps to the deepest available parent rather than inventing empty levels.
      const parentIndex = Math.min(depth, stack.length) - 1
      stack[parentIndex]!.children.push(node)
      stack.length = parentIndex + 1
      stack.push(node)
    }
  }
  return roots
}

/**
 * Emits one <text:list style-name=...> per maximal run of same-kind siblings
 * (a kind change — or the top-level call itself — starts a new list), each
 * item's own children recursing into their own nested text:list. A single
 * ilvl run can freely mix bullet and ordered sub-lists this way, and every
 * <text:list> carries the style-name a real ODF consumer (and parseOdt's own
 * collectListStyleKinds) needs to tell them apart — a bare, style-less
 * text:list has no way to say "this one was ordered".
 */
function serializeListForest(nodes: ListForestNode[], acc: StyleAcc): string {
  const out: string[] = []
  let run: ListForestNode[] = []
  const flush = () => {
    if (run.length === 0) return
    const styleName = acc.listStyle(run[0]!.kind)
    const items = run
      .map((n) => {
        const nested = n.children.length > 0 ? serializeListForest(n.children, acc) : ''
        return `<text:list-item>${n.itemXml}${nested}</text:list-item>`
      })
      .join('')
    out.push(`<text:list text:style-name="${styleName}">${items}</text:list>`)
    run = []
  }
  for (const node of nodes) {
    if (run.length > 0 && run[run.length - 1]!.kind !== node.kind) flush()
    run.push(node)
  }
  flush()
  return out.join('')
}

/** Group consecutive listItem blocks into one (possibly nested) text:list per run; everything else passes through. */
function wrapLists(entries: BlockEntry[], acc: StyleAcc): string {
  const out: string[] = []
  let run: Array<{ xml: string; ilvl: number; kind: 'bullet' | 'ordered' }> = []
  const flush = () => {
    if (run.length === 0) return
    out.push(serializeListForest(buildListForest(run), acc))
    run = []
  }
  for (const entry of entries) {
    if (!entry.isListItem) {
      flush()
      out.push(entry.xml)
      continue
    }
    run.push({ xml: entry.xml, ilvl: entry.ilvl, kind: entry.kind })
  }
  flush()
  return out.join('')
}

function contentXml(blocks: OdtSaveBlock[], acc: StyleAcc): string {
  const entries = blocks.map((b) => blockXml(b, acc))
  const body = wrapLists(entries, acc)
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-content ' +
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
    'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
    'office:version="1.2">' +
    `<office:automatic-styles>${acc.toXml()}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body>` +
    '</office:document-content>'
  )
}

function stylesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
    'office:version="1.2">' +
    '<office:styles/>' +
    '<office:automatic-styles>' +
    '<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>' +
    '</office:automatic-styles>' +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"/></office:master-styles>' +
    '</office:document-styles>'
  )
}

function metaXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">' +
    '<office:meta><meta:generator>GenOffice</meta:generator></office:meta>' +
    '</office:document-meta>'
  )
}

function manifestXml(media: Array<{ path: string; mime: string }>): string {
  const entries = media
    .map(
      (m) =>
        `<manifest:file-entry manifest:full-path="${escapeXmlAttr(m.path)}" manifest:media-type="${m.mime}"/>`,
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">' +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIME}"/>` +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    entries +
    '</manifest:manifest>'
  )
}

export async function saveOdt(blocks: OdtSaveBlock[]): Promise<Uint8Array> {
  const acc = new StyleAcc()
  let xml = contentXml(blocks, acc)

  const zip = new JSZip()
  const media: Array<{ path: string; mime: string }> = []
  acc.images.forEach((img, i) => {
    const ext = img.mime.split('/')[1]?.replace('svg+xml', 'svg') ?? 'png'
    const path = `Pictures/image${i + 1}.${ext}`
    zip.file(path, img.base64, { base64: true })
    media.push({ path, mime: img.mime })
    xml = xml.split(img.token).join(path)
  })

  zip.file('mimetype', ODT_MIME, { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', manifestXml(media))
  zip.file('meta.xml', metaXml())
  zip.file('styles.xml', stylesXml())
  zip.file('content.xml', xml)
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

export async function saveOdtToFile(blocks: OdtSaveBlock[], filePath: string): Promise<void> {
  const { writeFile } = await import('node:fs/promises')
  await writeFile(filePath, await saveOdt(blocks))
}
