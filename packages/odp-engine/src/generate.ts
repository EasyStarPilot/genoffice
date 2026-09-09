/**
 * pptx-engine's Slide/SlideElement model -> .odp bytes.
 *
 * No byte-fidelity: every save fully regenerates content.xml from the current
 * in-memory model (there is no ODF equivalent of pptx-engine's byte-anchor
 * patch strategy to reuse, and building one would buy nothing since this
 * package never round-trips through the *original* .odp bytes as text). Media
 * already referenced by a canonical `Pictures/imageN.ext` path is kept as-is
 * (idempotent across repeated saves); anything else (freshly inserted via
 * pptx-engine's own addPicture, which writes ppt/media/imageN.ext) is
 * re-embedded under a new canonical path.
 */
import JSZip from 'jszip'
import type {
  Fill,
  OpenedPptx,
  Paragraph,
  PictureElement,
  Slide,
  SlideElement,
  Stroke,
  TextAlign,
  TextElement,
  TextRun,
  Transform,
} from '@genoffice/pptx-engine'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { emuToOdfLength, ooxmlRotToOdfRadians, ptToOdfLength } from './units'

const ODP_MIME = 'application/vnd.oasis.opendocument.presentation'

export async function saveOdp(opened: OpenedPptx): Promise<Uint8Array> {
  const zip = buildOdpZip(opened)
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

/** Same output as saveOdp, streamed straight to disk (mirrors pptx-engine's savePptxToFile). */
export async function saveOdpToFile(opened: OpenedPptx, filePath: string): Promise<void> {
  const { createWriteStream } = await import('node:fs')
  const { pipeline } = await import('node:stream/promises')
  const zip = buildOdpZip(opened)
  const source = zip.generateNodeStream({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    streamFiles: true,
  })
  await pipeline(source, createWriteStream(filePath))
}

/** Clears pptx-engine's edit dirty flags (set by its editing operations, which this app reuses verbatim regardless of source format) now that a save has captured the current state. Call after a successful saveOdp/saveOdpToFile. */
export function commitOdpSaved(opened: OpenedPptx): void {
  for (const slide of opened.deck.slides) {
    delete slide.structureDirty
    for (const el of slide.elements) {
      delete el.dirty
      delete el.dirtyTransform
      delete el.dirtyFill
      delete el.dirtyStroke
      delete el.dirtySrcRect
      delete el.dirtyPPr
    }
  }
}

// ── media: assign canonical Pictures/imageN.ext paths, idempotently ──

const CANONICAL_MEDIA_RE = /^Pictures\/image(\d+)\.(\w+)$/i

const MEDIA_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

interface MediaPlan {
  /** old mediaRef -> new canonical path (identity for already-canonical refs) */
  pathFor: Map<string, string>
  manifestEntries: Array<{ path: string; mediaType: string }>
}

function planMedia(opened: OpenedPptx): MediaPlan {
  const pathFor = new Map<string, string>()
  const manifestEntries: Array<{ path: string; mediaType: string }> = []
  let used = 0
  for (const m of opened.archive.entries.keys()) {
    const match = CANONICAL_MEDIA_RE.exec(m)
    if (match) used = Math.max(used, Number.parseInt(match[1]!, 10))
  }
  for (const slide of opened.deck.slides) {
    for (const el of slide.elements) {
      if (el.type !== 'picture') continue
      const ref = el.mediaRef
      if (pathFor.has(ref)) continue
      const already = CANONICAL_MEDIA_RE.exec(ref)
      const path = already ? ref : (() => {
        const ext = ref.includes('.') ? ref.slice(ref.lastIndexOf('.') + 1).toLowerCase() : 'png'
        return `Pictures/image${++used}.${ext}`
      })()
      pathFor.set(ref, path)
      const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
      manifestEntries.push({ path, mediaType: MEDIA_MIME[ext] ?? 'application/octet-stream' })
    }
  }
  return { pathFor, manifestEntries }
}

// ── fill / stroke -> ODF attrs (solid + gradient) ──

class GradientAcc {
  private fragments: string[] = []
  private n = 0
  private names = new Map<string, string>()

  getGradient(fill: Fill): string | undefined {
    if (fill.type !== 'gradient') return undefined
    const key = `${fill.stops.map((s) => `${s.pos}:${s.color}`).join(',')}|${fill.angle ?? 0}`
    const existing = this.names.get(key)
    if (existing) return existing
    const name = `Grad${++this.n}`
    this.names.set(key, name)
    const startColor = fill.stops[0]?.color ?? '#000000'
    const endColor = fill.stops[fill.stops.length - 1]?.color ?? '#ffffff'
    const angle = fill.angle ?? 0
    this.fragments.push(
      `<draw:gradient draw:name="${escapeXmlAttr(name)}" draw:style="linear" ` +
      `draw:start-color="${escapeXmlAttr(startColor)}" draw:end-color="${escapeXmlAttr(endColor)}" ` +
      `draw:angle="${angle}"/>`,
    )
    return name
  }

  toXml(): string {
    return this.fragments.length > 0
      ? `<office:styles>${this.fragments.join('')}</office:styles>`
      : '<office:styles/>'
  }
}

function fillAttrs(fill: Fill | undefined, gradAcc?: GradientAcc): string {
  if (!fill || fill.type === 'none') return 'draw:fill="none"'
  if (fill.type === 'solid') return `draw:fill="solid" draw:fill-color="${escapeXmlAttr(fill.color)}"`
  if (fill.type === 'gradient') {
    const gradName = gradAcc?.getGradient(fill)
    if (gradName) {
      return `draw:fill="gradient" draw:fill-gradient-name="${escapeXmlAttr(gradName)}"`
    }
    // Fallback: approximate to first stop color
    const color = fill.stops[0]?.color ?? '#808080'
    return `draw:fill="solid" draw:fill-color="${escapeXmlAttr(color)}"`
  }
  return 'draw:fill="none"'
}

function strokeAttrs(stroke: Stroke | undefined): string {
  if (!stroke || stroke.fill.type === 'none') return 'draw:stroke="none"'
  const color = stroke.fill.type === 'solid' ? stroke.fill.color : '#000000'
  return `draw:stroke="solid" svg:stroke-color="${escapeXmlAttr(color)}" svg:stroke-width="${emuToOdfLength(stroke.width)}"`
}

// ── automatic-styles accumulator ──

class StyleAcc {
  private fragments: string[] = []
  private n = 0
  readonly gradients = new GradientAcc()
  private id(prefix: string): string {
    return `${prefix}${++this.n}`
  }
  graphic(fill: Fill | undefined, stroke: Stroke | undefined): string {
    const name = this.id('gr')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${fillAttrs(fill, this.gradients)} ${strokeAttrs(stroke)} draw:textarea-horizontal-align="center" draw:textarea-vertical-align="middle" draw:auto-grow-height="false"/></style:style>`,
    )
    return name
  }
  drawingPage(fill: Fill | undefined): string {
    const name = this.id('dp')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="drawing-page"><style:drawing-page-properties ${fillAttrs(fill, this.gradients)}/></style:style>`,
    )
    return name
  }
  paragraph(align: TextAlign | undefined): string | undefined {
    if (!align) return undefined
    const name = this.id('P')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties fo:text-align="${ALIGN_TO_ODF[align]}"/></style:style>`,
    )
    return name
  }
  text(run: TextRun): string | undefined {
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
    if (run.strike) attrs.push('style:text-line-through-style="solid"', 'style:text-line-through-type="single"')
    if (run.fontSize !== undefined) attrs.push(`fo:font-size="${ptToOdfLength(run.fontSize)}"`)
    if (run.color) attrs.push(`fo:color="${escapeXmlAttr(run.color)}"`)
    if (run.fontFamily) attrs.push(`style:font-name="${escapeXmlAttr(run.fontFamily)}"`)
    if (attrs.length === 0) return undefined
    const name = this.id('T')
    this.fragments.push(
      `<style:style style:name="${name}" style:family="text"><style:text-properties ${attrs.join(' ')}/></style:style>`,
    )
    return name
  }
  toXml(): string {
    return this.fragments.join('')
  }
}

const ALIGN_TO_ODF: Record<TextAlign, string> = {
  left: 'start',
  right: 'end',
  center: 'center',
  justify: 'justify',
}

// ── geometry ──

function geometryAttrs(transform: Transform): string {
  const { offset, rot } = transform
  if (!rot) {
    return `svg:x="${emuToOdfLength(offset.x)}" svg:y="${emuToOdfLength(offset.y)}" svg:width="${emuToOdfLength(offset.cx)}" svg:height="${emuToOdfLength(offset.cy)}"`
  }
  const rad = ooxmlRotToOdfRadians(rot)
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  const centerPageX = offset.x + offset.cx / 2
  const centerPageY = offset.y + offset.cy / 2
  // Invert parse.ts's read-back exactly (proven algebraically, see units.ts):
  // local top-left-relative-to-transform = R(rad)^-1 * pageCenter, R^-1 = R(-rad).
  const lx = centerPageX * cos - centerPageY * sin
  const ly = centerPageX * sin + centerPageY * cos
  const tx = lx - offset.cx / 2
  const ty = ly - offset.cy / 2
  const radStr = (Math.round(rad * 1e6) / 1e6).toString()
  return `svg:x="0cm" svg:y="0cm" svg:width="${emuToOdfLength(offset.cx)}" svg:height="${emuToOdfLength(offset.cy)}" draw:transform="rotate (${radStr}) translate (${emuToOdfLength(tx)} ${emuToOdfLength(ty)})"`
}

// ── text ──

function runXml(run: TextRun, acc: StyleAcc): string {
  const styleName = acc.text(run)
  const attr = styleName ? ` text:style-name="${styleName}"` : ''
  return `<text:span${attr}>${escapeXmlText(run.text)}</text:span>`
}

function paragraphXml(p: Paragraph, acc: StyleAcc): string {
  const styleName = acc.paragraph(p.align)
  const attr = styleName ? ` text:style-name="${styleName}"` : ''
  const body = p.runs.length > 0 ? p.runs.map((r) => runXml(r, acc)).join('') : ''
  return `<text:p${attr}>${body}</text:p>`
}

function textBodyXml(paragraphs: Paragraph[], acc: StyleAcc): string {
  return paragraphs.map((p) => paragraphXml(p, acc)).join('')
}

// ── elements ──

function textElementXml(el: TextElement, acc: StyleAcc): string {
  const styleName = acc.graphic(el.fill, el.stroke)
  const geo = geometryAttrs(el.transform)
  const paragraphs = el.text?.paragraphs ?? []
  if (el.type === 'text') {
    return `<draw:frame draw:style-name="${styleName}" ${geo}><draw:text-box>${textBodyXml(paragraphs, acc)}</draw:text-box></draw:frame>`
  }
  // 'shape' — a v1 autoshape always round-trips as a rectangle (presetGeometry
  // beyond rect/ellipse isn't modeled by the parser, see parse.ts)
  const tag = el.presetGeometry === 'ellipse' ? 'draw:ellipse' : 'draw:rect'
  return `<${tag} draw:style-name="${styleName}" ${geo}>${textBodyXml(paragraphs, acc)}</${tag}>`
}

function pictureElementXml(el: PictureElement, acc: StyleAcc, mediaPath: string): string {
  const styleName = acc.graphic(undefined, undefined)
  const geo = geometryAttrs(el.transform)
  return `<draw:frame draw:style-name="${styleName}" ${geo}><draw:image xlink:href="${escapeXmlAttr(mediaPath)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`
}

function elementXml(el: SlideElement, acc: StyleAcc, media: MediaPlan): string {
  if (el.type === 'text' || el.type === 'shape') return textElementXml(el, acc)
  if (el.type === 'picture') {
    const path = media.pathFor.get(el.mediaRef) ?? el.mediaRef
    return pictureElementXml(el, acc, path)
  }
  // table/chart/group/passthrough: emit the original XML preserved during
  // parse so these elements survive a round-trip even though this engine
  // cannot edit them. If no original XML was captured (e.g. a newly-created
  // passthrough), emit nothing rather than invalid output.
  if (el.type === 'passthrough' && el.anchor.originalXml) {
    return el.anchor.originalXml
  }
  return ''
}

function slideXml(slide: Slide, acc: StyleAcc, media: MediaPlan): string {
  const bgStyle = slide.background ? acc.drawingPage(slide.background) : undefined
  const styleAttr = bgStyle ? ` draw:style-name="${bgStyle}"` : ''
  const elements = slide.elements.map((el) => elementXml(el, acc, media)).join('')
  return `<draw:page draw:master-page-name="Default"${styleAttr}>${elements}</draw:page>`
}

// ── package assembly ──

function contentXml(opened: OpenedPptx, media: MediaPlan): string {
  const acc = new StyleAcc()
  const pages = opened.deck.slides.map((s) => slideXml(s, acc, media)).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-content ' +
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
    'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
    'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" ' +
    'office:version="1.2">' +
    acc.gradients.toXml() +
    '<office:automatic-styles>' +
    acc.toXml() +
    '</office:automatic-styles>' +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    '</office:document-content>'
  )
}

function stylesXml(opened: OpenedPptx): string {
  const { cx, cy } = opened.deck.size
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-styles ' +
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
    'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" ' +
    'office:version="1.2">' +
    '<office:styles/>' +
    '<office:automatic-styles>' +
    `<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="${emuToOdfLength(cx)}" fo:page-height="${emuToOdfLength(cy)}" style:print-orientation="landscape"/></style:page-layout>` +
    '</office:automatic-styles>' +
    '<office:master-styles>' +
    '<style:master-page style:name="Default" style:page-layout-name="PM1"/>' +
    '</office:master-styles>' +
    '</office:document-styles>'
  )
}

function metaXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
    'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">' +
    `<office:meta><meta:generator>GenOffice</meta:generator></office:meta>` +
    '</office:document-meta>'
  )
}

function manifestXml(media: MediaPlan): string {
  const entries = media.manifestEntries
    .map(
      (m) =>
        `<manifest:file-entry manifest:full-path="${escapeXmlAttr(m.path)}" manifest:media-type="${m.mediaType}"/>`,
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">' +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODP_MIME}"/>` +
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>' +
    entries +
    '</manifest:manifest>'
  )
}

function buildOdpZip(opened: OpenedPptx): JSZip {
  const media = planMedia(opened)
  const zip = new JSZip()
  // mimetype must be the first entry and stored uncompressed (ODF package rule)
  zip.file('mimetype', ODP_MIME, { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', manifestXml(media))
  zip.file('meta.xml', metaXml())
  zip.file('styles.xml', stylesXml(opened))
  zip.file('content.xml', contentXml(opened, media))
  for (const [oldRef, newPath] of media.pathFor) {
    const bytes = opened.archive.readBytes(oldRef)
    if (bytes) zip.file(newPath, bytes)
  }
  // Keep mediaRef consistent with what was just written, so a second save
  // without a reopen re-embeds the same (now-canonical) path idempotently.
  for (const slide of opened.deck.slides) {
    for (const el of slide.elements) {
      if (el.type === 'picture') {
        const mapped = media.pathFor.get(el.mediaRef)
        if (mapped) el.mediaRef = mapped
      }
    }
  }
  return zip
}
