/**
 * .odp -> pptx-engine's Slide/SlideElement model.
 *
 * Scope (v1): slide geometry (position/size/rotation), text boxes and basic
 * autoshapes (solid fill/stroke only) with paragraph/run text formatting
 * (bold/italic/underline/strike/size/color/font, paragraph alignment), and
 * pictures. Everything else recognized-but-unhandled (tables, charts,
 * embedded objects, groups, connectors) becomes a PassthroughElement so the
 * editor shows a placeholder chip instead of losing the shape silently —
 * mirroring pptx-engine's own passthrough philosophy for its own no-go areas.
 *
 * No byte-fidelity/dirty-flag bookkeeping: saveOdp always regenerates the
 * whole content.xml from the current in-memory model (see generate.ts).
 */
import { XMLParser } from 'fast-xml-parser'
import { PackageArchive } from '@genoffice/pptx-engine'
import type {
  Fill,
  OpenedPptx,
  Paragraph,
  PassthroughElement,
  PictureElement,
  Slide,
  SlideElement,
  SlideSize,
  Stroke,
  TextAlign,
  TextElement,
  TextRun,
} from '@genoffice/pptx-engine'
import { asXmlNode, xmlArray, xmlText, type XmlNode } from './xml-utils'
import { DEFAULT_SLIDE_SIZE, odfRadiansToOoxmlRot, parseOdfLength, parseOdfPt } from './units'

const ARRAY_TAGS = new Set([
  'draw:page',
  'draw:frame',
  'draw:custom-shape',
  'draw:rect',
  'draw:ellipse',
  'draw:g',
  'draw:image',
  'text:p',
  'text:span',
  'text:line-break',
  'style:style',
  'style:page-layout',
  'style:master-page',
])

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (name) => ARRAY_TAGS.has(name),
})

const ODP_MIME = 'application/vnd.oasis.opendocument.presentation'

export async function parseOdp(bytes: Uint8Array): Promise<OpenedPptx> {
  const archive = await PackageArchive.open(bytes)
  const mimetype = archive.readText('mimetype')?.trim()
  if (mimetype && !mimetype.startsWith('application/vnd.oasis.opendocument')) {
    throw new Error(`not an OpenDocument file: mimetype is ${mimetype}`)
  }
  if (mimetype && mimetype !== ODP_MIME) {
    throw new Error(`OpenDocument file (${mimetype}), not a presentation — expected .odp`)
  }
  const contentXml = archive.readText('content.xml')
  if (!contentXml) throw new Error('odp: missing content.xml')

  const size = readSlideSize(archive.readText('styles.xml'))
  const content = asXmlNode(xmlParser.parse(contentXml))
  const root = asXmlNode(content['office:document-content'])
  const styles = collectStyles(root)

  const body = asXmlNode(root['office:body'])
  const presentation = asXmlNode(body['office:presentation'])
  const pages = xmlArray(presentation['draw:page'])

  const slides: Slide[] = pages.map((page, i) => parsePage(page, i, styles))

  return { deck: { slides, size, originalHash: archive.originalHash }, archive }
}

// ── styles.xml: slide size only (per-shape/paragraph styles live in content.xml's automatic-styles) ──

function readSlideSize(stylesXml: string | null): SlideSize {
  if (!stylesXml) return DEFAULT_SLIDE_SIZE
  try {
    const doc = asXmlNode(xmlParser.parse(stylesXml))
    const root = asXmlNode(doc['office:document-styles'])
    const autoStyles = asXmlNode(root['office:automatic-styles'])
    const layout = xmlArray(autoStyles['style:page-layout'])[0]
    if (!layout) return DEFAULT_SLIDE_SIZE
    const props = asXmlNode(layout['style:page-layout-properties'])
    const cx = parseOdfLength(props['@_fo:page-width'] as string | undefined)
    const cy = parseOdfLength(props['@_fo:page-height'] as string | undefined)
    return cx > 0 && cy > 0 ? { cx, cy } : DEFAULT_SLIDE_SIZE
  } catch {
    return DEFAULT_SLIDE_SIZE
  }
}

// ── content.xml automatic-styles: style:name -> raw style node ──

function collectStyles(root: XmlNode): Map<string, XmlNode> {
  const map = new Map<string, XmlNode>()
  const autoStyles = asXmlNode(root['office:automatic-styles'])
  for (const style of xmlArray(autoStyles['style:style'])) {
    const name = style['@_style:name'] as string | undefined
    if (name) map.set(name, style)
  }
  return map
}

function styleOf(styles: Map<string, XmlNode>, name: unknown): XmlNode | undefined {
  return typeof name === 'string' ? styles.get(name) : undefined
}

// ── geometry ──

interface Geometry {
  x: number
  y: number
  cx: number
  cy: number
  rot: number
}

/** svg:x/y/width/height (+ optional draw:transform="rotate(rad) translate(x y)") -> pptx-engine's unrotated-rect + rot pair. */
function parseGeometry(node: XmlNode): Geometry {
  const cx = parseOdfLength(node['@_svg:width'] as string | undefined)
  const cy = parseOdfLength(node['@_svg:height'] as string | undefined)
  const transform = node['@_draw:transform'] as string | undefined
  const rotMatch = transform ? /rotate\s*\(\s*(-?[\d.]+)/.exec(transform) : null
  const translateMatch = transform ? /translate\s*\(\s*(-?[\d.]+\S*)\s+(-?[\d.]+\S*)\s*\)/.exec(transform) : null
  if (rotMatch) {
    const rad = Number.parseFloat(rotMatch[1]!)
    const tx = translateMatch ? parseOdfLength(translateMatch[1]) : parseOdfLength(node['@_svg:x'] as string | undefined)
    const ty = translateMatch ? parseOdfLength(translateMatch[2]) : parseOdfLength(node['@_svg:y'] as string | undefined)
    // Inverse of generate.ts's placement formula: the composed transform maps
    // local point (lx,ly) to R(rad)*(lx+tx, ly+ty); the shape's own center is
    // at local (cx/2, cy/2).
    const cos = Math.cos(rad)
    const sin = Math.sin(rad)
    const lx = tx + cx / 2
    const ly = ty + cy / 2
    const centerX = lx * cos + ly * sin
    const centerY = -lx * sin + ly * cos
    return { x: centerX - cx / 2, y: centerY - cy / 2, cx, cy, rot: odfRadiansToOoxmlRot(rad) }
  }
  const x = parseOdfLength(node['@_svg:x'] as string | undefined)
  const y = parseOdfLength(node['@_svg:y'] as string | undefined)
  return { x, y, cx, cy, rot: 0 }
}

// ── fill / stroke (solid only — gradients approximate to their fill-color when present, hatch/bitmap fall back to no fill) ──

function parseFill(props: XmlNode): Fill | undefined {
  const kind = props['@_draw:fill'] as string | undefined
  if (kind === 'none') return { type: 'none' }
  if (kind === 'solid' || kind === 'gradient') {
    const color = props['@_draw:fill-color'] as string | undefined
    if (color) return { type: 'solid', color }
  }
  return undefined
}

function parseStroke(props: XmlNode): Stroke | undefined {
  const kind = (props['@_draw:stroke'] as string | undefined) ?? 'none'
  if (kind === 'none') return undefined
  const color = (props['@_svg:stroke-color'] as string) || '#000000'
  const width = parseOdfLength(props['@_svg:stroke-width'] as string | undefined) || 12700
  return { fill: { type: 'solid', color }, width }
}

function graphicFillStroke(style: XmlNode | undefined): { fill?: Fill; stroke?: Stroke } {
  if (!style) return {}
  const props = asXmlNode(style['style:graphic-properties'])
  const fill = parseFill(props)
  const stroke = parseStroke(props)
  return { ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}) }
}

// ── text ──

const ALIGN_MAP: Record<string, TextAlign> = {
  start: 'left',
  left: 'left',
  end: 'right',
  right: 'right',
  center: 'center',
  justify: 'justify',
}

function paragraphAlign(style: XmlNode | undefined): TextAlign | undefined {
  const align = asXmlNode(style?.['style:paragraph-properties'])['@_fo:text-align'] as
    | string
    | undefined
  return align ? ALIGN_MAP[align] : undefined
}

function runPropsFromStyle(style: XmlNode | undefined): Partial<TextRun> {
  if (!style) return {}
  const tp = asXmlNode(style['style:text-properties'])
  const out: Partial<TextRun> = {}
  if ((tp['@_fo:font-weight'] as string | undefined) === 'bold') out.bold = true
  const fstyle = tp['@_fo:font-style'] as string | undefined
  if (fstyle === 'italic' || fstyle === 'oblique') out.italic = true
  const underline = tp['@_style:text-underline-style'] as string | undefined
  if (underline && underline !== 'none') out.underline = true
  const strike = tp['@_style:text-line-through-style'] as string | undefined
  if (strike && strike !== 'none') out.strike = true
  const size = parseOdfPt(tp['@_fo:font-size'] as string | undefined)
  if (size !== undefined) out.fontSize = size
  const color = tp['@_fo:color'] as string | undefined
  if (color && color !== 'transparent') out.color = color
  const font =
    (tp['@_style:font-name'] as string | undefined) ?? (tp['@_fo:font-family'] as string | undefined)
  if (font) out.fontFamily = font.replace(/^['"]|['"]$/g, '')
  return out
}

/**
 * Flatten a text:span's children into runs. Known limitation: fast-xml-parser
 * collapses same-name repeated children into one array, losing the relative
 * order between bare text and <text:span> siblings when both appear at the
 * same level (e.g. "before <span>bold</span> after") — spans come out in
 * their own document order, but interleaved bare text does not. The common
 * cases (a uniformly-formatted paragraph, or one fully wrapped in spans) are
 * unaffected; only fine-grained inline formatting can reorder.
 */
function spanRuns(span: XmlNode, styles: Map<string, XmlNode>, inherited: Partial<TextRun>): TextRun[] {
  const own = { ...inherited, ...runPropsFromStyle(styleOf(styles, span['@_text:style-name'])) }
  const nested = xmlArray(span['text:span'])
  if (nested.length > 0) return nested.flatMap((s) => spanRuns(s, styles, own))
  const text = xmlText(span)
  return text ? [{ text, ...own }] : []
}

function paragraphRuns(p: XmlNode, styles: Map<string, XmlNode>): TextRun[] {
  const spans = xmlArray(p['text:span'])
  if (spans.length > 0) return spans.flatMap((s) => spanRuns(s, styles, {}))
  const text = xmlText(p)
  return text ? [{ text }] : []
}

/** A text:p may contain text:line-break; pptx-engine's own model treats soft breaks as separate paragraphs (TextRun doc comment), so one text:p can yield several Paragraphs. */
function parseTextBodyParagraphs(paragraphs: XmlNode[], styles: Map<string, XmlNode>): Paragraph[] {
  const out: Paragraph[] = []
  for (const p of paragraphs) {
    const align = paragraphAlign(styleOf(styles, p['@_text:style-name']))
    const breaks = xmlArray(p['text:line-break']).length
    const runs = paragraphRuns(p, styles)
    if (breaks === 0) {
      out.push({ runs, ...(align ? { align } : {}) })
      continue
    }
    // Best-effort: line-break position within the run sequence isn't tracked
    // by fast-xml-parser's collapsed structure either, so split the whole
    // paragraph's runs evenly is wrong; instead keep all text on the first
    // line and emit empty trailing lines for the remaining breaks — visually
    // preserves line count without inventing text placement.
    out.push({ runs, ...(align ? { align } : {}) })
    for (let i = 0; i < breaks; i++) out.push({ runs: [], ...(align ? { align } : {}) })
  }
  return out.length > 0 ? out : [{ runs: [] }]
}

// ── shapes ──

let passthroughId = 0

function passthrough(
  geo: Geometry,
  kind: PassthroughElement['kind'],
  spIndex: number,
): PassthroughElement {
  return {
    id: `odp-el-${spIndex}-${passthroughId++}`,
    type: 'passthrough',
    kind,
    anchor: { spIndex, originalXml: '', range: [0, 0] },
    transform: { offset: { x: geo.x, y: geo.y, cx: geo.cx, cy: geo.cy }, rot: geo.rot, flipH: false, flipV: false },
  }
}

function parseTextBoxFrame(
  frame: XmlNode,
  textBox: XmlNode,
  styles: Map<string, XmlNode>,
  spIndex: number,
): TextElement {
  const geo = parseGeometry(frame)
  const style = styleOf(styles, frame['@_draw:style-name'])
  const { fill, stroke } = graphicFillStroke(style)
  const paragraphs = parseTextBodyParagraphs(xmlArray(textBox['text:p']), styles)
  return {
    id: `odp-el-${spIndex}`,
    type: 'text',
    anchor: { spIndex, originalXml: '', range: [0, 0] },
    transform: { offset: { x: geo.x, y: geo.y, cx: geo.cx, cy: geo.cy }, rot: geo.rot, flipH: false, flipV: false },
    ...(fill ? { fill } : {}),
    ...(stroke ? { stroke } : {}),
    text: { paragraphs },
  }
}

function parseCustomShape(shape: XmlNode, styles: Map<string, XmlNode>, spIndex: number): TextElement {
  const geo = parseGeometry(shape)
  const style = styleOf(styles, shape['@_draw:style-name'])
  const { fill, stroke } = graphicFillStroke(style)
  const paragraphs = parseTextBodyParagraphs(xmlArray(shape['text:p']), styles)
  return {
    id: `odp-el-${spIndex}`,
    type: 'shape',
    presetGeometry: 'rect',
    anchor: { spIndex, originalXml: '', range: [0, 0] },
    transform: { offset: { x: geo.x, y: geo.y, cx: geo.cx, cy: geo.cy }, rot: geo.rot, flipH: false, flipV: false },
    ...(fill ? { fill } : {}),
    ...(stroke ? { stroke } : {}),
    text: { paragraphs },
  }
}

function parsePrimitiveShape(
  shape: XmlNode,
  preset: 'rect' | 'ellipse',
  styles: Map<string, XmlNode>,
  spIndex: number,
): TextElement {
  const geo = parseGeometry(shape)
  const style = styleOf(styles, shape['@_draw:style-name'])
  const { fill, stroke } = graphicFillStroke(style)
  const paragraphs = parseTextBodyParagraphs(xmlArray(shape['text:p']), styles)
  return {
    id: `odp-el-${spIndex}`,
    type: 'shape',
    presetGeometry: preset,
    anchor: { spIndex, originalXml: '', range: [0, 0] },
    transform: { offset: { x: geo.x, y: geo.y, cx: geo.cx, cy: geo.cy }, rot: geo.rot, flipH: false, flipV: false },
    ...(fill ? { fill } : {}),
    ...(stroke ? { stroke } : {}),
    text: { paragraphs },
  }
}

function parsePictureFrame(frame: XmlNode, image: XmlNode, spIndex: number): PictureElement {
  const geo = parseGeometry(frame)
  const href = (image['@_xlink:href'] as string | undefined) ?? ''
  return {
    id: `odp-el-${spIndex}`,
    type: 'picture',
    mediaRef: href,
    anchor: { spIndex, originalXml: '', range: [0, 0] },
    transform: { offset: { x: geo.x, y: geo.y, cx: geo.cx, cy: geo.cy }, rot: geo.rot, flipH: false, flipV: false },
  }
}

function parseFrame(frame: XmlNode, styles: Map<string, XmlNode>, spIndex: number): SlideElement {
  if (frame['draw:text-box']) {
    return parseTextBoxFrame(frame, asXmlNode(frame['draw:text-box']), styles, spIndex)
  }
  const images = xmlArray(frame['draw:image'])
  if (images[0]) return parsePictureFrame(frame, images[0], spIndex)
  if (frame['table:table']) return passthrough(parseGeometry(frame), 'table', spIndex)
  if (frame['draw:object'] || frame['draw:object-ole']) {
    return passthrough(parseGeometry(frame), 'ole', spIndex)
  }
  return passthrough(parseGeometry(frame), 'unknown', spIndex)
}

function parsePageElements(page: XmlNode, styles: Map<string, XmlNode>): SlideElement[] {
  const out: SlideElement[] = []
  let spIndex = 0
  for (const frame of xmlArray(page['draw:frame'])) out.push(parseFrame(frame, styles, spIndex++))
  for (const shape of xmlArray(page['draw:custom-shape'])) {
    out.push(parseCustomShape(shape, styles, spIndex++))
  }
  for (const rect of xmlArray(page['draw:rect'])) out.push(parsePrimitiveShape(rect, 'rect', styles, spIndex++))
  for (const ellipse of xmlArray(page['draw:ellipse'])) {
    out.push(parsePrimitiveShape(ellipse, 'ellipse', styles, spIndex++))
  }
  for (const group of xmlArray(page['draw:g'])) out.push(passthrough(parseGeometry(group), 'unknown', spIndex++))
  return out
}

function pageBackground(page: XmlNode, styles: Map<string, XmlNode>): Fill | undefined {
  const style = styleOf(styles, page['@_draw:style-name'])
  const props = asXmlNode(style?.['style:drawing-page-properties'])
  return parseFill(props)
}

function parsePage(page: XmlNode, index: number, styles: Map<string, XmlNode>): Slide {
  const background = pageBackground(page, styles)
  return {
    path: `content.xml#page${index + 1}`,
    originalXml: '',
    bodyPrefix: '',
    bodySuffix: '',
    elements: parsePageElements(page, styles),
    ...(background ? { background } : {}),
  }
}
