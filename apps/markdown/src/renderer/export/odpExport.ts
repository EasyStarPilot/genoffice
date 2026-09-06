/**
 * md document (PM JSON) -> brand-new .odp bytes, fully local. Unlike docx/odt
 * (a flowing document, so the markdown block tree maps to it almost 1:1),
 * markdown has no native notion of a "slide" — this invents one deliberately
 * simple convention, the same one Marp/Slidev/reveal.js-markdown use:
 *
 *   - Each H1 heading starts a new slide and becomes its title.
 *   - A horizontal rule (`---`) also starts a new slide (manual break).
 *   - Everything else (paragraphs, lists, code blocks, blockquotes, tables,
 *     H2-H6) becomes a line in the current slide's body text box; sub-headings
 *     render as a bold body line rather than starting their own slide, so a
 *     deeply-sectioned document doesn't explode into one slide per subhead.
 *
 * Scope: text only. Images have no local from-scratch embed path in this
 * engine (real OOXML picture embedding needs new media parts + relationship
 * bookkeeping this codebase has no from-scratch primitive for yet), so an
 * image becomes a "[Image: alt]" placeholder line rather than being dropped
 * silently or half-embedded.
 *
 * Content is built with pptx-engine's own Paragraph/TextRun model via
 * addElement — the same primitive apps/slides' page-spec.ts uses to render a
 * generated deck from scratch — so slide XML is never hand-assembled here.
 */
import type { JSONContent } from '@tiptap/core'
import { addElement, createBlankPptx, insertBlankSlide, openPptx } from '@genoffice/pptx-engine'
import type { Paragraph, Slide } from '@genoffice/pptx-engine'
import { saveOdp } from '@genoffice/odp-engine'
import { plainText } from './docxExport'

type BulletKind = 'bullet' | 'ordered' | 'none'

interface BodyLine {
  text: string
  level: number
  bullet: BulletKind
  bold?: boolean
  mono?: boolean
}

interface SlideContent {
  title: string | null
  lines: BodyLine[]
}

const MONO_FONT = 'Consolas'

function flattenInlineText(content: JSONContent[] | undefined): string {
  const parts: string[] = []
  for (const child of content ?? []) {
    if (child.type === 'hardBreak') {
      parts.push(' ')
      continue
    }
    if (child.type === 'inlineMath') {
      const latex = String(child.attrs?.latex ?? '')
      if (latex) parts.push(`$${latex}$`)
      continue
    }
    if (child.type === 'text' && child.text) parts.push(child.text)
  }
  return parts.join('')
}

/** Walks the PM doc into a flat SlideContent[], splitting on H1/hr (see file doc comment). */
export function splitIntoSlides(doc: JSONContent): SlideContent[] {
  const slides: SlideContent[] = []
  let current: SlideContent = { title: null, lines: [] }

  const hasContent = () => current.title !== null || current.lines.length > 0
  const startSlide = (title: string | null) => {
    if (hasContent()) slides.push(current)
    current = { title, lines: [] }
  }
  const addLine = (line: BodyLine) => {
    if (line.text.trim() || line.bullet !== 'none') current.lines.push(line)
  }

  function walkList(node: JSONContent, kind: 'bullet' | 'ordered', depth: number): void {
    for (const item of node.content ?? []) {
      if (item.type !== 'listItem' && item.type !== 'taskItem') continue
      for (const child of item.content ?? []) {
        if (child.type === 'paragraph') {
          let text = flattenInlineText(child.content)
          if (item.type === 'taskItem') text = `${item.attrs?.checked ? '☑' : '☐'} ${text}`
          addLine({ text, level: depth, bullet: kind })
        } else if (child.type === 'bulletList' || child.type === 'taskList') {
          walkList(child, 'bullet', depth + 1)
        } else if (child.type === 'orderedList') {
          walkList(child, 'ordered', depth + 1)
        } else {
          walk(child, depth + 1)
        }
      }
    }
  }

  function walk(node: JSONContent, depth: number): void {
    switch (node.type) {
      case 'heading': {
        const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6)
        const text = flattenInlineText(node.content)
        if (level === 1) startSlide(text || null)
        else addLine({ text, level: depth, bullet: 'none', bold: true })
        break
      }
      case 'horizontalRule':
        startSlide(null)
        break
      case 'paragraph':
        addLine({ text: flattenInlineText(node.content), level: depth, bullet: 'none' })
        break
      case 'bulletList':
      case 'taskList':
        walkList(node, 'bullet', depth)
        break
      case 'orderedList':
        walkList(node, 'ordered', depth)
        break
      case 'blockquote':
        for (const child of node.content ?? []) walk(child, depth + 1)
        break
      case 'codeBlock':
        for (const line of plainText(node).split('\n')) {
          addLine({ text: line, level: depth, bullet: 'none', mono: true })
        }
        break
      case 'table':
        for (const row of node.content ?? []) {
          if (row.type !== 'tableRow') continue
          const cells = (row.content ?? []).map((cell) => plainText(cell).trim()).join('  |  ')
          addLine({ text: cells, level: depth, bullet: 'none' })
        }
        break
      case 'image': {
        const alt = String(node.attrs?.alt ?? '') || String(node.attrs?.src ?? '')
        addLine({ text: `[Image: ${alt}]`, level: depth, bullet: 'none' })
        break
      }
      case 'blockMath':
        addLine({
          text: `$$${String(node.attrs?.latex ?? '')}$$`,
          level: depth,
          bullet: 'none',
          mono: true,
        })
        break
      default: {
        const text = plainText(node)
        if (text.trim()) addLine({ text, level: depth, bullet: 'none' })
      }
    }
  }

  for (const node of doc.content ?? []) walk(node, 0)
  if (hasContent() || slides.length === 0) slides.push(current)
  return slides
}

function titleParagraph(text: string): Paragraph {
  return { runs: [{ text, bold: true, fontSize: 32 }], bullet: { type: 'none' } }
}

function bodyParagraph(line: BodyLine): Paragraph {
  const bulleted = line.bullet !== 'none'
  const step = 228600 // 0.25in per indent level, matches the hanging bullet indent
  return {
    runs: [
      {
        text: line.text || ' ',
        fontSize: line.bold ? 20 : 18,
        ...(line.bold ? { bold: true } : {}),
        ...(line.mono ? { fontFamily: MONO_FONT, latinFont: MONO_FONT } : {}),
      },
    ],
    marL: bulleted ? step * (line.level + 1) : step * line.level,
    indent: bulleted ? -step : 0,
    spaceAfter: 8,
    bullet:
      line.bullet === 'bullet'
        ? { type: 'char', char: '•' }
        : line.bullet === 'ordered'
          ? { type: 'number', numType: 'arabicPeriod' }
          : { type: 'none' },
  }
}

const MARGIN_EMU = 457200 // 0.5in
const TITLE_HEIGHT_EMU = 1143000 // 1.25in
const GAP_EMU = 182880 // 0.2in

/** md document (PM JSON) -> brand-new .odp bytes, fully local (see file doc comment for the slide-split rule). */
export async function exportOdpBytes(doc: JSONContent): Promise<Uint8Array> {
  const slides = splitIntoSlides(doc)
  const opened = await openPptx(await createBlankPptx())
  const canvas = opened.deck.size
  const bodyTop = MARGIN_EMU + TITLE_HEIGHT_EMU + GAP_EMU
  const bodyWidth = canvas.cx - MARGIN_EMU * 2

  const targets: Slide[] = [opened.deck.slides[0]!]
  for (let i = 1; i < slides.length; i++) {
    const slide = insertBlankSlide(opened, opened.deck.slides.length - 1)
    if (!slide) throw new Error('odp export: failed to insert slide')
    targets.push(slide)
  }

  slides.forEach((content, i) => {
    const slide = targets[i]!
    if (content.title) {
      addElement(slide, {
        kind: 'textbox',
        offset: { x: MARGIN_EMU, y: MARGIN_EMU, cx: bodyWidth, cy: TITLE_HEIGHT_EMU },
        paragraphs: [titleParagraph(content.title)],
        bodyPr: { wrap: 'square', anchor: 't' },
      })
    }
    if (content.lines.length) {
      addElement(slide, {
        kind: 'textbox',
        offset: { x: MARGIN_EMU, y: bodyTop, cx: bodyWidth, cy: canvas.cy - bodyTop - MARGIN_EMU },
        paragraphs: content.lines.map(bodyParagraph),
        bodyPr: { wrap: 'square', anchor: 't', autoFit: 'shrink' },
      })
    }
  })

  return saveOdp(opened)
}
