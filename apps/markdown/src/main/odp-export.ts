/**
 * SlideContent[] (the plain-data slide split from the renderer's
 * odpExport.ts) -> brand-new .odp bytes, fully local. Lives in the main
 * process because pptx-engine/odp-engine need node:crypto/node:fs for zip
 * archive handling, which the renderer (browser) bundle target does not
 * have — see odpExport.ts's file doc comment.
 *
 * Content is built with pptx-engine's own Paragraph/TextRun model via
 * addElement — the same primitive apps/slides' page-spec.ts uses to render a
 * generated deck from scratch — so slide XML is never hand-assembled here.
 */
import { addElement, createBlankPptx, insertBlankSlide, openPptx } from '@genoffice/pptx-engine'
import type { Paragraph, Slide } from '@genoffice/pptx-engine'
import { saveOdp } from '@genoffice/odp-engine'
import type { BodyLine, SlideContent } from '../shared/ipc'

const MONO_FONT = 'Consolas'

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

/** SlideContent[] -> brand-new .odp bytes (see odpExport.ts for the markdown slide-split rule). */
export async function buildOdpBytes(slides: SlideContent[]): Promise<Uint8Array> {
  const ordered = slides.length > 0 ? slides : [{ title: null, lines: [] }]
  const opened = await openPptx(await createBlankPptx())
  const canvas = opened.deck.size
  const bodyTop = MARGIN_EMU + TITLE_HEIGHT_EMU + GAP_EMU
  const bodyWidth = canvas.cx - MARGIN_EMU * 2

  const targets: Slide[] = [opened.deck.slides[0]!]
  for (let i = 1; i < ordered.length; i++) {
    const slide = insertBlankSlide(opened, opened.deck.slides.length - 1)
    if (!slide) throw new Error('odp export: failed to insert slide')
    targets.push(slide)
  }

  ordered.forEach((content, i) => {
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
