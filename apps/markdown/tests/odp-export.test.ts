import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseOdp } from '@genoffice/odp-engine'
import type { TextElement } from '@genoffice/pptx-engine'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { splitIntoSlides } from '../src/renderer/export/odpExport'
import { buildOdpBytes } from '../src/main/odp-export'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(md: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

function textElements(el: { elements: TextElement[] }): TextElement[] {
  return el.elements.filter((e): e is TextElement => e.type === 'text')
}

function allText(el: TextElement): string {
  return (el.text?.paragraphs ?? []).map((p) => p.runs.map((r) => r.text).join('')).join('\n')
}

describe('splitIntoSlides', () => {
  it('starts a new slide on each H1, keeping sub-headings as body lines', () => {
    const editor = createEditor('# One\n\nintro\n\n## Sub\n\nmore\n\n# Two\n\nsecond')
    const slides = splitIntoSlides(editor.getJSON())
    expect(slides.map((s) => s.title)).toEqual(['One', 'Two'])
    expect(slides[0]?.lines.map((l) => l.text)).toEqual(['intro', 'Sub', 'more'])
    expect(slides[0]?.lines[1]).toMatchObject({ bold: true })
  })

  it('splits on a horizontal rule even without a heading', () => {
    const editor = createEditor('first\n\n---\n\nsecond')
    const slides = splitIntoSlides(editor.getJSON())
    expect(slides).toHaveLength(2)
    expect(slides[0]?.title).toBeNull()
    expect(slides[0]?.lines.map((l) => l.text)).toEqual(['first'])
    expect(slides[1]?.lines.map((l) => l.text)).toEqual(['second'])
  })

  it('content before the first heading becomes an untitled slide 1', () => {
    const editor = createEditor('lead-in text\n\n# Title')
    const slides = splitIntoSlides(editor.getJSON())
    expect(slides).toHaveLength(2)
    expect(slides[0]).toMatchObject({ title: null })
    expect(slides[0]?.lines.map((l) => l.text)).toEqual(['lead-in text'])
    expect(slides[1]?.title).toBe('Title')
  })

  it('an empty document still yields exactly one (blank) slide', () => {
    const editor = createEditor('')
    const slides = splitIntoSlides(editor.getJSON())
    expect(slides).toEqual([{ title: null, lines: [] }])
  })

  it('list items become bulleted/ordered body lines with their nesting level', () => {
    const editor = createEditor('# T\n\n- one\n  - nested\n\n1. first')
    const slides = splitIntoSlides(editor.getJSON())
    const lines = slides[0]?.lines ?? []
    expect(lines[0]).toMatchObject({ text: 'one', bullet: 'bullet', level: 0 })
    expect(lines[1]).toMatchObject({ text: 'nested', bullet: 'bullet', level: 1 })
    expect(lines[2]).toMatchObject({ text: 'first', bullet: 'ordered', level: 0 })
  })
})

describe('odp export (buildOdpBytes, the main-process byte builder)', () => {
  it('produces one slide per H1 with a real title shape and bullet body', async () => {
    const editor = createEditor('# Intro\n\n- point one\n- point two\n\n# Conclusion\n\nthe end')
    const slides = splitIntoSlides(editor.getJSON())
    const bytes = await buildOdpBytes(slides)
    expect(bytes.length).toBeGreaterThan(500)
    const opened = await parseOdp(bytes)
    expect(opened.deck.slides).toHaveLength(2)

    const slide1 = textElements(opened.deck.slides[0]!)
    expect(slide1.some((e) => allText(e) === 'Intro')).toBe(true)
    const body1 = slide1.find((e) => allText(e).includes('point one'))
    expect(body1 && allText(body1)).toContain('point two')

    const slide2 = textElements(opened.deck.slides[1]!)
    expect(slide2.some((e) => allText(e) === 'Conclusion')).toBe(true)
    expect(slide2.some((e) => allText(e).includes('the end'))).toBe(true)
  })

  it('a document with no heading at all still exports a single valid slide', async () => {
    const editor = createEditor('just some text')
    const slides = splitIntoSlides(editor.getJSON())
    const bytes = await buildOdpBytes(slides)
    const opened = await parseOdp(bytes)
    expect(opened.deck.slides).toHaveLength(1)
    const body = textElements(opened.deck.slides[0]!).find((e) =>
      allText(e).includes('just some text'),
    )
    expect(body).toBeDefined()
  })

  it('an empty slides array still produces one valid blank slide', async () => {
    const bytes = await buildOdpBytes([])
    const opened = await parseOdp(bytes)
    expect(opened.deck.slides).toHaveLength(1)
  })
})
