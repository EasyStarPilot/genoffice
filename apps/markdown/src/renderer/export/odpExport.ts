/**
 * md document (PM JSON) -> SlideContent[], the plain-data slide split for a
 * .odp export. Markdown has no native notion of a "slide" — this invents one
 * deliberately simple convention, the same one Marp/Slidev/reveal.js-markdown
 * use:
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
 * This module is deliberately Node-free (browser-safe): the actual bytes are
 * built in the main process (see apps/markdown/src/main/odp-export.ts) from
 * the SlideContent[] this produces, sent over IPC — pptx-engine/odp-engine
 * need node:crypto/node:fs for zip archive handling and would break the
 * renderer bundle if imported here (electron-vite build target is the
 * browser, which has no such module).
 */
import type { JSONContent } from '@tiptap/core'
import type { BodyLine, SlideContent } from '../../shared/ipc'
import { plainText } from './docxExport'

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
