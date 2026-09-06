import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseOdt } from '@genoffice/odt-engine'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { exportOdtBytes, mapDocToOdtSaveBlocks } from '../src/renderer/export/odtExport'

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

const noImages = () => Promise.resolve(null)

/** md → .odt bytes → parseOdt: the exported file must be a valid OpenDocument Text file */
async function exportAndParse(md: string) {
  const editor = createEditor(md)
  const bytes = await exportOdtBytes(editor.getJSON(), noImages)
  expect(bytes.length).toBeGreaterThan(500)
  return parseOdt(bytes)
}

describe('odt export', () => {
  it('headings, paragraphs and inline marks survive into the odt', async () => {
    const parsed = await exportAndParse('# Title\n\nSome **bold** and *italic* and `code` text.')
    const heading = parsed.blocks.find((b) => b.type === 'heading')
    expect(heading?.level).toBe(1)
    expect(heading?.runs?.map((r) => r.text).join('')).toBe('Title')
    const para = parsed.blocks.find((b) => b.type === 'paragraph')
    const runs = para?.runs ?? []
    expect(runs.find((r) => r.bold)?.text).toBe('bold')
    expect(runs.find((r) => r.italic)?.text).toBe('italic')
    expect(runs.map((r) => r.text).join('')).toContain('code')
  })

  it('bullet and ordered lists become native odt list items', async () => {
    const parsed = await exportAndParse('- one\n- two\n\n1. first\n2. second')
    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items.length).toBe(4)
    expect(items.filter((b) => b.list?.kind === 'bullet').length).toBe(2)
    expect(items.filter((b) => b.list?.kind === 'ordered').length).toBe(2)
  })

  it('nested list levels carry ilvl', async () => {
    const parsed = await exportAndParse('- top\n  - nested')
    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items[0]?.list?.ilvl).toBe(0)
    expect(items[1]?.list?.ilvl).toBe(1)
  })

  it('separate ordered lists each restart numbering (odt auto-restarts per contiguous run)', async () => {
    const parsed = await exportAndParse('1. a\n\ntext between\n\n1. b')
    const items = parsed.blocks.filter((b) => b.type === 'listItem')
    expect(items.map((b) => b.runs?.map((r) => r.text).join(''))).toEqual(['a', 'b'])
  })

  it('tables become native odt tables', async () => {
    const parsed = await exportAndParse('| Name | Value |\n| --- | --- |\n| a | 1 |')
    const table = parsed.blocks.find((b) => b.type === 'table')
    expect(table?.table?.rows.length).toBe(2)
    expect(table?.table?.rows[0]?.[0]?.paras.join('')).toBe('Name')
    expect(table?.table?.rows[1]?.[1]?.paras.join('')).toBe('1')
  })

  it('block math has no ODF equation form and always keeps its LaTeX source visible', async () => {
    const editor = createEditor('$$\n\\frac{a}{b}\n$$')
    const blocks = await mapDocToOdtSaveBlocks(editor.getJSON(), noImages)
    const texts = blocks.map((b) =>
      b.kind === 'text' ? (b.block.runs ?? []).map((r) => r.text).join('') : '',
    )
    expect(texts.join('\n')).toContain('$$\\frac{a}{b}$$')
  })

  it('inline math keeps its LaTeX in the run text', async () => {
    const editor = createEditor('value $x_{1}$ end')
    const blocks = await mapDocToOdtSaveBlocks(editor.getJSON(), noImages)
    const texts = blocks.map((b) =>
      b.kind === 'text' ? (b.block.runs ?? []).map((r) => r.text).join('') : '',
    )
    expect(texts.join('\n')).toContain('$x_{1}$')
  })

  it('task lists render checkbox glyphs', async () => {
    const parsed = await exportAndParse('- [x] done\n- [ ] open')
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.some((t) => t.startsWith('☑'))).toBe(true)
    expect(texts.some((t) => t.startsWith('☐'))).toBe(true)
  })

  it('code blocks keep their text in a monospace paragraph', async () => {
    const parsed = await exportAndParse('```js\nconst a = 1\nconst b = 2\n```')
    const code = parsed.blocks.find((b) =>
      (b.runs ?? []).some((r) => r.text.includes('const a = 1')),
    )
    expect(code).toBeDefined()
    expect(code?.runs?.[0]?.font).toBe('Consolas')
    // odt has no run-internal line-break escape: the embedded \n round-trips as
    // its own text:line-break run, so the two lines land in separate runs.
    expect(code?.runs?.map((r) => r.text).join('')).toContain('const b = 2')
  })

  it('unresolvable images fall back to alt text', async () => {
    const parsed = await exportAndParse('![architecture diagram](assets/missing.png)')
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.some((t) => t.includes('architecture diagram'))).toBe(true)
  })

  it('a resolvable image becomes an embedded odt image block', async () => {
    const tinyPng =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
    const editor = createEditor('![alt](pic.png)')
    const bytes = await exportOdtBytes(editor.getJSON(), () =>
      Promise.resolve({ base64: tinyPng, mime: 'image/png', widthPx: 1, heightPx: 1 }),
    )
    const parsed = await parseOdt(bytes)
    const image = parsed.blocks.find((b) => b.type === 'image')
    expect(image?.imageDataUrl).toContain('base64,')
  })
})
