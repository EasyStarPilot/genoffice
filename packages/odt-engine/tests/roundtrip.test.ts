import { describe, expect, it } from 'vitest'
import { parseOdt } from '../src/parse'
import { saveOdt, type OdtSaveBlock } from '../src/generate'
import { ODT_BULLET_NUM_ID, ODT_ORDERED_NUM_ID } from '../src/numbering'

// A 1x1 red PNG, matching the fixture convention used elsewhere in this repo.
const TINY_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

describe('odt round-trip: plain blocks -> saveOdt -> parseOdt', () => {
  it('preserves a heading, a formatted paragraph, and paragraph alignment', async () => {
    const blocks: OdtSaveBlock[] = [
      { kind: 'text', block: { type: 'heading', level: 2, format: { align: 'center' }, runs: [{ text: 'Title' }] } },
      {
        kind: 'text',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'Hello ' },
            { text: 'bold', bold: true, color: '112233', sizeHalfPoints: 28, font: 'Georgia' },
            { text: ' world' },
          ],
        },
      },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    expect(parsed.blocks).toHaveLength(2)
    expect(parsed.blocks[0]).toMatchObject({ type: 'heading', level: 2, format: { align: 'center' } })
    expect(parsed.blocks[0]!.runs).toEqual([{ text: 'Title' }])

    const p = parsed.blocks[1]!
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      { text: 'Hello ' },
      { text: 'bold', bold: true, color: '112233', sizeHalfPoints: 28, font: 'Georgia' },
      { text: ' world' },
    ])
  })

  it('preserves a nested bullet list as listItem blocks with increasing ilvl', async () => {
    const blocks: OdtSaveBlock[] = [
      { kind: 'text', block: { type: 'listItem', list: { kind: 'bullet', numId: ODT_BULLET_NUM_ID, ilvl: 0 }, runs: [{ text: 'One' }] } },
      { kind: 'text', block: { type: 'listItem', list: { kind: 'bullet', numId: ODT_BULLET_NUM_ID, ilvl: 1 }, runs: [{ text: 'One-a' }] } },
      { kind: 'text', block: { type: 'listItem', list: { kind: 'bullet', numId: ODT_BULLET_NUM_ID, ilvl: 0 }, runs: [{ text: 'Two' }] } },
      { kind: 'text', block: { type: 'paragraph', runs: [{ text: 'After the list' }] } },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    expect(parsed.blocks.map((b) => [b.type, b.list?.ilvl, b.runs?.[0]?.text])).toEqual([
      ['listItem', 0, 'One'],
      ['listItem', 1, 'One-a'],
      ['listItem', 0, 'Two'],
      ['paragraph', undefined, 'After the list'],
    ])
  })

  it('preserves an ordered list distinctly from a bullet list (not silently downgraded on save)', async () => {
    const blocks: OdtSaveBlock[] = [
      { kind: 'text', block: { type: 'listItem', list: { kind: 'ordered', numId: ODT_ORDERED_NUM_ID, ilvl: 0 }, runs: [{ text: 'First' }] } },
      { kind: 'text', block: { type: 'listItem', list: { kind: 'ordered', numId: ODT_ORDERED_NUM_ID, ilvl: 0 }, runs: [{ text: 'Second' }] } },
      { kind: 'text', block: { type: 'listItem', list: { kind: 'bullet', numId: ODT_BULLET_NUM_ID, ilvl: 1 }, runs: [{ text: 'Nested bullet under an ordered item' }] } },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    expect(parsed.blocks.map((b) => [b.list?.kind, b.list?.ilvl, b.runs?.[0]?.text])).toEqual([
      ['ordered', 0, 'First'],
      ['ordered', 0, 'Second'],
      ['bullet', 1, 'Nested bullet under an ordered item'],
    ])
  })

  it('preserves a table\'s cell text and a column span', async () => {
    const blocks: OdtSaveBlock[] = [
      {
        kind: 'table',
        model: {
          rows: [
            [{ paras: ['Metric'] }, { paras: ['Value'] }],
            [{ paras: ['Revenue'], colSpan: 2 }],
          ],
        },
      },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    expect(parsed.blocks).toHaveLength(1)
    const table = parsed.blocks[0]!.table!
    expect(table.rows[0]!.map((c) => c.paras)).toEqual([['Metric'], ['Value']])
    expect(table.rows[1]![0]).toMatchObject({ paras: ['Revenue'], colSpan: 2 })
  })

  it('re-embeds a standalone image block and resolves it back to the same bytes as a data URL', async () => {
    const blocks: OdtSaveBlock[] = [
      { kind: 'image', dataUrl: TINY_PNG, widthPx: 40, heightPx: 40 },
      { kind: 'text', block: { type: 'paragraph', runs: [{ text: 'caption' }] } },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    expect(parsed.blocks).toHaveLength(2)
    expect(parsed.blocks[0]).toMatchObject({ type: 'image', imageDataUrl: TINY_PNG })
    expect(parsed.blocks[1]!.runs).toEqual([{ text: 'caption' }])
  })

  it('re-embeds an inline (as-char) image mixed with text in one paragraph', async () => {
    // Known limitation (documented on collectRuns in parse.ts): mixing inline
    // text spans with an inline image at the same level doesn't guarantee
    // their relative order survives the round trip — assert presence and each
    // run's own content, not their exact sequence.
    const blocks: OdtSaveBlock[] = [
      {
        kind: 'text',
        block: {
          type: 'paragraph',
          runs: [{ text: 'before ' }, { text: '', image: { dataUrl: TINY_PNG, xml: '', widthPx: 20, heightPx: 20 } }, { text: ' after' }],
        },
      },
    ]
    const bytes = await saveOdt(blocks)
    const parsed = await parseOdt(bytes)

    const runs = parsed.blocks[0]!.runs!
    expect(runs).toHaveLength(3)
    expect(runs.some((r) => r.text === 'before ')).toBe(true)
    expect(runs.some((r) => r.text === ' after')).toBe(true)
    expect(runs.some((r) => r.image?.dataUrl === TINY_PNG)).toBe(true)
  })

  it('produces a ParsedDocFull-shaped stub with sensible blank-document defaults', async () => {
    const bytes = await saveOdt([{ kind: 'text', block: { type: 'paragraph', runs: [{ text: 'x' }] } }])
    const parsed = await parseOdt(bytes)
    expect(parsed.comments).toEqual([])
    expect(parsed.footnotes).toEqual([])
    expect(parsed.protection).toBeNull()
    expect(parsed.numbering.size).toBeGreaterThan(0)
    expect(parsed.internal.originalBytes).toBeInstanceOf(Uint8Array)
  })
})
