import { describe, expect, it } from 'vitest'
import { parseOdt, saveOdt } from '@genoffice/odt-engine'
import { pmDocToOdtSaveBlocks } from '../src/renderer/save-odt'
import { blocksToPmDoc, type PmNode } from '../src/renderer/editor/convert'

// A 1x1 red PNG, matching the fixture convention used in odt-engine's own tests.
const TINY_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

describe('pmDocToOdtSaveBlocks', () => {
  it('turns docHeading/docListItem/docParagraph/docTable into the matching OdtSaveBlock kinds', () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        { type: 'docHeading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Title' }] },
        {
          type: 'docListItem',
          attrs: { kind: 'bullet', numId: '1', ilvl: 0 },
          content: [{ type: 'text', text: 'Item one' }],
        },
        { type: 'docParagraph', content: [{ type: 'text', text: 'Body text' }] },
        {
          type: 'docTable',
          attrs: {},
          content: [
            {
              type: 'docTableRow',
              attrs: {},
              content: [
                {
                  type: 'docTableCell',
                  attrs: { colspan: 1, rowspan: 1 },
                  content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Cell' }] }],
                },
              ],
            },
          ],
        },
      ],
    }

    const blocks = pmDocToOdtSaveBlocks(doc)
    expect(blocks.map((b) => b.kind)).toEqual(['text', 'text', 'text', 'table'])
    expect(blocks[0]).toMatchObject({ kind: 'text', block: { type: 'heading', level: 1 } })
    expect(blocks[1]).toMatchObject({ kind: 'text', block: { type: 'listItem' } })
    expect(blocks[2]).toMatchObject({ kind: 'text', block: { type: 'paragraph' } })
    expect(blocks[3]).toMatchObject({ kind: 'table', model: { rows: [[{ paras: ['Cell'] }]] } })
  })

  it('reads a docProtected standalone image node into an image block', () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: {
            blockType: 'image',
            imageDataUrl: TINY_PNG,
            imageWidthPx: 40,
            imageHeightPx: 40,
            imageAlign: 'center',
          },
        },
      ],
    }
    const blocks = pmDocToOdtSaveBlocks(doc)
    expect(blocks).toEqual([
      { kind: 'image', dataUrl: TINY_PNG, widthPx: 40, heightPx: 40, align: 'center' },
    ])
  })

  it('drops a docProtected node with no resolvable image (e.g. a chart or OLE object) rather than guessing', () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        { type: 'docProtected', attrs: { blockType: 'chart', chartDisplay: { title: 'Q3' } } },
        { type: 'docParagraph', content: [{ type: 'text', text: 'after' }] },
      ],
    }
    const blocks = pmDocToOdtSaveBlocks(doc)
    expect(blocks).toHaveLength(1)
    expect(blocks[0]).toMatchObject({ kind: 'text', block: { type: 'paragraph' } })
  })

  it('round-trips a full doc through saveOdt/parseOdt/blocksToPmDoc unchanged in substance', async () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        { type: 'docHeading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Report' }] },
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'bold', marks: [{ type: 'bold' }] },
            { type: 'text', text: ' plain' },
          ],
        },
        {
          type: 'docListItem',
          attrs: { kind: 'ordered', numId: '2', ilvl: 0 },
          content: [{ type: 'text', text: 'Step one' }],
        },
      ],
    }

    const bytes = await saveOdt(pmDocToOdtSaveBlocks(doc))
    const parsed = await parseOdt(bytes)
    const rebuilt = blocksToPmDoc(parsed.blocks)

    const headingNode = rebuilt.content?.[0]
    expect(headingNode?.type).toBe('docHeading')
    expect(headingNode?.attrs?.level).toBe(2)

    const paraNode = rebuilt.content?.[1]
    expect(paraNode?.type).toBe('docParagraph')
    const boldRun = paraNode?.content?.find((n) => n.text === 'bold')
    expect(boldRun?.marks).toEqual([{ type: 'bold' }])

    const listNode = rebuilt.content?.[2]
    expect(listNode?.type).toBe('docListItem')
    expect(listNode?.attrs?.kind).toBe('ordered')
  })
})
