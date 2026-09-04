import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { PackageArchive } from '@genoffice/pptx-engine'
import type { OpenedPptx, Slide } from '@genoffice/pptx-engine'
import { parseOdp } from '../src/parse'
import { saveOdp } from '../src/generate'

async function emptyArchive(): Promise<PackageArchive> {
  const bytes = await new JSZip().generateAsync({ type: 'uint8array' })
  return PackageArchive.open(bytes)
}

describe('odp round-trip: model -> saveOdp -> parseOdp', () => {
  it('preserves an unrotated text box exactly', async () => {
    const slides: Slide[] = [
      {
        path: 'p1',
        originalXml: '',
        bodyPrefix: '',
        bodySuffix: '',
        elements: [
          {
            id: 'e1',
            type: 'text',
            anchor: { spIndex: 0, originalXml: '', range: [0, 0] },
            transform: {
              offset: { x: 914400, y: 457200, cx: 3657600, cy: 1828800 },
              rot: 0,
              flipH: false,
              flipV: false,
            },
            fill: { type: 'solid', color: '#FF0000' },
            text: {
              paragraphs: [
                {
                  runs: [{ text: 'Hello world', bold: true, color: '#112233', fontSize: 24 }],
                  align: 'center',
                },
              ],
            },
          },
        ],
      },
    ]
    const opened: OpenedPptx = {
      deck: { slides, size: { cx: 9144000, cy: 6858000 }, originalHash: 'x' },
      archive: await emptyArchive(),
    }
    const bytes = await saveOdp(opened)
    const reopened = await parseOdp(bytes)

    expect(reopened.deck.size).toEqual({ cx: 9144000, cy: 6858000 })
    expect(reopened.deck.slides).toHaveLength(1)
    const el = reopened.deck.slides[0]!.elements[0]!
    expect(el.type).toBe('text')
    expect(el.transform.rot).toBe(0)
    expect(el.transform.offset).toEqual({ x: 914400, y: 457200, cx: 3657600, cy: 1828800 })
    if (el.type !== 'text') throw new Error('expected text')
    expect(el.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(el.text?.paragraphs).toHaveLength(1)
    const p = el.text!.paragraphs[0]!
    expect(p.align).toBe('center')
    expect(p.runs).toEqual([{ text: 'Hello world', bold: true, color: '#112233', fontSize: 24 }])
  })

  it('preserves a rotated shape\'s center and size (the load-bearing transform math)', async () => {
    // A 2in x 1in shape centered at (4in, 3in), rotated 37 degrees clockwise (OOXML rot).
    const cx = 1828800 // 2in
    const cy = 914400 // 1in
    const centerX = 3657600 // 4in
    const centerY = 2743200 // 3in
    const rot = Math.round(37 * 60000)
    const offset = { x: centerX - cx / 2, y: centerY - cy / 2, cx, cy }
    const slides: Slide[] = [
      {
        path: 'p1',
        originalXml: '',
        bodyPrefix: '',
        bodySuffix: '',
        elements: [
          {
            id: 'e1',
            type: 'shape',
            presetGeometry: 'rect',
            anchor: { spIndex: 0, originalXml: '', range: [0, 0] },
            transform: { offset, rot, flipH: false, flipV: false },
            text: { paragraphs: [{ runs: [] }] },
          },
        ],
      },
    ]
    const opened: OpenedPptx = {
      deck: { slides, size: { cx: 9144000, cy: 6858000 }, originalHash: 'x' },
      archive: await emptyArchive(),
    }
    const bytes = await saveOdp(opened)
    const reopened = await parseOdp(bytes)
    const el = reopened.deck.slides[0]!.elements[0]!

    // cm round-trip (4 decimal places) plus float rotation trig loses a small
    // amount of precision — a few hundred EMU (a tiny fraction of a mm) is noise.
    const EPS = 500
    expect(Math.abs(el.transform.offset.x - offset.x)).toBeLessThan(EPS)
    expect(Math.abs(el.transform.offset.y - offset.y)).toBeLessThan(EPS)
    expect(Math.abs(el.transform.offset.cx - cx)).toBeLessThan(EPS)
    expect(Math.abs(el.transform.offset.cy - cy)).toBeLessThan(EPS)
    // rot: allow a fraction of a degree of drift from the cm-rounded translate
    expect(Math.abs(el.transform.rot - rot)).toBeLessThan(600) // 0.01 degree
  })

  it('re-embeds a freshly-added picture under a canonical Pictures/ path', async () => {
    const archive = await emptyArchive()
    archive.entries.set('ppt/media/image1.png', new Uint8Array([1, 2, 3, 4]))
    const slides: Slide[] = [
      {
        path: 'p1',
        originalXml: '',
        bodyPrefix: '',
        bodySuffix: '',
        elements: [
          {
            id: 'e1',
            type: 'picture',
            mediaRef: 'ppt/media/image1.png',
            anchor: { spIndex: 0, originalXml: '', range: [0, 0] },
            transform: {
              offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
              rot: 0,
              flipH: false,
              flipV: false,
            },
          },
        ],
      },
    ]
    const opened: OpenedPptx = {
      deck: { slides, size: { cx: 9144000, cy: 6858000 }, originalHash: 'x' },
      archive,
    }
    const bytes = await saveOdp(opened)
    // saveOdp also rewrites the in-memory mediaRef to the new canonical path
    expect((slides[0]!.elements[0] as { mediaRef: string }).mediaRef).toMatch(/^Pictures\/image\d+\.png$/)

    const reopened = await parseOdp(bytes)
    const el = reopened.deck.slides[0]!.elements[0]!
    expect(el.type).toBe('picture')
    if (el.type !== 'picture') throw new Error('expected picture')
    expect(el.mediaRef).toMatch(/^Pictures\/image\d+\.png$/)
    expect(reopened.archive.readBytes(el.mediaRef)).toEqual(new Uint8Array([1, 2, 3, 4]))
  })

  it('rejects a non-presentation OpenDocument file with a clear error', async () => {
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
    zip.file('content.xml', '<?xml version="1.0"?><office:document-content/>')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    await expect(parseOdp(bytes)).rejects.toThrow(/not a presentation/)
  })
})
