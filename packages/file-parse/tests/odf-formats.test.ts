import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { buildOdpFixture, buildOdsFixture, buildOdtFixture, writeFixture } from './helpers/fixtures'

describe('parseFileToText: odt', () => {
  it('extracts headings, paragraphs and tables', async () => {
    const path = writeFixture('report.odt', await buildOdtFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('Annual Report')
    expect(result.text).toContain('First paragraph hello odt')
    expect(result.text).toContain('Metric | Value')
    expect(result.text).toContain('Revenue | 100')
  })
})

describe('parseFileToText: odp', () => {
  it('extracts one section per slide in document order', async () => {
    const path = writeFixture('deck.odp', await buildOdpFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('## Slide 1\nProduct Intro\nFirst slide subtitle')
    // a custom-shape's own direct text:p (no draw:text-box wrapper)
    expect(result.text).toContain('## Slide 2\nMarket Analysis')
  })

  it('keeps text:line-break as a line break', async () => {
    const path = writeFixture('deck.odp', await buildOdpFixture())
    const result = await parseFileToText(path)
    expect(result.text).toContain('## Slide 3\nBefore\n\nAfter')
  })

  it('fails gracefully on a corrupt file', async () => {
    const path = writeFixture('broken.odp', Buffer.from('not a zip'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})

describe('parseFileToText: ods', () => {
  it('extracts sheet name and rows, with a repeated empty trailing cell', async () => {
    const path = writeFixture('table.ods', await buildOdsFixture())
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('# Grades')
    expect(result.text).toContain('Name | Scores')
    expect(result.text).toContain('Alice | 95 | ')
  })

  it('fails gracefully on a corrupt file', async () => {
    const path = writeFixture('broken.ods', Buffer.from('not a zip'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})
