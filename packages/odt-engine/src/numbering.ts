import type { NumberingDef, NumberingLevel } from '@genoffice/docx-engine'

/** Shared numIds for every bullet/ordered list in an odt document (matching docx-engine's own blank-document convention — BLANK_BULLET_NUM_ID/BLANK_ORDERED_NUM_ID — so odt lists render with the same defaults a brand-new docx list would use). Real per-list-style formatting (mixed bullet glyphs, per-level number formats) is not modeled in v1: every bullet list looks the same, every ordered list looks the same. */
export const ODT_BULLET_NUM_ID = '1'
export const ODT_ORDERED_NUM_ID = '2'

function levels(build: (ilvl: number) => NumberingLevel): Record<number, NumberingLevel> {
  const out: Record<number, NumberingLevel> = {}
  for (let ilvl = 0; ilvl < 9; ilvl++) out[ilvl] = build(ilvl)
  return out
}

export function odtNumberingDefs(): Map<string, NumberingDef> {
  const map = new Map<string, NumberingDef>()
  map.set(ODT_BULLET_NUM_ID, {
    numId: ODT_BULLET_NUM_ID,
    abstractNumId: '0',
    levels: levels((ilvl) => ({
      numFmt: 'bullet',
      lvlText: '•',
      start: 1,
      indentLeft: 720 * (ilvl + 1),
      hanging: 360,
    })),
    startOverrides: {},
  })
  map.set(ODT_ORDERED_NUM_ID, {
    numId: ODT_ORDERED_NUM_ID,
    abstractNumId: '1',
    levels: levels((ilvl) => ({
      numFmt: 'decimal',
      lvlText: '%1.',
      start: 1,
      indentLeft: 720 * (ilvl + 1),
      hanging: 360,
    })),
    startOverrides: {},
  })
  return map
}
