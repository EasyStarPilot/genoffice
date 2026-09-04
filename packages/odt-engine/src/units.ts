/** ODF <-> docx-engine unit conversion. docx-engine is twips-native (1/20 pt = 1/1440 inch) with font sizes in half-points — ODF uses CSS-style lengths (cm/mm/in/pt). */

const TWIPS_PER_INCH = 1440
const TWIPS_PER_CM = TWIPS_PER_INCH / 2.54
const TWIPS_PER_MM = TWIPS_PER_CM / 10
const TWIPS_PER_PT = 20

const UNIT_TO_TWIPS: Record<string, number> = {
  cm: TWIPS_PER_CM,
  mm: TWIPS_PER_MM,
  in: TWIPS_PER_INCH,
  pt: TWIPS_PER_PT,
}

/** ODF length ("2.5cm", "1in") -> twips, rounded. Missing/unit-less values default to cm (real files always carry a unit). */
export function parseOdfLengthTwips(value: string | undefined | null): number {
  if (!value) return 0
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt)?$/.exec(value.trim())
  if (!m) return 0
  const num = Number.parseFloat(m[1]!)
  if (!Number.isFinite(num)) return 0
  const unit = m[2] ?? 'cm'
  return Math.round(num * (UNIT_TO_TWIPS[unit] ?? UNIT_TO_TWIPS.cm!))
}

export function twipsToOdfLength(twips: number): string {
  return `${trimNumber(twips / TWIPS_PER_CM)}cm`
}

/** ODF font sizes are always pt ("24pt") -> half-points (docx-engine's Run.sizeHalfPoints). */
export function parseOdfHalfPoints(value: string | undefined | null): number | undefined {
  if (!value) return undefined
  const m = /^(-?[\d.]+)\s*pt$/.exec(value.trim())
  if (!m) return undefined
  const num = Number.parseFloat(m[1]!)
  return Number.isFinite(num) ? Math.round(num * 2) : undefined
}

export function halfPointsToOdfPt(halfPoints: number): string {
  return `${trimNumber(halfPoints / 2)}pt`
}

function trimNumber(n: number): string {
  return (Math.round(n * 10000) / 10000).toString()
}
