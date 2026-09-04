/**
 * ODF <-> pptx-engine unit conversion. pptx-engine's model is EMU-native
 * (1 inch = 914400 EMU) with rotation in 1/60000 degree, clockwise (OOXML's
 * native units, per its own doc comment) — ODF uses CSS-style lengths (cm/mm/
 * in/pt/pc/px) and a different rotation convention entirely (see below).
 */

const EMU_PER_INCH = 914400
const EMU_PER_CM = EMU_PER_INCH / 2.54
const EMU_PER_MM = EMU_PER_CM / 10
const EMU_PER_PT = EMU_PER_INCH / 72
const EMU_PER_PC = EMU_PER_PT * 12
const EMU_PER_PX_96 = EMU_PER_INCH / 96

const UNIT_TO_EMU: Record<string, number> = {
  cm: EMU_PER_CM,
  mm: EMU_PER_MM,
  in: EMU_PER_INCH,
  pt: EMU_PER_PT,
  pc: EMU_PER_PC,
  px: EMU_PER_PX_96,
}

/** ODF length ("10.5cm", "2in", "24pt") -> EMU, rounded. Missing/unit-less values are read as cm (every real producer writes an explicit unit; this is only a defensive default). */
export function parseOdfLength(value: string | undefined | null): number {
  if (!value) return 0
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(value.trim())
  if (!m) return 0
  const num = Number.parseFloat(m[1]!)
  if (!Number.isFinite(num)) return 0
  const unit = m[2] ?? 'cm'
  return Math.round(num * (UNIT_TO_EMU[unit] ?? UNIT_TO_EMU.cm!))
}

/** EMU -> an ODF length string in cm (the LibreOffice/OpenOffice convention). */
export function emuToOdfLength(emu: number): string {
  return `${trimNumber(emu / EMU_PER_CM)}cm`
}

/** ODF font sizes are always pt ("24pt"), like OOXML's own half-point-doubled sizes. */
export function parseOdfPt(value: string | undefined | null): number | undefined {
  if (!value) return undefined
  const m = /^(-?[\d.]+)\s*pt$/.exec(value.trim())
  if (!m) return undefined
  const num = Number.parseFloat(m[1]!)
  return Number.isFinite(num) ? num : undefined
}

export function ptToOdfLength(pt: number): string {
  return `${trimNumber(pt)}pt`
}

/** Round to 4 decimal places and drop a trailing ".0000"/trailing zeros — keeps generated XML compact and diffable. */
function trimNumber(n: number): string {
  const rounded = Math.round(n * 10000) / 10000
  return rounded.toString()
}

/**
 * ODF draw:transform's rotate() argument: unitless values are RADIANS in every
 * real producer/consumer (LibreOffice, OpenOffice, odfpy), despite the ODF 1.2
 * spec text describing the shared <angle> datatype as degrees-by-default — a
 * known spec/implementation split (AOO bug 123879, "dr3d rotation angle is
 * degree in spec ODF1.2 but radian in AOO"). Positive = counter-clockwise on
 * screen (LibreOffice's documented convention) — the opposite sense from
 * OOXML's `rot`, where positive is clockwise. `rot` is 1/60000 degree,
 * normalized to [0, 21600000) by pptx-engine's own convention.
 *
 * NOTE: this direction/composition convention could not be checked against a
 * real LibreOffice-authored file in this environment (no GUI/rendering
 * available) — verified only by cross-referencing public bug reports and
 * library source examples, plus round-tripping generate->parse in this
 * package's own tests. Treat rotated-shape fidelity as the least-verified
 * part of this engine.
 */
export function ooxmlRotToOdfRadians(rot: number): number {
  const deg = -(rot / 60000)
  return (deg * Math.PI) / 180
}

export function odfRadiansToOoxmlRot(rad: number): number {
  const deg = (-rad * 180) / Math.PI
  const norm = ((deg % 360) + 360) % 360
  return Math.round(norm * 60000)
}

export const DEFAULT_SLIDE_SIZE = { cx: 9144000, cy: 6858000 }
