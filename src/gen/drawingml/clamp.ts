/**
 * PptxGenJS: DrawingML value clamps
 *
 * Clamp font/character/line spacing values (points) into their ST_* schema
 * ranges and convert to the hundredths-of-a-point units the attributes expect.
 * Out-of-range values make PowerPoint report the package as needing repair.
 */

import { HUNDREDTHS_PER_POINT, ptToHundredths } from '../../units.js'
import { warnOnce } from '../../log.js'

/**
 * Clamp a font size (points) into ST_TextFontSize (1-4000pt) and return it in
 * hundredths of a point for the `sz` attribute. Out-of-range sizes make
 * PowerPoint report the package as needing repair (e.g. `sz` > 400000 or < 100).
 */
export function clampFontSizeSz(fontSizePts: number): number {
	const raw = ptToHundredths(fontSizePts)
	const clamped = Math.min(400000, Math.max(100, raw))
	if (clamped !== raw)
		warnOnce(`fontSize ${fontSizePts} is outside the valid range 1-4000pt; using ${clamped / HUNDREDTHS_PER_POINT}.`)
	return clamped
}

/** Clamp character spacing (points) into ST_TextPoint (-4000..4000pt); returns hundredths for the `spc` attribute. */
export function clampCharSpacingSpc(charSpacingPts: number): number {
	const raw = ptToHundredths(charSpacingPts)
	const clamped = Math.min(400000, Math.max(-400000, raw))
	if (clamped !== raw)
		warnOnce(
			`charSpacing ${charSpacingPts} is outside the valid range -4000..4000pt; using ${clamped / HUNDREDTHS_PER_POINT}.`
		)
	return clamped
}

/** Clamp line spacing (points) into ST_TextSpacingPoint (0..1584pt); returns hundredths for `<a:spcPts val>`. */
export function clampLineSpacingPts(lineSpacingPts: number): number {
	const raw = ptToHundredths(lineSpacingPts)
	const clamped = Math.min(158400, Math.max(0, raw))
	if (clamped !== raw)
		warnOnce(
			`lineSpacing ${lineSpacingPts} is outside the valid range 0-1584pt; using ${clamped / HUNDREDTHS_PER_POINT}.`
		)
	return clamped
}
