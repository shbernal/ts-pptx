/**
 * ts-pptx: DrawingML value clamps
 *
 * Clamp font/character/line spacing values (points) into their ST_* schema
 * ranges and convert to the hundredths-of-a-point units the attributes expect.
 * Out-of-range values make PowerPoint report the package as needing repair.
 */

import { EMU_PER_POINT, HUNDREDTHS_PER_POINT, ptToHundredths } from '../../units.js'
import { valToPts } from '../../units-internal.js'
import { warnOnce } from '../../diagnostics.js'

/**
 * Clamp a font size (points) into ST_TextFontSize (1-4000pt) and return it in
 * hundredths of a point for the `sz` attribute. Out-of-range sizes make
 * PowerPoint report the package as needing repair (e.g. `sz` > 400000 or < 100).
 */
export function clampFontSizeSz(fontSizePts: number): number {
	const raw = ptToHundredths(fontSizePts)
	const clamped = Math.min(400000, Math.max(100, raw))
	if (clamped !== raw)
		warnOnce(
			'font/size-out-of-range',
			`fontSize ${fontSizePts} is outside the valid range 1-4000pt; using ${clamped / HUNDREDTHS_PER_POINT}.`
		)
	return clamped
}

/** Clamp character spacing (points) into ST_TextPoint (-4000..4000pt); returns hundredths for the `spc` attribute. */
export function clampCharSpacingSpc(charSpacingPts: number): number {
	const raw = ptToHundredths(charSpacingPts)
	const clamped = Math.min(400000, Math.max(-400000, raw))
	if (clamped !== raw)
		warnOnce(
			'text/char-spacing-out-of-range',
			`charSpacing ${charSpacingPts} is outside the valid range -4000..4000pt; using ${clamped / HUNDREDTHS_PER_POINT}.`
		)
	return clamped
}

/**
 * Clamp a paragraph's left margin (points) into ST_TextMargin (0..4032pt, i.e. 0..51206400 EMU);
 * returns EMU for `a:pPr/@marL`. The type is unsigned — a negative margin is not a narrower one,
 * it is a value PowerPoint reports as needing repair.
 */
export function clampParaMarginEmu(marginPts: number): number {
	const raw = valToPts(marginPts)
	const clamped = Math.min(51206400, Math.max(0, raw))
	if (clamped !== raw)
		warnOnce(
			'text/paragraph-margin-out-of-range',
			`paraMarginLeft ${marginPts} is outside the valid range 0-4032pt; using ${clamped / EMU_PER_POINT}.`
		)
	return clamped
}

/**
 * Clamp a first-line indent (points) into ST_TextIndent (-4032..4032pt); returns EMU for
 * `a:pPr/@indent`. Unlike the margin this one is signed: a negative indent is the hanging
 * indent every bulleted paragraph uses.
 */
export function clampParaIndentEmu(indentPts: number): number {
	const raw = valToPts(indentPts)
	const clamped = Math.min(51206400, Math.max(-51206400, raw))
	if (clamped !== raw)
		warnOnce(
			'text/paragraph-indent-out-of-range',
			`paraIndent ${indentPts} is outside the valid range -4032..4032pt; using ${clamped / EMU_PER_POINT}.`
		)
	return clamped
}

/** Clamp line spacing (points) into ST_TextSpacingPoint (0..1584pt); returns hundredths for `<a:spcPts val>`. */
export function clampLineSpacingPts(lineSpacingPts: number): number {
	const raw = ptToHundredths(lineSpacingPts)
	const clamped = Math.min(158400, Math.max(0, raw))
	if (clamped !== raw)
		warnOnce(
			'text/line-spacing-out-of-range',
			`lineSpacing ${lineSpacingPts} is outside the valid range 0-1584pt; using ${clamped / HUNDREDTHS_PER_POINT}.`
		)
	return clamped
}
