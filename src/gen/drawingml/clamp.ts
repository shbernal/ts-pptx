/**
 * ts-pptx: DrawingML value clamps
 *
 * Clamp font/character/line spacing values into their ST_* schema ranges and convert to the
 * units the attributes expect — hundredths of a point for most of them, thousandths of a
 * percent for the one measure the caller states as a multiple rather than a size.
 * Out-of-range values make PowerPoint report the package as needing repair.
 */

import { EMU_PER_INCH, EMU_PER_POINT, HUNDREDTHS_PER_POINT, PERCENT_SCALE, ptToHundredths } from '../../units.js'
import { inch2Emu, ptsToEmuLenient } from '../../units-internal.js'
import { warnOnce } from '../../diagnostics.js'
import type { DiagnosticCode } from '../../codes.js'

/**
 * One clamped text measure: how to convert the caller's value, what range the schema type
 * allows in the converted unit, and how to say so if the value falls outside it.
 */
interface ClampSpec {
	/** Diagnostic code raised when the value is clamped. */
	code: DiagnosticCode
	/** Option name as the caller spells it, opening the warning. */
	label: string
	/** The valid range in the caller's own unit, as the warning states it. */
	range: string
	/** The caller's unit to the attribute's own unit. */
	convert: (value: number) => number
	/** Converted units per caller unit — divides the clamped value back for the warning. */
	perUnit: number
	min: number
	max: number
}

/**
 * Clamp a caller's value into its schema range and warn once if it did not already fit.
 *
 * The six measures below differ only in the fields of {@link ClampSpec}; the arithmetic and
 * the shape of the warning are the same for all of them. `warnOnce` keys its dedupe set on the
 * code plus the message, so the wording is part of the contract, not decoration.
 */
function clampWithWarn(value: number, spec: ClampSpec, label = spec.label): number {
	const raw = spec.convert(value)
	const clamped = Math.min(spec.max, Math.max(spec.min, raw))
	if (clamped !== raw)
		warnOnce(spec.code, `${label} ${value} is outside the valid range ${spec.range}; using ${clamped / spec.perUnit}.`)
	return clamped
}

const FONT_SIZE: ClampSpec = {
	code: 'font/size-out-of-range',
	label: 'fontSize',
	range: '1-4000pt',
	convert: ptToHundredths,
	perUnit: HUNDREDTHS_PER_POINT,
	min: 100,
	max: 400000,
}

const CHAR_SPACING: ClampSpec = {
	code: 'text/char-spacing-out-of-range',
	label: 'charSpacing',
	range: '-4000..4000pt',
	convert: ptToHundredths,
	perUnit: HUNDREDTHS_PER_POINT,
	min: -400000,
	max: 400000,
}

const PARA_MARGIN: ClampSpec = {
	code: 'text/paragraph-margin-out-of-range',
	label: 'paraMarginLeft',
	range: '0-4032pt',
	convert: ptsToEmuLenient,
	perUnit: EMU_PER_POINT,
	min: 0,
	max: 51206400,
}

const PARA_INDENT: ClampSpec = {
	code: 'text/paragraph-indent-out-of-range',
	label: 'paraIndent',
	range: '-4032..4032pt',
	convert: ptsToEmuLenient,
	perUnit: EMU_PER_POINT,
	min: -51206400,
	max: 51206400,
}

/** {@link PARA_MARGIN}, for the callers that state the measure in inches. 4032pt is 56in. */
const PARA_MARGIN_INCHES: ClampSpec = {
	...PARA_MARGIN,
	range: '0-56in',
	convert: inch2Emu,
	perUnit: EMU_PER_INCH,
}

/** {@link PARA_INDENT}, for the callers that state the measure in inches. */
const PARA_INDENT_INCHES: ClampSpec = {
	...PARA_INDENT,
	range: '-56..56in',
	convert: inch2Emu,
	perUnit: EMU_PER_INCH,
}

const LINE_SPACING: ClampSpec = {
	code: 'text/line-spacing-out-of-range',
	label: 'lineSpacing',
	range: '0-1584pt',
	convert: ptToHundredths,
	perUnit: HUNDREDTHS_PER_POINT,
	min: 0,
	max: 158400,
}

const LINE_SPACING_MULTIPLE: ClampSpec = {
	code: 'text/line-spacing-out-of-range',
	label: 'lineSpacingMultiple',
	range: '0-132',
	convert: (multiple) => Math.round(multiple * PERCENT_SCALE),
	perUnit: PERCENT_SCALE,
	min: 0,
	max: 13200000,
}

/**
 * Clamp a font size (points) into ST_TextFontSize (1-4000pt) and return it in
 * hundredths of a point for the `sz` attribute. Out-of-range sizes make
 * PowerPoint report the package as needing repair (e.g. `sz` > 400000 or < 100).
 *
 * Every `sz` in the package goes through here, and the nine options that reach it are spelled
 * nine different ways -- `fontSize` on a run, `legendFontSize` and six siblings on a chart.
 * `label` is what the warning quotes back, so a caller who set `catAxisLabelFontSize` is not
 * told about `fontSize`.
 * @param fontSizePts - the caller's size in points
 * @param label - option name as the caller spells it
 */
export function clampFontSizeSz(fontSizePts: number, label?: string): number {
	return clampWithWarn(fontSizePts, FONT_SIZE, label)
}

/** Clamp character spacing (points) into ST_TextPoint (-4000..4000pt); returns hundredths for the `spc` attribute. */
export function clampCharSpacingSpc(charSpacingPts: number): number {
	return clampWithWarn(charSpacingPts, CHAR_SPACING)
}

/**
 * Clamp a paragraph's left margin (points) into ST_TextMargin (0..4032pt, i.e. 0..51206400 EMU);
 * returns EMU for `a:pPr/@marL`. The type is unsigned — a negative margin is not a narrower one,
 * it is a value PowerPoint reports as needing repair.
 */
export function clampParaMarginEmu(marginPts: number): number {
	return clampWithWarn(marginPts, PARA_MARGIN)
}

/**
 * Clamp a first-line indent (points) into ST_TextIndent (-4032..4032pt); returns EMU for
 * `a:pPr/@indent`. Unlike the margin this one is signed: a negative indent is the hanging
 * indent every bulleted paragraph uses.
 */
export function clampParaIndentEmu(indentPts: number): number {
	return clampWithWarn(indentPts, PARA_INDENT)
}

/**
 * Clamp a left margin the caller stated in INCHES into ST_TextMargin, returning EMU for
 * `a:pPr/@marL`.
 *
 * A master's text styles state their measures in inches where a run's paragraph properties
 * state them in points. The schema bound is the same one either way, so the two spellings
 * share {@link PARA_MARGIN}'s range and differ only in the unit the warning quotes back.
 * @param marginInches - the caller's margin in inches
 * @param label - option name as the caller spells it
 */
export function clampParaMarginInchesEmu(marginInches: number, label: string): number {
	return clampWithWarn(marginInches, PARA_MARGIN_INCHES, label)
}

/**
 * Clamp a first-line indent the caller stated in INCHES into ST_TextIndent, returning EMU for
 * `a:pPr/@indent`. The inch-stated half of {@link clampParaIndentEmu}; see
 * {@link clampParaMarginInchesEmu} for why there are two.
 * @param indentInches - the caller's indent in inches
 * @param label - option name as the caller spells it
 */
export function clampParaIndentInchesEmu(indentInches: number, label: string): number {
	return clampWithWarn(indentInches, PARA_INDENT_INCHES, label)
}

/** Clamp line spacing (points) into ST_TextSpacingPoint (0..1584pt); returns hundredths for `<a:spcPts val>`. */
export function clampLineSpacingPts(lineSpacingPts: number): number {
	return clampWithWarn(lineSpacingPts, LINE_SPACING)
}

/**
 * Clamp a line-spacing multiple into ST_TextSpacingPercentOrPercentString (0..13200000, i.e. a
 * multiple of 0..132); returns thousandths of a percent for `<a:spcPct val>`. The bound is the
 * schema's, not the `0.0-9.99` the option documents: a multiple of 12 is unusual, not invalid,
 * and only the values PowerPoint reports as needing repair are worth moving.
 */
export function clampLineSpacingMultiplePct(multiple: number): number {
	return clampWithWarn(multiple, LINE_SPACING_MULTIPLE)
}
