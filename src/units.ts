/**
 * Public unit conversion helpers and standard PowerPoint slide-layout constants.
 */

export const EMU_PER_INCH = 914400
export const EMU_PER_POINT = 12700
export const POINTS_PER_INCH = 72

export type StandardLayoutName = 'LAYOUT_4x3' | 'LAYOUT_16x9' | 'LAYOUT_16x10' | 'LAYOUT_WIDE'

export interface StandardLayout {
	/** PptxGenJS layout key used with `pptx.layout`. */
	readonly layout: StandardLayoutName
	/** PresentationML slide-size preset name, or `custom` for PowerPoint widescreen. */
	readonly name: string
	/** Slide width in inches. */
	readonly widthIn: number
	/** Slide height in inches. */
	readonly heightIn: number
	/** Slide width in English Metric Units. */
	readonly widthEmu: number
	/** Slide height in English Metric Units. */
	readonly heightEmu: number
}

export function inchesToEmu(inches: number): number {
	assertFiniteNumber(inches, 'inches')
	return Math.round(inches * EMU_PER_INCH)
}

export function pointsToEmu(points: number): number {
	assertFiniteNumber(points, 'points')
	return Math.round(points * EMU_PER_POINT)
}

export function pixelsToEmu(pixels: number, dpi: number): number {
	assertFiniteNumber(pixels, 'pixels')
	assertPositiveFiniteNumber(dpi, 'dpi')
	return inchesToEmu(pixels / dpi)
}

export function emuToInches(emu: number): number {
	assertFiniteNumber(emu, 'emu')
	return emu / EMU_PER_INCH
}

export function emuToPoints(emu: number): number {
	assertFiniteNumber(emu, 'emu')
	return emu / EMU_PER_POINT
}

export function emuToPixels(emu: number, dpi: number): number {
	assertFiniteNumber(emu, 'emu')
	assertPositiveFiniteNumber(dpi, 'dpi')
	return Math.round(emuToInches(emu) * dpi)
}

function standardLayout(layout: StandardLayoutName, name: string, widthIn: number, heightIn: number): StandardLayout {
	return Object.freeze({
		layout,
		name,
		widthIn,
		heightIn,
		widthEmu: inchesToEmu(widthIn),
		heightEmu: inchesToEmu(heightIn),
	})
}

export const STANDARD_LAYOUTS: Readonly<Record<StandardLayoutName, StandardLayout>> = Object.freeze({
	LAYOUT_4x3: standardLayout('LAYOUT_4x3', 'screen4x3', 10, 7.5),
	LAYOUT_16x9: standardLayout('LAYOUT_16x9', 'screen16x9', 10, 5.625),
	LAYOUT_16x10: standardLayout('LAYOUT_16x10', 'screen16x10', 10, 6.25),
	LAYOUT_WIDE: standardLayout('LAYOUT_WIDE', 'custom', 40 / 3, 7.5),
})

function assertFiniteNumber(value: number, name: string): void {
	if (!Number.isFinite(value)) throw new Error(`${name} must be a finite number`)
}

function assertPositiveFiniteNumber(value: number, name: string): void {
	assertFiniteNumber(value, name)
	if (value <= 0) throw new Error(`${name} must be greater than 0`)
}
