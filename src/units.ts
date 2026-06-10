/**
 * Public unit conversion helpers and standard PowerPoint slide-layout constants.
 */

export const EMU_PER_INCH = 914400
export const EMU_PER_POINT = 12700
export const POINTS_PER_INCH = 72

/**
 * English Metric Units — the integer unit OOXML serializes geometry in (914400 per inch).
 *
 * Branded so a value that has already been resolved to EMU cannot be silently fed back into a
 * unit converter: {@link coordToEmu} only accepts a user `Coord`, so passing an `Emu` into it is
 * a compile-time error. This replaces the old runtime "a number ≥ 100 must already be EMU"
 * magnitude guess, which silently mis-rendered values near the threshold.
 */
export type Emu = number & { readonly __unit: 'emu' }

/** A bare number larger than this (in inches) is almost certainly a mistake — likely a raw EMU
 *  value passed where inches are expected. We interpret it as inches (the documented contract) but
 *  warn, pointing at the explicit `"<n>emu"` form. ~1000in is far beyond any real slide. */
const IMPLAUSIBLE_INCHES = 1000

export type StandardLayoutName = 'LAYOUT_4x3' | 'LAYOUT_16x9' | 'LAYOUT_16x10' | 'LAYOUT_WIDE'

export interface StandardLayout {
	/** PptxGenJS layout key used with `pptx.layout`. */
	readonly layout: StandardLayoutName
	/** PresentationML slide-size preset name, or `custom` for PowerPoint widescreen. */
	readonly name: string
	/** Slide width in inches. Alias of {@link StandardLayout.widthIn} — inches is PptxGenJS's default coordinate unit, so this is the value to use for `addText`/`addShape` math. */
	readonly width: number
	/** Slide height in inches. Alias of {@link StandardLayout.heightIn}. */
	readonly height: number
	/** Slide width in inches. */
	readonly widthIn: number
	/** Slide height in inches. */
	readonly heightIn: number
	/** Slide width in English Metric Units. */
	readonly widthEmu: number
	/** Slide height in English Metric Units. */
	readonly heightEmu: number
}

export function inchesToEmu(inches: number): Emu {
	assertFiniteNumber(inches, 'inches')
	return Math.round(inches * EMU_PER_INCH) as Emu
}

export function pointsToEmu(points: number): Emu {
	assertFiniteNumber(points, 'points')
	return Math.round(points * EMU_PER_POINT) as Emu
}

export function pixelsToEmu(pixels: number, dpi: number): Emu {
	assertFiniteNumber(pixels, 'pixels')
	assertPositiveFiniteNumber(dpi, 'dpi')
	return inchesToEmu(pixels / dpi)
}

/**
 * Resolve a percentage of an axis length to EMU.
 * @param percent - percentage value (e.g. `50` for 50%)
 * @param axisEmu - the axis length in EMU (slide width for x/w, height for y/h)
 */
export function percentToEmu(percent: number, axisEmu: number): Emu {
	assertFiniteNumber(percent, 'percent')
	assertFiniteNumber(axisEmu, 'axisEmu')
	return Math.round((percent / 100) * axisEmu) as Emu
}

/**
 * The single user-coordinate → EMU boundary. Convert each user-supplied coordinate exactly once.
 *
 * Accepts (see {@link Coord}):
 * - a bare `number` → **always inches** (the documented unit); no magnitude guessing
 * - `"<n>%"` → percentage of `axisEmu`
 * - `"<n>in"` / `"<n>pt"` / `"<n>emu"` → explicit units (the escape hatch for non-inch values)
 *
 * Throws on non-finite or unparseable input rather than silently emitting a degenerate 0-size.
 * @param value - user coordinate
 * @param axisEmu - axis length in EMU, used only to resolve percentages
 */
export function coordToEmu(value: number | string, axisEmu: number): Emu {
	if (typeof value === 'number') {
		assertFiniteNumber(value, 'coordinate')
		if (Math.abs(value) > IMPLAUSIBLE_INCHES) {
			console.warn(
				`PptxGenJS: coordinate ${value} interpreted as ${value} inches. A bare number is always inches; ` +
					`if you meant EMU, pass it as a string like "${Math.round(value)}emu".`
			)
		}
		return inchesToEmu(value)
	}

	const match = /^\s*(-?\d*\.?\d+)\s*(%|in|pt|emu)\s*$/.exec(value)
	if (!match) {
		throw new Error(
			`PptxGenJS: invalid coordinate "${value}". Expected a number (inches) or a string like "50%", "5in", "72pt", or "914400emu".`
		)
	}
	const n = Number(match[1])
	switch (match[2]) {
		case '%':
			return percentToEmu(n, axisEmu)
		case 'in':
			return inchesToEmu(n)
		case 'pt':
			return pointsToEmu(n)
		default: // 'emu'
			assertFiniteNumber(n, 'coordinate')
			return Math.round(n) as Emu
	}
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
		width: widthIn,
		height: heightIn,
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
