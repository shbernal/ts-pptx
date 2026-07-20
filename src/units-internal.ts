/**
 * Generator-side unit conversion — the lenient layer over `units.ts`.
 *
 * `units.ts` holds the strict, public primitives: one branded `Emu` type and converters that
 * take a number and throw on anything they cannot represent. This module holds what the
 * generators actually need on top of that: converters that accept the loose shapes the
 * authoring API has always accepted (numeric strings, `"5in"`, `null`), that clamp to the
 * ranges PowerPoint will load without offering to repair the file, and that warn about legacy
 * inputs instead of failing.
 *
 * The split is deliberate: the leniency and the back-compat warnings are policy, not
 * arithmetic, and policy should not be part of the published API. Nothing here is exported
 * from an entrypoint.
 */

import { EMU, ONEPT } from './core-enums.js'
import { warn, warnOnce } from './log.js'
import {
	ANGLE_UNITS_PER_DEGREE,
	coordToEmu,
	FIXED_PCT_PER_PERCENT,
	inchesToEmu,
	PERCENT_SCALE,
	type Emu,
} from './units.js'
import type { Coord, PresLayout } from './core-interfaces.js'

/**
 * Resolve a user `Coord` (x/y/w/h) to EMU — the single user-coordinate → EMU boundary.
 * - bare `number` → **inches** (no magnitude guessing); `"<n>%"` → percent of the slide axis;
 *   `"<n>in"`/`"<n>pt"`/`"<n>emu"` → explicit units (see {@link Coord} / {@link coordToEmu})
 * - `null`/`undefined` → 0 (callers may omit a coordinate)
 * - throws on a non-finite number rather than silently collapsing the object to zero size
 * @param {Coord|null|undefined} size - user coordinate
 * @param {'X' | 'Y'} xyDir - axis (selects slide width vs height for percentages)
 * @param {PresLayout} layout - presentation layout (EMU dimensions)
 * @returns {Emu} resolved EMU value
 */
export function getSmartParseNumber(size: Coord | null | undefined, xyDir: 'X' | 'Y', layout: PresLayout): Emu {
	if (size === null || size === undefined) return 0 as Emu

	// GUARD: A NaN/Infinity coordinate is always a mistake (commonly arithmetic on an
	// `undefined` layout dimension). Fail loud with a targeted hint instead of the generic
	// converter message, since this is the most common way a deck collapses to zero-size.
	if (typeof size === 'number' && !isFinite(size)) {
		throw new Error(
			`Invalid ${xyDir || 'coordinate'} value: expected a finite number but received ${String(size)}. ` +
				'This usually means a layout dimension was read from a missing property (e.g. `layout.width` returning `undefined`). ' +
				'Use `slide.width`/`slide.height` or `STANDARD_LAYOUTS.<NAME>.width`/`.height` (inches).'
		)
	}

	return coordToEmu(size, xyDir === 'Y' ? layout.height : layout.width)
}

/**
 * Convert inches into EMU.
 * - accepts a number (inches) or a numeric/`"<n>in"` string
 * - no magnitude guessing: values are always treated as inches (use {@link coordToEmu} for
 *   user coordinates that may carry other units)
 * @param {number|string} inches - inches as number or string
 * @returns {Emu} EMU value
 */
export function inch2Emu(inches: number | string): Emu {
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return inchesToEmu(inches)
}

/**
 * Convert a single `margin` component (table cell/table margin, or text-box/placeholder body
 * inset) to EMU.
 *
 * Margins are INCHES, consistent with the positional API (`x`/`y`/`w`/`h`) and the value
 * PowerPoint's own dialog shows (both the table cell-margin and the text-box internal-margin
 * fields are inches). Historically the library read these as POINTS (table cells used a magnitude
 * heuristic — `>= 1` points, `< 1` inches; text-box margins were straight points), so a legitimate
 * fraction-of-an-inch value entered from the PowerPoint dialog became a tiny points value. Every
 * value is now inches. A `>= 1` value is honored as inches but warns once, because it is almost
 * certainly a legacy points value that should be divided by 72 (e.g. `10` points → `0.139` inches).
 *
 * Shared by every margin site so they stay in lockstep: the cell XML emitter and text-box/slide-
 * number insets (`gen-xml`), the autoPage row-height pass (`gen-tables`), and the measured-fit pass
 * (`measure-fit`).
 * @param {number} inches - margin component in inches
 * @returns {Emu} EMU value
 */
export function marginToEmu(inches: number): Emu {
	if (inches >= 1)
		warnOnce(
			'margins (table cell and text-box) are interpreted as inches (matching the rest of the API and the ' +
				'PowerPoint dialog); a value >= 1 is likely a legacy points value — divide by 72 to convert (e.g. 10pt => 0.139in).'
		)
	return inch2Emu(inches)
}

/**
 * Resolve a table's column widths to EMU, the single source of truth shared by the
 * table XML emitter and the measured-fit pass (so a fitted cell sees the same grid
 * the renderer draws).
 * - an explicit `colW` **array** is per-column inches (`inch2Emu`); a non-finite slot
 *   falls back to the even-distribution width
 * - otherwise the table's already-resolved width (`totalWidthEmu`, EMU) is split evenly
 *   across `colCount` columns. (A scalar `colW` never reaches here — `addTableDefinition`
 *   converts it to `w` and clears `colW`.)
 *
 * IMPORTANT: the even path divides an **EMU** width. Passing the raw inches `options.w`
 * instead (the historical bug) produced ~0-EMU columns (e.g. `w=9` → `gridCol w="3"`),
 * collapsing auto-width tables to a sliver in PowerPoint/LibreOffice.
 * @param {Coord[]|Coord|undefined} colW - explicit per-column inches, or scalar/undefined
 * @param {number} totalWidthEmu - the table's resolved width in EMU
 * @param {number} colCount - number of grid columns (counting colspans)
 * @returns {number[]} per-column widths in EMU (length `colCount`)
 */
export function resolveTableColWidthsEmu(
	colW: Coord[] | Coord | undefined,
	totalWidthEmu: number,
	colCount: number
): number[] {
	if (!(colCount > 0)) return []
	const even = totalWidthEmu > 0 ? Math.round(totalWidthEmu / colCount) : EMU
	if (Array.isArray(colW)) {
		return Array.from({ length: colCount }, (_, i) => {
			// Guard before inch2Emu: it throws on non-finite input. A missing/NaN slot
			// falls back to the even-distribution width.
			const n = colW[i]
			return typeof n === 'number' && Number.isFinite(n) ? Math.round(inch2Emu(n)) : even
		})
	}
	return new Array<number>(colCount).fill(even)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts(pt: number | string): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert a transparency percentage (0-100) into a schema-valid `<a:alpha>` value
 * (ST_PositiveFixedPercentage, 0-100000). Out-of-range transparency yields an
 * alpha that PowerPoint rejects as needing repair, so clamp into range and warn.
 */
export function transparencyToAlpha(transparency: number): number {
	const pct = Math.min(100, Math.max(0, transparency))
	if (pct !== transparency) warn(`transparency ${transparency} is outside the valid range 0-100; using ${pct}.`)
	return Math.round((100 - pct) * FIXED_PCT_PER_PERCENT)
}

/** Convert an opacity (0-1) into a schema-valid `<a:alpha>` value (0-100000); clamps + warns on out-of-range input. */
export function opacityToAlpha(opacity: number): number {
	const o = Math.min(1, Math.max(0, opacity))
	if (o !== opacity) warn(`opacity ${opacity} is outside the valid range 0-1; using ${o}.`)
	return Math.round(o * PERCENT_SCALE)
}

/**
 * Convert a line width (points) to EMU clamped into ST_LineWidth (0..20116800 EMU,
 * i.e. 0-1584pt). Out-of-range widths make PowerPoint report the package as needing
 * repair, so clamp into range and warn.
 */
export function lineWidthToEmu(widthPts: number | string): number {
	const raw = valToPts(widthPts)
	const clamped = Math.min(20116800, Math.max(0, raw))
	if (clamped !== raw) warn(`line width ${widthPts} is outside the valid range 0-1584pt; using ${clamped / ONEPT}.`)
	return clamped
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees(d: number): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * ANGLE_UNITS_PER_DEGREE)
}

/**
 * Convert a freeform arc angle (degrees) to an `<a:arcTo>` ST_AdjAngle value (60000ths).
 * Unlike a shape rotation, a sweep is not modular: a 400 degree swAng draws a different
 * arc than a 40 degree one, so the value is never wrapped into 0..360.
 * @param {number} d degrees
 * @param {'stAng' | 'swAng'} attr - attribute being emitted, for the error message
 * @returns {number} ST_AdjAngle value (60000ths of a degree)
 */
export function convertArcAngle(d: number, attr: 'stAng' | 'swAng'): number {
	if (typeof d !== 'number' || !Number.isFinite(d))
		throw new Error(`Arc ${attr} must be a finite number of degrees; received ${String(d)}.`)
	return Math.round(d * ANGLE_UNITS_PER_DEGREE)
}
