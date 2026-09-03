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

import { DEF_CELL_MARGIN_IN, DEF_SLIDE_MARGIN_IN, LINEH_MODIFIER } from './constants-internal.js'
import { warn, warnOnce } from './diagnostics.js'
import { InvalidOptionError } from './errors.js'
import {
	EMU_PER_INCH,
	EMU_PER_POINT,
	ANGLE_UNITS_PER_DEGREE,
	coordToEmu,
	FIXED_PCT_PER_PERCENT,
	inchesToEmu,
	PERCENT_SCALE,
	type Emu,
} from './units.js'
import type { DiagnosticCode, InvalidOptionErrorCode } from './codes.js'
import type { Coord, PresLayout } from './types/index.js'

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
	if (typeof size === 'number' && !Number.isFinite(size)) {
		throw new InvalidOptionError(
			'coord/non-finite',
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
 * number insets (`gen/slide/object.ts`), the autoPage row-height pass (`gen/table/autopage.ts`), and the measured-fit pass
 * (`measure-fit`).
 * @param {number} inches - margin component in inches
 * @returns {Emu} EMU value
 */
export function marginToEmu(inches: number): Emu {
	if (inches >= 1)
		warnOnce(
			'margin/legacy-points',
			'margins (table cell and text-box) are interpreted as inches (matching the rest of the API and the ' +
				'PowerPoint dialog); a value >= 1 is likely a legacy points value — divide by 72 to convert (e.g. 10pt => 0.139in).'
		)
	return inch2Emu(inches)
}

/**
 * A `[Top, Right, Bottom, Left]` margin (inches) as the four text insets in EMU, or `null` when
 * the caller stated no margin.
 *
 * A scalar broadcasts to all four sides. The index order is the whole point of naming this:
 * the arrays are CSS order while `a:bodyPr`/`a:tcPr` name their insets left-first, so every
 * site that mapped it by hand spelled the same 3/0/1/2 shuffle — the text box's `_bodyProp`,
 * the slide-number placeholder's `a:bodyPr`, and a table cell's `marL/R/T/B`.
 *
 * Each component goes through {@link marginToEmu}, so the legacy-points warning still fires
 * once for a value of an inch or more.
 * @param margin - the caller's `margin`, a scalar or a `[T, R, B, L]` array of inches
 */
export function resolveInsetsEmu(margin: number | number[] | undefined | null): {
	l: number
	t: number
	r: number
	b: number
} | null {
	if (typeof margin === 'number') {
		const all = marginToEmu(margin)
		return { l: all, t: all, r: all, b: all }
	}
	if (!Array.isArray(margin)) return null
	const at = (idx: number): number => marginToEmu(margin[idx] || 0)
	return { l: at(3), t: at(0), r: at(1), b: at(2) }
}

/**
 * One line of table text, in EMU: `fontSizePt * (LINEH_MODIFIER + lineWeight) / 100`.
 *
 * The auto-pager prices a row by counting lines and multiplying by this, and the measured-fit
 * pass estimates a cell's height with the same product. Three sites wrote the expression by
 * hand and one of them left the caller's `autoPageLineWeight` out, so the pager guarded a row
 * against a line it then measured at a different height.
 *
 * @param fontSizePt - the resolved font size in points
 * @param lineWeight - the caller's `autoPageLineWeight`, an addend on the modifier
 */
export function autoPageLineHeightEmu(fontSizePt: number, lineWeight = 0): number {
	return inch2Emu((fontSizePt * (LINEH_MODIFIER + lineWeight)) / 100)
}

/**
 * The four slide margins in inches, `[top, right, bottom, left]`, from the master's own margin
 * and the caller's `slideMargin`.
 *
 * Precedence is master, then caller, then {@link DEF_SLIDE_MARGIN_IN}. A scalar broadcasts to
 * all four sides; an array is taken as-is.
 *
 * Three sites derived this and they already disagreed. Two gated the master on
 * `typeof !== 'undefined'` and coerced with `Number.isFinite(Number(m))`; the HTML path gated
 * on truthiness and tested `Number.isFinite(m)` with no coercion. So a master with
 * `_margin: 0` took the master branch in two of them and the caller branch in the third, and
 * a master with `_margin: "0.25"` resolved in two and was ignored in the third.
 *
 * @param masterMargin - the master/layout's `_margin`, in inches
 * @param slideMargin - the caller's `slideMargin`, in inches
 */
export function resolveSlideMarginsInches(
	masterMargin: number | number[] | string | undefined | null,
	slideMargin?: number | number[] | string | null
): [number, number, number, number] {
	const broadcast = (value: number | number[] | string): [number, number, number, number] | null => {
		if (Array.isArray(value)) {
			const [t = 0, r = 0, b = 0, l = 0] = value
			return [t, r, b, l]
		}
		const n = Number(value)
		return Number.isFinite(n) ? [n, n, n, n] : null
	}
	if (masterMargin !== undefined && masterMargin !== null) {
		const resolved = broadcast(masterMargin)
		if (resolved) return resolved
	}
	if (slideMargin !== undefined && slideMargin !== null) {
		const resolved = broadcast(slideMargin)
		if (resolved) return resolved
	}
	return [...DEF_SLIDE_MARGIN_IN]
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
	const even = totalWidthEmu > 0 ? Math.round(totalWidthEmu / colCount) : EMU_PER_INCH
	if (Array.isArray(colW)) {
		return Array.from({ length: colCount }, (_, i) => {
			// Guard before inch2Emu: it throws on non-finite input. A slot that is present but
			// unusable falls back to the even-distribution width AND says so, for the reason
			// `pinnedRowHeightInches` gives about the analogous `rowH` entry: it is something the
			// caller wrote on purpose and the even split is not what they meant. An absent slot is
			// not that — a sparse array is how a caller spells "distribute this one" — so it is
			// silent, which is the same line that helper draws.
			const n = colW[i]
			if (typeof n === 'number' && Number.isFinite(n)) return Math.round(inch2Emu(n))
			if (n !== undefined && n !== null)
				warnOnce(
					'table/invalid-col-width',
					`colW entry ${String(n)} is not a number of inches; that column takes an even share of the table width instead.`
				)
			return even
		})
	}
	return new Array<number>(colCount).fill(even)
}

/**
 * The slide width a table may occupy, in EMU: the slide, less the table's own left
 * edge and the right slide margin.
 *
 * The width twin of the auto-pager's `calcSlideTabH`, and the same reading
 * `addTableDefinition` applies when a table states neither `w` nor `colW`. The pager
 * used to *add* the two margins here instead of subtracting them, which made its
 * fallback width about one inch however wide the slide was, and every column a sliver.
 *
 * @param presLayout - the presentation layout, for the slide width
 * @param xEmu - the table's resolved left edge in EMU; `0`/absent falls back to the left slide margin
 * @param marginsIn - the four slide margins in inches, as the pager holds them
 * @returns the usable width in EMU, never below zero
 */
export function usableTableWidthEmu(
	presLayout: PresLayout,
	xEmu: number,
	marginsIn: readonly [number, number, number, number]
): number {
	// `marginsIn` is TRBL, so the left margin is index 3 and the right one index 1. Both were
	// read one index off, which cancels out for a symmetric margin -- the default, and what every
	// gate deck sets -- and is wrong by their difference for anything else. `define/table.ts` had
	// already found and fixed this exact swap at its own sibling site.
	const startIn = xEmu ? xEmu / EMU_PER_INCH : marginsIn[3]
	return Math.max(0, presLayout.width - inch2Emu(startIn) - inch2Emu(marginsIn[1]))
}

/**
 * The single rule for whether one `rowH` entry pins a table row, in inches.
 *
 * An entry pins its row when it reads as a **finite number greater than zero**. Everything
 * else returns `null` — the row is sized from the table's own height, or grows to fit — and
 * an entry that is present but unusable warns, because `rowH: [0, …]` and `rowH: [-1, …]`
 * are things a caller wrote on purpose and none of the readings this replaced was what they
 * meant. A missing slot is not that: an array with holes is how the auto-pager spells
 * "this row is auto-height", so `undefined` is silent.
 *
 * Three call sites read `rowH` and they disagreed. The writer and the export-time fit pass
 * both used a truthiness test, so `0` fell through to the even split of `h`;
 * `pptx.tableLayout()` used `typeof === 'number'`, so `0` pinned the row at zero inches and
 * was then rescued by a default-line fallback — reporting 0.2in against the file's 2.0in,
 * and flagging it `heightExact`. A negative entry reached `<a:tr h="-914400">`.
 *
 * @param entry - one `rowH` array slot, or the scalar `rowH`
 * @returns the pinned height in inches, or `null` when the row is not pinned by `rowH`
 */
export function pinnedRowHeightInches(entry: unknown): number | null {
	if (entry === undefined || entry === null) return null
	const inches = Number(entry)
	if (Number.isFinite(inches) && inches > 0) return inches
	warnOnce(
		'table/invalid-row-height',
		`rowH entry ${String(entry)} is not a positive number of inches; the row is sized from the table height instead.`
	)
	return null
}

/**
 * A table cell's `[Top, Right, Bottom, Left]` margins in inches, defaulted and validated.
 *
 * Three sites spelled this out -- two of them character-identical, the third self-described as
 * "mirroring" one of them -- and a fourth in the auto-pager gated each side on truthiness, so a
 * cell stating `margin: [0, …]` fell through to the table's margin instead of taking its own
 * zero. All of them swapped an unusable margin for the default in silence, while every sibling
 * resolver in this file warns for the same class of input.
 *
 * `resolveInsetsEmu` above is deliberately NOT a fifth copy: it resolves a text box's
 * `a:bodyPr` insets, where stating nothing means writing no attribute at all rather than
 * taking a cell default, so its absent case has a different answer.
 * @param margin - the caller's `margin`: a scalar, a `[T, R, B, L]` array, or nothing
 * @returns four finite inch measures
 */
export function resolveCellMarginsInches(
	margin: number | number[] | undefined | null
): [number, number, number, number] {
	if (margin === undefined || margin === null) return [...DEF_CELL_MARGIN_IN]
	const usable = (value: unknown): boolean => typeof value === 'number' && Number.isFinite(value)
	if (typeof margin === 'number') {
		if (usable(margin)) return [margin, margin, margin, margin]
	} else if (Array.isArray(margin) && margin.length === 4 && margin.every(usable)) {
		return [margin[0] as number, margin[1] as number, margin[2] as number, margin[3] as number]
	}
	warnOnce(
		'table/invalid-margin',
		`table margin ${JSON.stringify(margin)} is not a number of inches or a [top, right, bottom, left] array of four; using the default.`
	)
	return [...DEF_CELL_MARGIN_IN]
}

/**
 * Resolve one table row's height to EMU — the row-height twin of
 * {@link resolveTableColWidthsEmu}, shared by the table XML emitter, the export-time
 * measured-fit pass and `pptx.tableLayout()` so a prediction cannot disagree with what the
 * export bakes.
 *
 * - a `rowH` entry that {@link pinnedRowHeightInches} accepts pins the row
 * - otherwise the table's already-resolved height (`totalHeightEmu`, EMU) is split evenly
 *   across `rowCount` rows
 * - otherwise `null`: the row is auto-height and grows to fit its content
 *
 * @param rowH - the caller's `rowH`, scalar or per-row array
 * @param rowIndex - which row is being resolved
 * @param totalHeightEmu - the table's resolved height in EMU, or 0 when it has none
 * @param rowCount - number of rows to split that height across
 * @returns the row height in EMU, or `null` for an auto-height row
 */
export function resolveTableRowHeightEmu(
	rowH: number | number[] | undefined,
	rowIndex: number,
	totalHeightEmu: number,
	rowCount: number
): number | null {
	const pinned = pinnedRowHeightInches(Array.isArray(rowH) ? rowH[rowIndex] : rowH)
	if (pinned !== null) return Math.round(inch2Emu(pinned))
	return totalHeightEmu > 0 && rowCount > 0 ? Math.round(totalHeightEmu / rowCount) : null
}

/**
 * Convert points to EMU, leniently: anything that does not read as a finite number — a `NaN`,
 * an `Infinity`, a non-numeric string, `null` — becomes `0` rather than throwing.
 *
 * That is the whole difference from {@link pointsToEmu} in `units.ts`, which throws on the same
 * input. Emitters use this one for the many optional size/width/offset options where a missing
 * or malformed value should collapse the feature (a zero-width line, no shadow offset) rather
 * than take the whole deck down.
 *
 * @param pt - points, as a number or a numeric string
 * @returns the value in EMU, or `0` if it does not read as a finite number
 */
export function ptsToEmuLenient(pt: number | string): number {
	const points = Number(pt)
	return Number.isFinite(points) ? Math.round(points * EMU_PER_POINT) : 0
}

/**
 * Clamp a caller-supplied number into the range its schema type allows, warning when it had to
 * move. Returns the value in the caller's own unit — each helper below applies its own
 * arithmetic afterwards, because the inverting and non-inverting forms do not round alike.
 *
 * `NaN` throws rather than clamping: there is no nearest in-range value for it, and
 * `Math.min`/`Math.max` propagate it straight through to the attribute (`val="NaN"`), which is
 * exactly the degenerate output the range check exists to prevent. `Infinity` is not in that
 * category — it clamps to the bound like any other out-of-range number, and warns.
 *
 * **This is the one policy for an out-of-range number**, and it is the policy
 * `docs/diagnostics.md` ("Warn or throw?") describes: a finite value has a nearest legal
 * neighbour, so the deck still comes out recognisable and the move is a warning; a value that
 * is not a number at all has no neighbour, so the request is discarded and that throws.
 * Rejecting a finite out-of-range value and emitting nothing is neither, and was how five
 * options behaved before they were routed through here: it discards the request and reports
 * it as a warning, which is the combination the rule exists to rule out.
 *
 * Most of what passes through here is a percentage, which is why the throw names one by
 * default; `nonFiniteCode` is for the callers whose option is not one, such as a shadow's
 * angle in degrees.
 *
 * @param value - the caller's value
 * @param min - inclusive lower bound, in the caller's unit
 * @param max - inclusive upper bound, in the caller's unit
 * @param code - diagnostic code raised when the value is clamped
 * @param label - option name as the caller spells it, opening the warning
 * @param nonFiniteCode - error code thrown when the value is not a number at all
 */
export function clampRangedInput(
	value: number,
	min: number,
	max: number,
	code: DiagnosticCode,
	label: string,
	nonFiniteCode: InvalidOptionErrorCode = 'percent/non-finite'
): number {
	if (typeof value !== 'number' || Number.isNaN(value))
		throw new InvalidOptionError(
			nonFiniteCode,
			`${label} must be a number from ${min} to ${max}; received ${String(value)}.`
		)
	const clamped = Math.min(max, Math.max(min, value))
	if (clamped !== value) warn(code, `${label} ${value} is outside the valid range ${min}-${max}; using ${clamped}.`)
	return clamped
}

/**
 * Convert a transparency percentage (0-100) into a schema-valid `<a:alpha>` value
 * (ST_PositiveFixedPercentage, 0-100000). Out-of-range transparency yields an
 * alpha that PowerPoint rejects as needing repair, so clamp into range and warn.
 *
 * This is the **inverting** form: 0 transparency is full opacity (`val="100000"`). Use
 * {@link percentToFixedPercent} where the option already reads as opacity.
 */
export function transparencyToAlpha(transparency: number): number {
	const pct = clampRangedInput(transparency, 0, 100, 'transparency/out-of-range', 'transparency')
	return Math.round((100 - pct) * FIXED_PCT_PER_PERCENT)
}

/** Convert an opacity (0-1) into a schema-valid `<a:alpha>` value (0-100000); clamps + warns on out-of-range input. */
export function opacityToAlpha(opacity: number): number {
	return fractionToFixedPercent(opacity, 'opacity/out-of-range', 'opacity')
}

/**
 * Convert a percentage into a schema-valid fixed-percentage value (thousandths of a percent) —
 * the **non-inverting** sibling of {@link transparencyToAlpha}, for options the caller already
 * states as "how much", such as a chart series' fill opacity.
 *
 * `min`/`max` default to the 0-100 nearly every such option allows. They are parameters because
 * the fixed-percentage attributes do not all share one range: `<a:buSzPct>` is
 * ST_TextBulletSizePercent (25-400), so the bullet-size option needs its own bounds while
 * keeping this function's policy.
 * @param value - the caller's percentage
 * @param code - diagnostic code raised when the value is clamped
 * @param label - option name as the caller spells it, opening the warning
 * @param min - inclusive lower bound in percent
 * @param max - inclusive upper bound in percent
 */
export function percentToFixedPercent(value: number, code: DiagnosticCode, label: string, min = 0, max = 100): number {
	return Math.round(clampRangedInput(value, min, max, code, label) * FIXED_PCT_PER_PERCENT)
}

/**
 * Convert a 0-1 fraction into a schema-valid fixed-percentage value (0-100000), for the options
 * spelled as a fraction rather than a percentage (an opacity, a luminance threshold).
 * @param value - the caller's fraction
 * @param code - diagnostic code raised when the value is clamped
 * @param label - option name as the caller spells it, opening the warning
 */
export function fractionToFixedPercent(value: number, code: DiagnosticCode, label: string): number {
	return Math.round(clampRangedInput(value, 0, 1, code, label) * PERCENT_SCALE)
}

/**
 * Convert a length stated in points to EMU, leniently, clamped into the schema range of the
 * attribute it is headed for.
 *
 * The leniency and the clamp answer two different inputs and must not be confused. A value that
 * does not read as a number at all collapses the feature to `0` via {@link ptsToEmuLenient} —
 * a zero-width line, an unoffset shadow — because these are optional decorations on an object
 * that still has to be emitted. A value that *is* a number but falls outside the attribute's
 * range clamps to the nearest bound and warns, per `docs/diagnostics.md`; left alone it would
 * make PowerPoint report the whole package as needing repair.
 *
 * @param pts - the caller's length in points, or any of the loose shapes `ptsToEmuLenient` takes
 * @param maxEmu - inclusive upper bound of the attribute's schema type, in EMU
 * @param code - diagnostic code raised when the value is clamped
 * @param label - option name as the caller spells it, opening the warning
 * @param range - the valid range in points, as the warning states it
 */
function clampLengthEmu(
	pts: number | string,
	maxEmu: number,
	code: DiagnosticCode,
	label: string,
	range: string
): number {
	const raw = ptsToEmuLenient(pts)
	const clamped = Math.min(maxEmu, Math.max(0, raw))
	if (clamped !== raw)
		warn(code, `${label} ${pts} is outside the valid range ${range}; using ${clamped / EMU_PER_POINT}.`)
	return clamped
}

/**
 * Convert a line width (points) to EMU clamped into ST_LineWidth (0..20116800 EMU,
 * i.e. 0-1584pt). Out-of-range widths make PowerPoint report the package as needing
 * repair, so clamp into range and warn.
 */
export function lineWidthToEmu(widthPts: number | string): number {
	return clampLengthEmu(widthPts, 20116800, 'line/width-out-of-range', 'line width', '0-1584pt')
}

/**
 * Convert a shadow's blur radius or offset distance (points) to EMU clamped into
 * ST_PositiveCoordinate (0..27273042316900 EMU, i.e. 0-2147483647pt) — the type both
 * `blurRad` and `dist` carry on `a:outerShdw`/`a:innerShdw`.
 *
 * **The bound is the schema's, not the `0-100` / `0-200` those options document.** Those two
 * numbers are the limits of PowerPoint's own spinners, and a blur past them is unusual rather
 * than invalid: the file loads and paints. The one input that genuinely breaks the package is a
 * *negative* one, because the type is unsigned — so that is what this moves, on the same
 * reasoning `clampLineSpacingMultiplePct` in `gen/drawingml/clamp.ts` clamps to
 * ST_TextSpacingPercent rather than to the `0.0-9.99` its own option describes. Only values
 * PowerPoint reports as needing repair are worth moving.
 */
export function positiveCoordinateEmu(pts: number | string, code: DiagnosticCode, label: string): number {
	return clampLengthEmu(pts, 27273042316900, code, label, '0-2147483647pt')
}

/**
 * Convert degrees to a DrawingML angle (60000ths of a degree), **without** wrapping.
 *
 * Most angles in the format are not modular: a connection site at 400 degrees, a polar adjust
 * handle whose range runs to 540, an arc that sweeps 400 degrees — each means something a
 * reduction into 0..360 would destroy. Use this for those, and {@link convertRotationDegrees}
 * for the ones that genuinely are a rotation.
 *
 * Non-finite input throws rather than reaching the attribute: `Math.round(Infinity * 60000)` is
 * `Infinity`, which serializes as `ang="Infinity"` and makes PowerPoint offer to repair the file.
 *
 * @param d - degrees
 * @param what - what is being converted, for the error message
 * @param code - diagnostic code to throw under, where a caller has a more specific one
 */
export function convertAngleUnits(d: number, what: string, code: AngleErrorCode = 'coord/non-finite'): number {
	assertFiniteDegrees(d, what, code)
	return Math.round(d * ANGLE_UNITS_PER_DEGREE)
}

/** The codes {@link convertAngleUnits} may throw under, each declared in `codes.ts`. */
type AngleErrorCode = 'coord/non-finite' | 'geometry/arc-angle-non-finite'

function assertFiniteDegrees(d: number, what: string, code: AngleErrorCode): void {
	if (typeof d !== 'number' || !Number.isFinite(d))
		throw new InvalidOptionError(code, `${what} must be a finite number of degrees; received ${String(d)}.`)
}

/**
 * Convert a shape or gradient rotation (degrees) to a DrawingML angle (60000ths of a degree).
 *
 * A rotation *is* modular — 800 degrees and 80 degrees point the same way — so the value is
 * reduced with `% 360`. The reduction keeps the sign: `-45` stays `-45` rather than becoming
 * `315`, because both are valid `ST_Angle`, PowerPoint writes negative rotations itself, and
 * the read side reports back what was authored. Only an input outside -360..360 moves.
 *
 * @param d - degrees
 * @returns the `rot` value in 60000ths of a degree
 */
export function convertRotationDegrees(d: number): number {
	const degrees = d || 0
	// Guard before reducing, so `Infinity` is reported as itself rather than as the `NaN` that
	// `Infinity % 360` would hand on.
	assertFiniteDegrees(degrees, 'rotation', 'coord/non-finite')
	return Math.round((degrees % 360) * ANGLE_UNITS_PER_DEGREE)
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
	return convertAngleUnits(d, `Arc ${attr}`, 'geometry/arc-angle-non-finite')
}
