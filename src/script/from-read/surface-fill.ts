/**
 * The non-solid halves of a surface's fill — `a:gradFill` and `a:pattFill` — as
 * `ShapeFillProps`.
 *
 * Shared by a shape's fill, a table's background, and a table cell's, which are the same
 * two elements read from three containers. `ShapeProps.fill`, `TableProps.tableFill` and
 * `TableCellProps.fill` are all `ShapeFillProps`, so the mapping is identical and only the
 * note's `where` prefix differs — the same reasoning that keeps `pictureFillOption` in one
 * place rather than one copy per surface.
 */
import type { GradientFill } from '../../read/api/gradient.js'
import type { PatternFill } from '../../read/api/pattern-fill.js'
import type { NoteScope } from '../fidelity.js'
import type { IrValue } from '../ir.js'
import { alphaToTransparency, compact, isWritableSchemeToken, literalColor } from './values.js'

/** How a note names the surface a gradient sits on: `fill.gradient`, `table.fill.gradient`, … */
export type FillNoteScope = 'fill' | 'line' | 'table.fill' | 'table.cell.fill'

/**
 * A read gradient as `GradientFillProps`. Stop positions convert from the read model's
 * 0–1 fraction to the write API's 0–100 percentage; the angle needs no conversion, since
 * both sides use OOXML degrees (clockwise from 3 o'clock).
 */
export function gradientStops(gradient: GradientFill, notes: NoteScope, where: FillNoteScope): IrValue | undefined {
	const stops = gradient.stops
		.map((stop) => {
			const color = isWritableSchemeToken(stop.schemeColor)
				? (stop.schemeColor as string)
				: stop.effectiveHex === null
					? null
					: literalColor(stop.effectiveHex)
			if (color === null) return null
			return compact({
				color,
				position: Math.round((stop.position ?? 0) * 100),
				transparency: alphaToTransparency(stop.alpha),
			}) as IrValue
		})
		.filter((stop): stop is IrValue => stop !== null)

	if (stops.length < 2) {
		notes.note(
			`${where}.gradient`,
			'dropped',
			'unsupported',
			'a gradient with fewer than two resolvable stops cannot be expressed, so this falls back to no gradient'
		)
		return undefined
	}

	if (gradient.kind === 'path') {
		// `a:path` covers circle/rect/shape; the write API models only the radial case, so
		// a rectangular or shape-following gradient becomes a circular one.
		if (gradient.path !== null && gradient.path !== 'circle') {
			notes.note(
				`${where}.gradient.path`,
				'approximated',
				'unwritable',
				`a "${gradient.path}" path gradient is emitted as a radial one; the write API models no other path shape`
			)
		}
		return { kind: 'radial', stops }
	}
	return compact({ kind: 'linear', angle: gradient.angleDeg ?? 0, stops })
}

/**
 * A read hatch as `ShapeFillProps` with `type: 'pattern'`, or `undefined` when the pattern
 * names no preset — `a:pattFill/@prst` is what selects the hatch, so without it there is
 * nothing to reproduce.
 *
 * Both colours are emitted as literals rather than scheme tokens: `PatternFillProps` types
 * `fgColor`/`bgColor` as hex, so a token would not survive the write side anyway.
 */
export function patternOption(pattern: PatternFill | null): IrValue | undefined {
	if (!pattern?.preset) return undefined
	return {
		type: 'pattern',
		pattern: compact({
			preset: pattern.preset,
			fgColor: pattern.foreground ? literalColor(pattern.foreground.effectiveHex) : undefined,
			bgColor: pattern.background ? literalColor(pattern.background.effectiveHex) : undefined,
		}) ?? { preset: pattern.preset },
	}
}
