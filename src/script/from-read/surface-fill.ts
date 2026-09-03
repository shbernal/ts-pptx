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
import type { PictureFill } from '../../read/api/picture-fill.js'
import type { ResolvedColor } from '../../read/api/theme-context.js'
import type { AssetResolver } from './context.js'
import { pictureFillOption, type PictureFillSubject } from './picture-fill.js'
import type { NoteScope } from '../fidelity.js'
import type { IrValue } from '../ir.js'
import { alphaToTransparency, colorOption, compact, isWritableSchemeToken, literalColor } from './values.js'

/** How a note names the surface a gradient sits on: `fill.gradient`, `table.fill.gradient`, … */
type FillNoteScope = 'fill' | 'line' | 'table.fill' | 'table.cell.fill'

/**
 * A read gradient as `GradientFillProps`. Stop positions convert from the read model's
 * 0–1 fraction to the write API's 0–100 percentage; the angle needs no conversion, since
 * both sides use OOXML degrees (clockwise from 3 o'clock).
 */
export function gradientStops(gradient: GradientFill, notes: NoteScope, where: FillNoteScope): IrValue | undefined {
	const stops = gradient.stops
		.map((stop) => {
			// Scoped by `where`, like the two notes below it. Hardcoding the shape spelling meant a
			// table background's gradient stop recorded `fill.gradient.schemeToken`, whose entry
			// excuses a difference on `fill` — while the difference itself lands on `tableFill`,
			// which the table mapper keeps deliberately distinct. A declared, genuine loss was
			// therefore reported as an undeclared defect, and on a line gradient the note's prose
			// said "fill" about a stroke.
			const color = colorOption(
				{ scheme: stop.schemeColor, resolvedHex: stop.effectiveHex },
				notes,
				`${where}.gradient.schemeToken`,
				`${where === 'line' ? 'line' : 'fill'} gradient stop`
			)
			if (color === undefined) return null
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

/**
 * The three surfaces a fill can sit on, and how a note names each. One record rather than three
 * constants in three modules, because every field of it is decided by `where`.
 */
const FILL_SURFACES: Record<FillSurface, { label: string; picture: PictureFillSubject }> = {
	fill: {
		label: 'fill',
		picture: { construct: 'fill.picture', subject: "this shape's surface", element: 'a:blipFill' },
	},
	'table.fill': {
		label: 'table fill',
		picture: { construct: 'table.fill.picture', subject: 'this table', element: 'a:tblPr/a:blipFill' },
	},
	'table.cell.fill': {
		label: 'cell fill',
		picture: { construct: 'table.cell.fill.picture', subject: 'this table cell', element: 'a:tcPr/a:blipFill' },
	},
}

/** The surfaces {@link surfaceFill} serves; `line` is a stroke and has its own ladder. */
export type FillSurface = 'fill' | 'table.fill' | 'table.cell.fill'

/**
 * What {@link surfaceFill} reads. Structural rather than a union of the three read-model
 * classes, because that is exactly the point: a shape, a table and a table cell expose the same
 * accessor names for the same `EG_FillProperties`, which is why three copies of this ladder
 * could exist and drift.
 */
export interface FillSubject {
	fillNoFill?: boolean
	gradientFill: GradientFill | null
	patternFill: PatternFill | null
	pictureFill: PictureFill | null
	fillSchemeColor: string | null
	/** The surface's OWN `a:srgbClr`, where the read model separates it from the resolved one. */
	fillColor?: string | null
	resolvedFill: ResolvedColor | null
	/** `false` on a styled surface means the fill came from the style, not from the surface. */
	hasOwnFill?: boolean
}

/**
 * A surface's `EG_FillProperties` as the write API's fill option, or `undefined` when the
 * surface states none.
 *
 * The order is the whole of it, and it is not arbitrary: an explicit `a:noFill` is a statement
 * and comes first; then the non-solid members, because a surface carrying one holds no
 * `a:solidFill` for the colour legs to find and `resolvedFill` would answer with a style's
 * banding colour the source never showed; then the colour ladder — a writable scheme token, the
 * surface's own literal, the resolved literal.
 *
 * Three copies of this walked three subjects that expose the same accessor names, and they had
 * drifted: only the shape's read `alphaModFix`, so a table or cell whose `a:solidFill` carried
 * one lost its transparency with nothing declaring the loss.
 *
 * @param subject - the shape, table or cell
 * @param ctx - the mapping context, for notes and the asset resolver
 * @param where - which surface this is, deciding the note constructs and prose
 * @param opts - `styled` short-circuits a styled surface with no fill of its own
 */
export function surfaceFill(
	subject: FillSubject,
	ctx: { notes: NoteScope; assets: AssetResolver },
	where: FillSurface,
	opts: { styled?: boolean } = {}
): IrValue | undefined {
	const { notes, assets } = ctx
	const surface = FILL_SURFACES[where]

	if (subject.fillNoFill) return { type: 'none' }

	const gradient = subject.gradientFill
	if (gradient) {
		const stops = gradientStops(gradient, notes, where)
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const pattern = patternOption(subject.patternFill)
	if (pattern) return pattern

	const picture = subject.pictureFill
	if (picture) return pictureFillOption(picture, assets, notes, surface.picture)

	// The token and the surface's own literal come before the style short-circuit: both are the
	// surface stating a colour of its own, which is what the short-circuit is testing for.
	const scheme = subject.fillSchemeColor
	if (isWritableSchemeToken(scheme)) return { color: scheme as string }
	if (subject.fillColor != null) return { color: literalColor(subject.fillColor) }

	// A styled surface with no fill of its own takes the style's banding, and the style GUID
	// travels with the table — so emitting nothing here is not a loss, it is what keeps the copy
	// responsive to its own style.
	if (opts.styled && subject.hasOwnFill === false) return undefined

	const resolved = subject.resolvedFill
	if (!resolved) return undefined
	const color = colorOption(
		{ scheme, resolvedHex: resolved.effectiveHex },
		notes,
		`${where}.schemeToken`,
		surface.label
	)
	// The transparency an `a:alphaModFix` on the surface's own `a:solidFill` carries. Only the
	// shape's copy of this ladder read it, so a table or a cell lost it with nothing declaring
	// the loss.
	return compact({ color, transparency: alphaToTransparency(resolved.alpha) })
}
