/**
 * The paint and placement a shape carries whatever kind it is: transform, fill, line, arrows,
 * shadow, glow, and the `p:style` tier behind them.
 *
 * Split out of `shape.ts`, where these sat above the five per-kind call builders and were
 * called by all of them. They are also where most of this converter's *fidelity notes* are
 * recorded, because the binding constraint is the read side: a shape whose outline comes from
 * `p:style/a:lnRef` reports a colour through `resolvedLine` and `null` for width and dash,
 * since there is a resolved-colour path and no resolved-width path. Each such loss is declared
 * here rather than quietly thinned.
 */
import type { AnyShape } from '../../read/api/shapes.js'
import type { NoteScope } from '../fidelity.js'
import { type IrValue } from '../ir.js'
import {
	alphaToTransparency,
	compact,
	isWritableSchemeToken,
	literalColor,
	orUndefined,
	WRITABLE_DASHES,
} from './values.js'
import { type FillSubject, gradientStops, surfaceFill } from './surface-fill.js'
import { type MapContext } from './context.js'

/** Arrowhead types `ShapeLineProps` accepts; `a:headEnd/@type` uses the same tokens. */
const WRITABLE_ARROWS = new Set(['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'])

/**
 * `a:ln/@cap` tokens → the `ShapeLineProps.cap` spelling that authors them. The read side
 * reports the raw OOXML token (`AutoShape.lineCap`) and the write side takes a friendly
 * name, so this is the whole of the mapping between them. `ST_LineCap` has exactly these
 * three members, so nothing falls off the end.
 */
const CAP_TOKENS: Record<string, 'flat' | 'square' | 'round'> = {
	flat: 'flat',
	sq: 'square',
	rnd: 'round',
}

/**
 * `p:cNvPr/@hidden` has no write-API counterpart, so a hidden shape would come back visible.
 * Omitting it preserves what the deck looks like, which is the lesser loss. Shared with the
 * layout-decoration mapper below, whose connector arm does not pass through {@link shapeCall}.
 */
export function noteHidden(notes: NoteScope): void {
	notes.note(
		'shape.hidden',
		'dropped',
		'unwritable',
		'a hidden shape has no write-API expression; it is omitted rather than emitted visible'
	)
}

/* ===== position, fill, line, effects — shared by every shape kind ===== */

/**
 * Rotation and flips. Taken from `absoluteFrame`, which composes enclosing group rotations
 * and XOR-composes group flips — the same reason position is absolute.
 */
export function transformOptions(shape: AnyShape): Record<string, IrValue | undefined> {
	const frame = shape.absoluteFrame
	const rotate = frame ? frame.rotation : (shape.rotation ?? 0)
	return {
		rotate: rotate === 0 ? undefined : rotate,
		flipH: (frame?.flipH ?? shape.flipH) ? true : undefined,
		flipV: (frame?.flipV ?? shape.flipV) ? true : undefined,
	}
}

/**
 * A shape's fill as `ShapeFillProps`.
 *
 * Prefers the raw scheme token over `resolvedFill` for the same reason run colour does: a
 * token keeps tracking the destination theme. `resolvedFill` is the fallback because it is
 * the only accessor that sees a `p:style/a:fillRef` — a shape styled entirely from the
 * theme has no `a:solidFill` of its own, so without it the shape would come out unfilled.
 */
export function fillOption(shape: AnyShape, ctx: MapContext): IrValue | undefined {
	// A shape has no explicit-noFill leg here: `lineOption` handles the stroke's, and a shape's
	// own `a:noFill` reaches the write side as an absent fill option, which is the same output.
	return surfaceFill({ ...fillAccessors(shape), fillColor: shape.fillColor }, ctx, 'fill')
}

/**
 * The `EG_FillProperties` accessors {@link surfaceFill} reads, projected off a read-model shape.
 *
 * A projection rather than passing the shape itself: `AnyShape` is a union whose members carry
 * far more than this, and naming the six keys is what makes "a shape, a table and a cell answer
 * the same six questions" a checked statement rather than an observation.
 */
function fillAccessors(shape: AnyShape): FillSubject {
	return {
		gradientFill: shape.gradientFill,
		patternFill: shape.patternFill,
		pictureFill: shape.pictureFill,
		fillSchemeColor: shape.fillSchemeColor,
		resolvedFill: shape.resolvedFill,
	}
}

/**
 * Everything about an outline that does not survive, noted.
 *
 * Three independent conditions, none of which decides any of `lineOption`'s emitted values --
 * they were interleaved with the option build, which is what made that function four
 * decisions wearing one name.
 *
 * @param values - the outline's already-read parts, so this does not re-read the DOM
 */
function noteLineLosses(
	shape: AnyShape,
	values: {
		widthPt: number | null
		dash: string | null
		scheme: string | null
		resolved: { effectiveHex: string } | null
	},
	notes: NoteScope
): void {
	const { widthPt, dash, scheme, resolved } = values

	// A colour only `resolvedLine` could supply means the stroke came from the theme style
	// list, whose width and dash are unreadable.
	if (resolved !== null && shape.lineColor === null && scheme === null && widthPt === null) {
		notes.note(
			'line.width',
			'dropped',
			'unread',
			'this outline comes from the theme style list (p:style/a:lnRef); its width and dash live in the theme fmtScheme a:lnStyleLst, which has no accessor, so only the colour carries'
		)
	}
	if (dash !== null && !WRITABLE_DASHES.has(dash)) {
		notes.note(
			'line.dash',
			'approximated',
			'unwritable',
			`dash style "${dash}" is outside the eight the write API accepts, so the outline falls back to solid`
		)
	}
	// `@algn` is the one outline attribute both sides can see and neither can carry: the read
	// model reports it (`lineAlign`) and `ShapeLineProps` has no option for it. Only `in` is
	// noted -- it insets the stroke by half its width, so it moves the border. `ctr` is what an
	// omitted `@algn` already renders as, so noting that would fire on most PowerPoint-authored
	// shapes while describing no loss at all.
	if (shape.lineAlign === 'in') {
		notes.note(
			'line.align',
			'dropped',
			'unwritable',
			'this outline is inset (a:ln/@algn="in", drawn wholly inside the shape); ShapeLineProps has no alignment option, so it comes back centred on the edge and sits half its width further out'
		)
	}
}

/**
 * A shape's outline as `ShapeLineProps`.
 *
 * This is where the measured silent loss lives. When a shape takes its outline from
 * `p:style/a:lnRef`, `resolvedLine` supplies the colour but `lineWidthPt` and `lineDash`
 * both report `null` — the theme's `a:lnStyleLst` entry, where the width and dash actually
 * live, has no accessor. Emitting the colour and noting the rest is the honest outcome;
 * inventing a width would be worse, because it would look deliberate.
 */
export function lineOption(shape: AnyShape, notes: NoteScope): Record<string, IrValue> | undefined {
	if (shape.lineNoFill) return { type: 'none' }

	const gradient = shape.lineGradient
	if (gradient) {
		const stops = gradientStops(gradient, notes, 'line')
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const widthPt = shape.lineWidthPt
	const dash = shape.lineDash
	const cap = shape.lineCap
	const scheme = shape.lineSchemeColor
	const resolved = shape.resolvedLine

	let color: string | undefined
	if (isWritableSchemeToken(scheme)) color = scheme as string
	else if (shape.lineColor !== null) color = literalColor(shape.lineColor)
	else if (resolved) color = literalColor(resolved.effectiveHex)

	noteLineLosses(shape, { widthPt, dash, scheme, resolved }, notes)

	const options = compact({
		color,
		width: orUndefined(widthPt),
		dashType: dash !== null && WRITABLE_DASHES.has(dash) ? dash : undefined,
		cap: cap === null ? undefined : CAP_TOKENS[cap],
		transparency: alphaToTransparency(resolved?.alpha),
		...arrowOptions(shape, notes),
	})
	return options
}

/** Arrowheads. Types map directly; the `sm`/`med`/`lg` size classes have no option. */
function arrowOptions(shape: AnyShape, notes: NoteScope): Record<string, IrValue | undefined> {
	const ends = shape.lineEnds
	if (!ends) return {}
	const sized = [ends.head, ends.tail].some((end) => end && (end.width !== null || end.length !== null))
	if (sized) {
		notes.note(
			'line.arrowSize',
			'dropped',
			'unwritable',
			'arrowhead width and length classes (@w / @len) have no write-API option, so arrowheads render at the default size'
		)
	}
	return {
		beginArrowType: ends.head && WRITABLE_ARROWS.has(ends.head.type) ? ends.head.type : undefined,
		endArrowType: ends.tail && WRITABLE_ARROWS.has(ends.tail.type) ? ends.tail.type : undefined,
	}
}

/** Outer shadow as `ShadowProps`; an inner shadow uses the same option with `type: 'inner'`. */
export function shadowOption(shape: AnyShape, notes: NoteScope): IrValue | undefined {
	if (shape.reflection || shape.softEdge) {
		notes.note(
			'shape.effects',
			'dropped',
			'unwritable',
			'a:reflection and a:softEdge are read but have no write-API emitter, so those effects are lost'
		)
	}

	const outer = shape.shadow
	const source = outer ?? shape.innerShadow
	if (!source) return undefined

	return compact({
		type: outer ? 'outer' : 'inner',
		color: source.color === null ? undefined : literalColor(source.color),
		blur: source.blurPt,
		offset: source.offsetPt,
		angle: source.angleDeg,
		transparency: alphaToTransparency(source.alpha),
	})
}

/** Glow. The read model reports opacity as a 0–1 alpha, which is the write API's unit too. */
export function glowOption(shape: AnyShape): IrValue | undefined {
	const glow = shape.glow
	if (!glow) return undefined
	return compact({
		size: glow.radiusPt,
		color: glow.color === null ? undefined : literalColor(glow.color),
		opacity: glow.alpha,
	})
}

/** The style block every shape kind shares. */
export function styleOptions(shape: AnyShape, ctx: MapContext): Record<string, IrValue | undefined> {
	const { notes } = ctx
	return {
		fill: fillOption(shape, ctx),
		line: lineOption(shape, notes),
		shadow: shadowOption(shape, notes),
		glow: glowOption(shape),
	}
}

/* ===== per-kind mappers ===== */

/**
 * An `AutoShape` becomes an `addShape` when it is bare geometry and an `addText` when it
 * carries text. A shape with both takes the `addText` form, because `addShape` has no text
 * argument while `addText` accepts a `shape` option — so only that form expresses both.
 */
