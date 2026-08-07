/**
 * ts-pptx: named clip silhouettes for images
 *
 * A `ClipShape` is declarative data — a named silhouette plus its options — and {@link clipPath}
 * resolves it to the freeform `points` path `addImage` emits as the `<a:custGeom>` clip mask on
 * the `<p:pic>`. Paired with `sizing: { type: 'cover' }` (an `<a:srcRect>` crop) it reproduces
 * what a PowerPoint *picture placeholder* does — a layout `custGeom` clipping an inherited
 * blipFill — but standalone, with no placeholder involved.
 *
 * These are silhouettes recurring often enough in real decks to be worth naming; hand-authored
 * `points` remain available for anything else. Nothing here is theme- or brand-aware: which
 * silhouette to use, and which edge it sits flush to, belongs to the caller.
 *
 * ## The coordinate trap
 *
 * Coordinates come back in the image box's OWN inch space (`0..w`, `0..h`), because `custGeom`
 * points expressed as `%` resolve against the SLIDE, not the box. A box-relative silhouette must
 * therefore be emitted in inches already scaled to the box, which is why {@link clipPath} takes
 * the box size: the fractions below are multiplied by `w` / `h` at build time, so one silhouette
 * scales with whatever region it is handed. Authoring the same shape as percentages would put
 * the clip in the wrong place on every box that is not the full slide.
 */

import type { GeometryPoint } from './types/index.js'

/**
 * Which edge of the box is the half-disc's FLAT side; the arc bulges toward the opposite edge.
 * `'right'` = flat right edge, arc bulging left.
 */
export type FlatSide = 'left' | 'right'

/**
 * Which half-disc proportion to trace. Both are flat-on-one-edge, bulging half-discs whose arc
 * reaches the opposite edge near mid-height; they differ only in how far the arc reaches.
 * - `'deep'` — the arc spans about 32% of the box width, symmetric about mid-height (the default).
 * - `'shallow'` — the arc spans about 13%, slightly asymmetric, with its apex just below mid-height.
 */
export type HalfDiscPreset = 'deep' | 'shallow'

/** A named clip silhouette, resolved to a `points` path by {@link clipPath}. */
export type ClipShape =
	/**
	 * A rectangle with one side replaced by an elliptical arc that bulges toward the opposite
	 * edge, reaching it near mid-height — the "D" a picture placeholder cuts on a cover slide.
	 */
	{ kind: 'half-disc'; flat: FlatSide; preset?: HalfDiscPreset }

/**
 * A half-disc as box fractions, flat side on the RIGHT (`x = 1`), the arc bulging LEFT toward
 * `x = 0` near mid-height. Two cubic Béziers trace the half-ellipse: the path runs
 * top-flat-corner → top-right → bottom-right → bottom-flat-corner → bottom tip, cubic up to the
 * mid-left apex, cubic back to the top tip, then closes.
 *
 * Each preset spells out every fraction rather than sharing a mirrored `topEdgeX`, because the
 * presets are independent freeforms — the shallow one is asymmetric top-to-bottom, so a
 * symmetric parameterization could not express it without changing its shape.
 */
interface HalfDiscSpec {
	/** x where the TOP edge departs the flat side toward the arc. */
	topEdgeX: number
	/** x where the BOTTOM edge departs the flat side (may differ from the top). */
	botEdgeX: number
	/** x where the arc leaves the top edge. */
	tipTopX: number
	/** y where the arc leaves the top edge. */
	tipTopY: number
	/** x where the arc leaves the bottom edge. */
	tipBotX: number
	/** y where the arc leaves the bottom edge. */
	tipBotY: number
	/** y of the bulge apex (at `x = 0`). */
	apexY: number
	/** Bottom-segment controls: from the bottom tip, cubic to the apex. */
	c1b: { x: number; y: number }
	c2b: { x: number; y: number }
	/** Top-segment controls: from the apex, cubic back to the top tip. */
	c1t: { x: number; y: number }
	c2t: { x: number; y: number }
}

/** `deep` — symmetric top/bottom, apex at mid-height. */
const DEEP_DISC: HalfDiscSpec = {
	topEdgeX: 0.31788,
	botEdgeX: 0.31788,
	tipTopX: 0.30061,
	tipTopY: 0.01078,
	tipBotX: 0.30061,
	tipBotY: 0.98922,
	apexY: 0.5,
	c1b: { x: 0.11594, y: 0.86833 },
	c2b: { x: 0.0, y: 0.69392 },
	c1t: { x: 0.0, y: 0.30609 },
	c2t: { x: 0.11594, y: 0.13167 },
}

/** `shallow` — a slightly asymmetric top/bottom inset and an apex just below mid-height. */
const SHALLOW_DISC: HalfDiscSpec = {
	topEdgeX: 0.12611,
	botEdgeX: 0.11128,
	tipTopX: 0.09976,
	tipTopY: 0.05448,
	tipBotX: 0.09976,
	tipBotY: 0.97617,
	apexY: 0.51533,
	c1b: { x: 0.03522, y: 0.83247 },
	c2b: { x: 0.0, y: 0.67725 },
	c1t: { x: 0.0, y: 0.3534 },
	c2t: { x: 0.03522, y: 0.19818 },
}

const HALF_DISC: Record<HalfDiscPreset, HalfDiscSpec> = { deep: DEEP_DISC, shallow: SHALLOW_DISC }

/**
 * Half-disc clip path for a `w × h` box. Authored flat-right; `flat: 'left'` mirrors `x → 1 - x`.
 */
function halfDisc(w: number, h: number, flat: FlatSide, preset: HalfDiscPreset): GeometryPoint[] {
	const s = HALF_DISC[preset]
	const fx = flat === 'left' ? (x: number): number => (1 - x) * w : (x: number): number => x * w
	const fy = (y: number): number => y * h
	return [
		{ x: fx(s.topEdgeX), y: fy(0), moveTo: true },
		{ x: fx(1), y: fy(0) },
		{ x: fx(1), y: fy(1) },
		{ x: fx(s.botEdgeX), y: fy(1) },
		{ x: fx(s.tipBotX), y: fy(s.tipBotY) },
		{
			x: fx(0),
			y: fy(s.apexY),
			curve: { type: 'cubic', x1: fx(s.c1b.x), y1: fy(s.c1b.y), x2: fx(s.c2b.x), y2: fy(s.c2b.y) },
		},
		{
			x: fx(s.tipTopX),
			y: fy(s.tipTopY),
			curve: { type: 'cubic', x1: fx(s.c1t.x), y1: fy(s.c1t.y), x2: fx(s.c2t.x), y2: fy(s.c2t.y) },
		},
		{ close: true },
	]
}

/**
 * Resolve a {@link ClipShape} to a freeform `points` path filling a `w × h` box, ready to hand
 * to `addImage({ points })`.
 *
 * `w`/`h` are the image box's size **in inches** and must match the `w`/`h` the picture is drawn
 * at — the path is emitted in the box's own coordinate space, not normalized (see the module
 * note above).
 * @param {ClipShape} shape - the named silhouette and its options
 * @param {number} w - image box width (inches)
 * @param {number} h - image box height (inches)
 * @returns {GeometryPoint[]} the freeform path, for `addImage({ points })`
 * @example
 * const w = 5.22, h = 7.5
 * slide.addImage({
 *   path: 'cover.jpg', x: 0, y: 0, w, h,
 *   points: clipPath({ kind: 'half-disc', flat: 'right' }, w, h),
 *   sizing: { type: 'cover' },   // center-crop the photo into the "D"
 * })
 */
export function clipPath(shape: ClipShape, w: number, h: number): GeometryPoint[] {
	switch (shape.kind) {
		case 'half-disc':
			return halfDisc(w, h, shape.flat, shape.preset ?? 'deep')
	}
}
