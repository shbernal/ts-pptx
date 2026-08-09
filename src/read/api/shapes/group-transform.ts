/**
 * Group-shape transform composition — the `p:grpSp` child-space mapping, shared
 * by every surface that needs slide-absolute geometry.
 *
 * A `p:grpSp` declares two frames on its `a:xfrm`: where it sits on the slide
 * (`a:off`/`a:ext`) and the coordinate space its children are authored in
 * (`a:chOff`/`a:chExt`). A child's own `a:off`/`a:ext` is expressed in that child
 * space, so it is not directly placeable on the slide — it must be mapped through
 * `off + (p - chOff) * (ext / chExt)` for every enclosing group, with the group's
 * `@rot`/`@flipH`/`@flipV` composed on top.
 *
 * This module is deliberately **pure**: it takes decoded numbers, not XML, so the
 * mapping stays one implementation no matter how many surfaces need it. It was
 * shared between two independent readers before the inspect surface became a
 * projection over the read model; today `Shape.absoluteFrame` is the only caller,
 * and the separation still earns its keep by keeping the arithmetic testable
 * without a package to parse.
 */

/** A position + extent pair in EMU (`a:off`/`a:ext`, or `a:chOff`/`a:chExt`). */
interface TransformBox {
	x: number
	y: number
	cx: number
	cy: number
}

/** A shape's own `a:xfrm`, decoded. `rotation` is in degrees (the raw `@rot` / 60000). */
interface ShapeTransform {
	box: TransformBox
	rotation: number
	flipH: boolean
	flipV: boolean
}

/**
 * An enclosing `p:grpSp`'s `a:xfrm`, decoded: `outer` is the group's own frame on
 * the slide (`a:off`/`a:ext`), `child` the space its children are authored in
 * (`a:chOff`/`a:chExt`).
 */
export interface GroupTransform {
	outer: TransformBox
	child: TransformBox
	rotation: number
	flipH: boolean
	flipV: boolean
}

/** A slide-absolute frame: the unrotated placement box plus effective orientation. */
interface ComposedFrame {
	box: TransformBox
	/** Effective clockwise rotation in degrees, normalised to [0, 360). */
	rotation: number
	/** Effective horizontal flip, XOR-composed across the group chain. */
	flipH: boolean
	/** Effective vertical flip, XOR-composed across the group chain. */
	flipV: boolean
}

/** Normalise a degree value into [0, 360). */
function normalizeDegrees(value: number): number {
	return ((value % 360) + 360) % 360
}

/** Rotate `point` clockwise about `center` by `degrees`. */
function rotatePoint(
	point: { x: number; y: number },
	center: { x: number; y: number },
	degrees: number
): { x: number; y: number } {
	if (degrees === 0) return point
	const angle = (degrees * Math.PI) / 180
	const cos = Math.cos(angle)
	const sin = Math.sin(angle)
	const dx = point.x - center.x
	const dy = point.y - center.y
	return {
		x: center.x + cos * dx - sin * dy,
		y: center.y + sin * dx + cos * dy,
	}
}

/**
 * Compose a shape's own transform with every enclosing group transform into a
 * slide-absolute frame.
 *
 * `groups` runs **innermost first** (the shape's immediate `p:grpSp` parent, then
 * outward). An empty list returns the shape's own box and orientation unchanged.
 *
 * The returned `box` is the *unrotated* placement box — the same box PowerPoint
 * writes when you Ungroup — with the effective rotation reported separately.
 * Values are exact; callers round or convert units as they see fit.
 *
 * Returns `null` when a group's `a:chExt` is degenerate (zero on either axis):
 * the child-space ratio is then undefined and there is no resolvable frame.
 */
export function composeGroupFrame(shape: ShapeTransform, groups: GroupTransform[]): ComposedFrame | null {
	let center = { x: shape.box.x + shape.box.cx / 2, y: shape.box.y + shape.box.cy / 2 }
	let width = shape.box.cx
	let height = shape.box.cy
	let flipH = shape.flipH
	let flipV = shape.flipV

	// Group centres are collected on the way out, then rotated on the way back in:
	// an outer group rotates the centres of the inner groups nested inside it, so
	// the two passes cannot be merged.
	const chain: { center: { x: number; y: number }; rotation: number; flipH: boolean; flipV: boolean }[] = []

	for (const group of groups) {
		if (group.child.cx === 0 || group.child.cy === 0) return null // degenerate child frame — no resolvable mapping
		const ratioX = group.outer.cx / group.child.cx
		const ratioY = group.outer.cy / group.child.cy

		const mapPoint = (point: { x: number; y: number }): { x: number; y: number } => {
			let x = group.outer.x + (point.x - group.child.x) * ratioX
			let y = group.outer.y + (point.y - group.child.y) * ratioY
			if (group.flipH) x = group.outer.x + group.outer.cx - (x - group.outer.x)
			if (group.flipV) y = group.outer.y + group.outer.cy - (y - group.outer.y)
			return { x, y }
		}
		center = mapPoint(center)
		for (const entry of chain) entry.center = mapPoint(entry.center)
		width *= Math.abs(ratioX)
		height *= Math.abs(ratioY)
		chain.push({
			center: { x: group.outer.x + group.outer.cx / 2, y: group.outer.y + group.outer.cy / 2 },
			rotation: group.rotation,
			flipH: group.flipH,
			flipV: group.flipV,
		})
		flipH = flipH !== group.flipH
		flipV = flipV !== group.flipV
	}

	let rotation = 0
	// A group flipped on exactly one axis mirrors the sense of every rotation
	// nested inside it, so track orientation as the chain is walked back inward.
	let orientation = 1
	for (let index = chain.length - 1; index >= 0; index--) {
		const group = chain[index]
		if (!group) continue
		const groupRotation = group.rotation * orientation
		center = rotatePoint(center, group.center, groupRotation)
		for (let innerIndex = 0; innerIndex < index; innerIndex++) {
			const inner = chain[innerIndex]
			if (inner) inner.center = rotatePoint(inner.center, group.center, groupRotation)
		}
		rotation += groupRotation
		if (group.flipH !== group.flipV) orientation *= -1
	}

	return {
		box: { x: center.x - width / 2, y: center.y - height / 2, cx: width, cy: height },
		rotation: normalizeDegrees(rotation + shape.rotation * orientation),
		flipH,
		flipV,
	}
}
