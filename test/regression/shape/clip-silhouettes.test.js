import { PNG_1X1, defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'
import { clipPath, EMU_PER_INCH } from '../../../dist/node.js'

// `clipPath` resolves a named silhouette to the freeform `points` path `addImage` emits as a
// `<a:custGeom>` clip. The one non-obvious property it must keep is the coordinate space:
// custGeom points given as `%` resolve against the SLIDE, not the picture box, so the path has
// to come back in the box's OWN inch space — every fraction pre-multiplied by w/h. A silhouette
// that leaked normalized (0..1) coordinates would clip a sliver at the slide's top-left corner
// on every box smaller than the slide, which is why the scaling is asserted directly.

/**
 * The resolved path as plain records. `GeometryPoint` is a union whose arms carry different
 * keys, and these tests deliberately look across all of them (is this node a curve? where does
 * it sit?), so the union is flattened once here rather than narrowed at every read.
 * @param {Parameters<typeof clipPath>[0]} shape
 * @param {number} w
 * @param {number} h
 * @returns {{ x?: number, y?: number, moveTo?: boolean, close?: boolean,
 *            curve?: { type: string, x1: number, y1: number, x2: number, y2: number } }[]}
 */
function path(shape, w, h) {
	return /** @type {any} */ (clipPath(shape, w, h))
}

/** Every x (or y) appearing in the path, control points included. */
function axisValues(points, axis) {
	const out = []
	for (const p of points) {
		const v = p[axis]
		if (typeof v === 'number') out.push(v)
		if (p.curve) out.push(p.curve[axis === 'x' ? 'x1' : 'y1'], p.curve[axis === 'x' ? 'x2' : 'y2'])
	}
	return out
}

defineRegressionSuite('Clip silhouettes (clipPath)', [
	{
		name: 'half-disc traces a closed path of straight edges plus two cubic segments',
		fn: () => {
			const p = path({ kind: 'half-disc', flat: 'right' }, 4, 6)
			assertEqual(p.length, 8, 'half-disc node count')
			assertEqual(p[0].moveTo, true, 'path must open with a moveTo')
			assertEqual(p[p.length - 1].close, true, 'path must close')
			const curves = p.filter((n) => n.curve)
			assertEqual(curves.length, 2, 'the half-ellipse is two cubic Béziers')
			assert(
				curves.every((n) => n.curve?.type === 'cubic'),
				'both curve nodes must be cubic'
			)
		},
	},
	{
		name: 'coordinates come back in the box inch space, scaled to w/h — never normalized',
		fn: () => {
			const w = 4
			const h = 6
			const p = path({ kind: 'half-disc', flat: 'right' }, w, h)
			const xs = axisValues(p, 'x')
			const ys = axisValues(p, 'y')
			assert(Math.max(...xs) === w, `widest x must reach the box width ${w}; got ${Math.max(...xs)}`)
			assert(Math.max(...ys) === h, `lowest y must reach the box height ${h}; got ${Math.max(...ys)}`)
			assert(Math.min(...xs) === 0 && Math.min(...ys) === 0, 'the path must reach the box origin')

			// The same silhouette at double the size is exactly double: one shape, any region.
			const doubled = path({ kind: 'half-disc', flat: 'right' }, w * 2, h * 2)
			assertEqual(
				JSON.stringify(axisValues(doubled, 'x')),
				JSON.stringify(xs.map((v) => v * 2)),
				'silhouette must scale linearly with the box'
			)
		},
	},
	{
		name: "flat: 'left' mirrors the silhouette about the box's vertical centre",
		fn: () => {
			const right = path({ kind: 'half-disc', flat: 'right' }, 4, 6)
			const left = path({ kind: 'half-disc', flat: 'left' }, 4, 6)
			assertEqual(
				JSON.stringify(axisValues(left, 'x')),
				JSON.stringify(axisValues(right, 'x').map((v) => 4 - v)),
				'left-flat must be the right-flat path mirrored in x'
			)
			assertEqual(
				JSON.stringify(axisValues(left, 'y')),
				JSON.stringify(axisValues(right, 'y')),
				'mirroring must not touch y'
			)
		},
	},
	{
		name: 'the two presets differ only in how far the arc bulges; deep is the default',
		fn: () => {
			const deep = path({ kind: 'half-disc', flat: 'right', preset: 'deep' }, 4, 6)
			const shallow = path({ kind: 'half-disc', flat: 'right', preset: 'shallow' }, 4, 6)
			assertEqual(
				JSON.stringify(path({ kind: 'half-disc', flat: 'right' }, 4, 6)),
				JSON.stringify(deep),
				'omitted preset must resolve to deep'
			)
			// Both open where the top edge leaves the flat side; with the flat side on the right
			// that x IS the bulge depth, so it is what separates the two presets.
			assert(
				Number(deep[0].x) > Number(shallow[0].x),
				`deep must bulge further into the box (deep ${deep[0].x} vs shallow ${shallow[0].x} of 4in)`
			)
			assert(
				deep.every((n, i) => Boolean(n.curve) === Boolean(shallow[i].curve)),
				'the presets must share one node structure'
			)
		},
	},
	{
		name: 'addImage({ points: clipPath(...) }) emits a custGeom clip in the picture EMU space',
		fn: async () => {
			const w = 4
			const h = 6
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: PNG_1X1,
					x: 0,
					y: 0,
					w,
					h,
					points: clipPath({ kind: 'half-disc', flat: 'right' }, w, h),
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:custGeom>/.test(xml), 'expected a custGeom clip; got: ' + xml)
			assert(!/<a:prstGeom/.test(xml.split('<p:pic>')[1] ?? ''), 'points must win over the preset geometry')
			assertEqual((xml.match(/<a:cubicBezTo>/g) || []).length, 2, 'expected two cubicBezTo segments')
			// The flat edge must land on the picture's own right edge, in EMU, not on the slide's.
			assert(
				xml.includes(`x="${w * EMU_PER_INCH}"`),
				`expected a point at the box's right edge (${w * EMU_PER_INCH} EMU); got: ` + xml
			)
		},
	},
])
