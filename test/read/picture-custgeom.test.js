// Read-model coverage for Shape.customGeometry on a picture. The oracle is picture-custgeom.pptx,
// authored in desktop PowerPoint (see test/read/fixtures/README.md) by Merge Shapes → Intersect of a
// picture and a freeform: PowerPoint keeps the result a `p:pic` and writes the freeform into its
// `p:spPr` as `a:custGeom`. Every number asserted below is a literal copied out of that fixture's
// slide1.xml.
//
// Pictures on the single slide:
// - pic-clip-lines: clipped to a triangle (moveTo + 2×lnTo + close).
// - pic-clip-curve: clipped to an arch. PowerPoint re-expresses the intersected outline starting
//   mid-curve, with a control point above the frame (negative y).
// - pic-preset-oval: cropped to a preset (`AutoShapeType`), the prstGeom negative control.
// - pic-plain: an untouched picture, `prstGeom rect`.

import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

async function pictureNamed(name) {
	const slide = (await openFixture('picture-custgeom')).slides[0]
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected shape named ${name}`)
	assertEqual(shape.shapeType, 'picture', `${name} is a picture`)
	return shape
}

describe('Shape.customGeometry on a picture: real PowerPoint XML (picture-custgeom.pptx)', () => {
	test('pic-clip-lines: a picture clipped to a triangle reads its clip path', async () => {
		const pic = await pictureNamed('pic-clip-lines')
		assertEqual(pic.presetGeometry, null, 'a freeform clip has no preset')
		assertEqual(
			JSON.stringify(pic.customGeometry),
			JSON.stringify({
				paths: [
					{
						w: 2794000,
						h: 2222500,
						fill: 'norm',
						stroke: true,
						commands: [
							{ cmd: 'moveTo', x: 0, y: 0 },
							{ cmd: 'lnTo', x: 2794000, y: 0 },
							{ cmd: 'lnTo', x: 1397000, y: 2222500 },
							{ cmd: 'close' },
						],
					},
				],
			}),
			'clip path'
		)
	})

	test('pic-clip-curve: cubic control points read in order, including one outside the frame', async () => {
		const pic = await pictureNamed('pic-clip-curve')
		const geom = pic.customGeometry
		assert(geom, 'pic-clip-curve has custom geometry')
		assertEqual(geom.paths.length, 1, 'one a:path')
		assertEqual(geom.paths[0].w, 2794000, 'path w')
		assertEqual(geom.paths[0].h, 2066402, 'path h')
		assertEqual(
			JSON.stringify(geom.paths[0].commands),
			JSON.stringify([
				{ cmd: 'moveTo', x: 1492250, y: 2652 },
				{ cmd: 'cubicBezTo', x1: 2095500, y1: 34402, x2: 2667000, y2: 351902, x: 2794000, y: 986902 },
				{ cmd: 'lnTo', x: 2794000, y: 2066402 },
				{ cmd: 'lnTo', x: 0, y: 2066402 },
				{ cmd: 'lnTo', x: 0, y: 732902 },
				{ cmd: 'cubicBezTo', x1: 254000, y1: 224902, x2: 889000, y2: -29098, x: 1492250, y: 2652 },
				{ cmd: 'close' },
			]),
			'ordered commands'
		)
	})

	test('pic-preset-oval and pic-plain: a preset-geometry picture reports customGeometry === null', async () => {
		const oval = await pictureNamed('pic-preset-oval')
		assertEqual(oval.customGeometry, null, 'a preset crop is not custom geometry')
		assertEqual(oval.presetGeometry, 'ellipse', 'the crop still reads as its preset')
		const plain = await pictureNamed('pic-plain')
		assertEqual(plain.customGeometry, null, 'a plain picture has no custom geometry')
		assertEqual(plain.presetGeometry, 'rect', 'a plain picture is a rect')
	})
})
