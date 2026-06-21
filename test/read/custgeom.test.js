// Read-model coverage for Shape.customGeometry — the `spPr/a:custGeom/a:pathLst`
// freeform path accessor. The oracle is custgeom.pptx, a minimal deck authored in
// desktop PowerPoint (see test/read/fixtures/README.md): a custom-geometry read
// can only be trusted against genuine Office XML, so every number asserted below
// is a literal copied out of that fixture's slide1.xml — never synthesized.
//
// Shapes on the single slide:
// - freeform-lines: closed triangle (moveTo + 2×lnTo + close) — the common case.
// - freeform-cubic: a cubicBezTo curve then two lnTo, left open — control-point order.
// - freeform-hole: rect-with-elliptical-hole from Merge Shapes → Subtract; PowerPoint
//   emits it as ONE a:path with two moveTo…close contours (ellipse = 4 cubics, then
//   the outer rect = 3 lnTo) — pins multi-contour single-path traversal.
// - preset-rect: <a:prstGeom prst="rect"> — the negative case (customGeometry === null).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function openCustgeom() {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', 'custgeom.pptx')))
}

function shapeNamed(slide, name) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected shape named ${name}`)
	return shape
}

/** Assert a path's viewport attrs and ordered commands against recorded literals. */
function assertPath(actual, expected) {
	assertEqual(actual.w, expected.w, 'path w')
	assertEqual(actual.h, expected.h, 'path h')
	assertEqual(actual.fill, expected.fill, 'path fill')
	assertEqual(actual.stroke, expected.stroke, 'path stroke')
	assertEqual(JSON.stringify(actual.commands), JSON.stringify(expected.commands), 'path commands')
}

describe('Shape.customGeometry — real PowerPoint XML (custgeom.pptx)', () => {
	test('freeform-lines: closed triangle from moveTo + lnTo + close', async () => {
		const slide = (await openCustgeom()).slides[0]
		const geom = shapeNamed(slide, 'freeform-lines').customGeometry
		assert(geom, 'freeform-lines has custom geometry')
		assertEqual(geom.paths.length, 1, 'one a:path')
		assertPath(geom.paths[0], {
			w: 2540001,
			h: 2540001,
			fill: 'norm', // no @fill → schema default
			stroke: true, // no @stroke → schema default
			commands: [
				{ cmd: 'moveTo', x: 0, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 0 },
				{ cmd: 'lnTo', x: 1270000, y: 2540000 },
				{ cmd: 'close' },
			],
		})
	})

	test('freeform-cubic: cubicBezTo control points read in c1,c2,end order', async () => {
		const slide = (await openCustgeom()).slides[0]
		const geom = shapeNamed(slide, 'freeform-cubic').customGeometry
		assert(geom, 'freeform-cubic has custom geometry')
		assertEqual(geom.paths.length, 1, 'one a:path')
		assertPath(geom.paths[0], {
			w: 2540001,
			h: 3302001,
			fill: 'norm',
			stroke: true,
			commands: [
				{ cmd: 'moveTo', x: 0, y: 762000 },
				{ cmd: 'cubicBezTo', x1: 635000, y1: 0, x2: 2032000, y2: 0, x: 2540000, y: 1270000 },
				{ cmd: 'lnTo', x: 2540000, y: 3302000 },
				{ cmd: 'lnTo', x: 0, y: 3302000 },
			],
		})
	})

	test('freeform-hole: one path carries two moveTo…close contours in document order', async () => {
		const slide = (await openCustgeom()).slides[0]
		const geom = shapeNamed(slide, 'freeform-hole').customGeometry
		assert(geom, 'freeform-hole has custom geometry')
		// PowerPoint's Merge Shapes → Subtract emits a single a:path, NOT two.
		assertEqual(geom.paths.length, 1, 'a hole is one a:path with multiple contours')
		const { commands } = geom.paths[0]
		assertEqual(geom.paths[0].w, 2540000, 'path w')
		assertEqual(geom.paths[0].h, 1524000, 'path h')
		// Two moveTo / two close: the inner ellipse contour then the outer rectangle.
		assertEqual(commands.filter((c) => c.cmd === 'moveTo').length, 2, 'two contours start with moveTo')
		assertEqual(commands.filter((c) => c.cmd === 'close').length, 2, 'each contour closes')
		assertEqual(
			JSON.stringify(commands),
			JSON.stringify([
				// Contour 1 — elliptical hole, four cubic Béziers.
				{ cmd: 'moveTo', x: 1270000, y: 381000 },
				{ cmd: 'cubicBezTo', x1: 919299, y1: 381000, x2: 635000, y2: 551580, x: 635000, y: 762000 },
				{ cmd: 'cubicBezTo', x1: 635000, y1: 972420, x2: 919299, y2: 1143000, x: 1270000, y: 1143000 },
				{ cmd: 'cubicBezTo', x1: 1620701, y1: 1143000, x2: 1905000, y2: 972420, x: 1905000, y: 762000 },
				{ cmd: 'cubicBezTo', x1: 1905000, y1: 551580, x2: 1620701, y2: 381000, x: 1270000, y: 381000 },
				{ cmd: 'close' },
				// Contour 2 — outer rectangle, three lnTo.
				{ cmd: 'moveTo', x: 0, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 1524000 },
				{ cmd: 'lnTo', x: 0, y: 1524000 },
				{ cmd: 'close' },
			]),
			'ordered multi-contour commands'
		)
	})

	test('preset-rect: a preset-geometry shape reports customGeometry === null', async () => {
		const slide = (await openCustgeom()).slides[0]
		const rect = shapeNamed(slide, 'preset-rect')
		assertEqual(rect.customGeometry, null, 'preset geometry → no custom geometry')
		assertEqual(rect.presetGeometry, 'rect', 'preset-rect still reports its preset name')
	})
})
