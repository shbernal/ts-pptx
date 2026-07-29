// Read-model coverage for the custom-geometry segment types + path-attr defaults
// in src/read/api/shapes/geometry.ts (readGeometryPath) that the custgeom.pptx fixture
// doesn't carry: quadBezTo, arcTo (with and without its optional attributes),
// path viewport / fill / stroke defaults, and the documented non-numeric-point
// degrade-to-0. custgeom.test.js covers moveTo / lnTo / cubicBezTo / close on a
// real deck; these hand-authored paths reach the remaining segment branches
// off-fixture through a synthetic AutoShape.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { AutoShape } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** An `AutoShape` over a hand-authored `p:sp` body (geometry reads need no theme). */
function sp(spPrInner) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp>${spPrInner}</p:sp></p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	return new AutoShape(el, /** @type {any} */ ({}))
}

/** The single path of a one-path custGeom built from `pathXml`. */
function onlyPath(pathXml) {
	const geom = sp(`<p:spPr><a:custGeom><a:pathLst>${pathXml}</a:pathLst></a:custGeom></p:spPr>`).customGeometry
	assert(geom, 'expected a custom geometry')
	assertEqual(geom.paths.length, 1, 'one a:path')
	return geom.paths[0]
}

describe('Shape.customGeometry — quadBezTo / arcTo segments', () => {
	test('quadBezTo reads its control + end point; arcTo reads radii and 60000ths angles', () => {
		const path = onlyPath(
			`<a:path w="200" h="100">
				<a:moveTo><a:pt x="0" y="0"/></a:moveTo>
				<a:quadBezTo><a:pt x="10" y="20"/><a:pt x="30" y="40"/></a:quadBezTo>
				<a:arcTo wR="50" hR="25" stAng="0" swAng="5400000"/>
				<a:close/>
			</a:path>`
		)
		assertEqual(path.commands.length, 4, 'moveTo + quadBezTo + arcTo + close')
		assertEqual(
			JSON.stringify(path.commands[1]),
			JSON.stringify({ cmd: 'quadBezTo', x1: 10, y1: 20, x: 30, y: 40 }),
			'quadBezTo control then end point'
		)
		assertEqual(
			JSON.stringify(path.commands[2]),
			JSON.stringify({ cmd: 'arcTo', wR: 50, hR: 25, stAng: 0, swAng: 90 }),
			'arcTo swAng 5400000 (60000ths) → 90°'
		)
		assertEqual(path.commands[3].cmd, 'close', 'trailing close')
	})

	test('an arcTo with no attributes defaults every value to 0', () => {
		const path = onlyPath(`<a:path w="10" h="10"><a:arcTo/></a:path>`)
		assertEqual(
			JSON.stringify(path.commands[0]),
			JSON.stringify({ cmd: 'arcTo', wR: 0, hR: 0, stAng: 0, swAng: 0 }),
			'missing wR/hR/stAng/swAng → 0'
		)
	})
})

describe('Shape.customGeometry — path attribute defaults + degenerate points', () => {
	test('a path with no w/h/fill/stroke uses the schema defaults', () => {
		const path = onlyPath(`<a:path><a:moveTo><a:pt x="1" y="2"/></a:moveTo></a:path>`)
		assertEqual(path.w, 0, 'missing @w → 0')
		assertEqual(path.h, 0, 'missing @h → 0')
		assertEqual(path.fill, 'norm', 'missing @fill → "norm"')
		assertEqual(path.stroke, true, 'missing @stroke → true')
	})

	test('explicit @fill / @stroke override the defaults', () => {
		const path = onlyPath(`<a:path w="5" h="5" fill="darken" stroke="0"><a:close/></a:path>`)
		assertEqual(path.fill, 'darken', 'explicit @fill is surfaced')
		assertEqual(path.stroke, false, 'stroke="0" reads false')
	})

	test('a non-numeric a:pt coordinate degrades to 0 rather than crashing', () => {
		const path = onlyPath(`<a:path w="10" h="10"><a:moveTo><a:pt x="not-a-number" y="also"/></a:moveTo></a:path>`)
		assertEqual(
			JSON.stringify(path.commands[0]),
			JSON.stringify({ cmd: 'moveTo', x: 0, y: 0 }),
			'unparseable coordinates → 0'
		)
	})

	test('an empty pathLst yields a geometry with no paths', () => {
		const geom = sp(`<p:spPr><a:custGeom><a:pathLst/></a:custGeom></p:spPr>`).customGeometry
		assert(geom, 'a custGeom with an empty pathLst still surfaces a geometry')
		assertEqual(geom.paths.length, 0, 'no a:path → empty paths')
	})
})
