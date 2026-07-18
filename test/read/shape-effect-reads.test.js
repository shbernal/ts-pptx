// Read-model coverage for the shape EFFECT / STROKE-DETAIL getters in
// src/read/api/shapes.ts that the fixture decks don't happen to carry:
// Shape.shadow (a:effectLst/a:outerShdw), Shape.lineEnds (a:ln head/tail
// arrowheads), and the a:path branch of Shape.gradientFill.
//
// These read only the shape's own spPr (shadow/gradient resolve colour through
// the slide theme, so a minimal themeContext stub is enough), so hand-authored
// OOXML exercises every branch without a round-trip through this library's
// writer — the same off-fixture pattern the style-accessor suite uses.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { AutoShape } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** Empty colour maps resolve `a:srgbClr` literally; no theme parts beyond that. */
function ctx() {
	return {
		clrMap: new Map(),
		clrScheme: new Map(),
		fmtScheme: null,
		fontScheme: null,
		layoutRoot: null,
		masterRoot: null,
	}
}

/** An `AutoShape` over a hand-authored `p:sp` body, resolving against a theme stub. */
function sp(spPrInner) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp>${spPrInner}</p:sp></p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	return new AutoShape(el, /** @type {any} */ ({ themeContext: () => ctx() }))
}

describe('Shape.shadow — outer drop shadow reads', () => {
	test('reads blur / offset / angle (EMU + 60000ths) and an srgb colour with alpha', () => {
		const shape = sp(
			`<p:spPr><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000">` +
				`<a:srgbClr val="808080"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr>`
		)
		const shadow = shape.shadow
		assert(shadow, 'an outerShdw surfaces a shadow')
		assertEqual(shadow.color, '808080', 'effectiveHex of the plain srgb shadow colour')
		assert(Math.abs(shadow.alpha - 0.4) < 1e-9, `alpha 40000 → 0.4, got ${shadow.alpha}`)
		assertEqual(shadow.blurPt, 4, 'blurRad 50800 EMU → 4pt')
		assertEqual(shadow.offsetPt, 3, 'dist 38100 EMU → 3pt')
		assertEqual(shadow.angleDeg, 45, 'dir 2700000 (60000ths) → 45°')
		assertEqual(shadow.colorToken, undefined, 'an srgb shadow carries no scheme colour token')
	})

	test('a scheme-coloured shadow surfaces its colorToken even when unresolved', () => {
		const shadow = sp(
			`<p:spPr><a:effectLst><a:outerShdw><a:schemeClr val="accent1"/></a:outerShdw></a:effectLst></p:spPr>`
		).shadow
		assert(shadow, 'the outerShdw still surfaces a shadow')
		assertEqual(shadow.colorToken, 'accent1', 'the scheme token is reported for a downstream resolver')
		assertEqual(shadow.color, null, 'with empty colour maps the scheme colour does not resolve')
	})

	test('no effectLst / no outerShdw → null', () => {
		assertEqual(sp(`<p:spPr/>`).shadow, null, 'no a:effectLst → null')
		assertEqual(sp(`<p:spPr><a:effectLst/></p:spPr>`).shadow, null, 'an effectLst with no outerShdw → null')
	})
})

describe('Shape.lineEnds — connector arrowheads', () => {
	test('reads both head and tail ends with their type / width / length', () => {
		const ends = sp(
			`<p:spPr><a:ln w="19050"><a:headEnd type="triangle" w="med" len="lg"/><a:tailEnd type="oval"/></a:ln></p:spPr>`
		).lineEnds
		assert(ends, 'a line with arrowheads surfaces lineEnds')
		assertEqual(ends.head.type, 'triangle', 'head type')
		assertEqual(ends.head.width, 'med', 'head width')
		assertEqual(ends.head.length, 'lg', 'head length')
		assertEqual(ends.tail.type, 'oval', 'tail type')
		assertEqual(ends.tail.width, null, 'tail has no explicit width → null')
		assertEqual(ends.tail.length, null, 'tail has no explicit length → null')
	})

	test('a headEnd with no @type defaults its type to none', () => {
		const ends = sp(`<p:spPr><a:ln w="12700"><a:headEnd/></a:ln></p:spPr>`).lineEnds
		assertEqual(ends.head.type, 'none', 'a bare headEnd reads type "none"')
		assertEqual(ends.tail, null, 'no tailEnd → null tail')
	})

	test('a line with neither end, and no line at all, both read null', () => {
		assertEqual(sp(`<p:spPr><a:ln w="12700"/></p:spPr>`).lineEnds, null, 'a plain line has no ends')
		assertEqual(sp(`<p:spPr/>`).lineEnds, null, 'no a:ln → null')
	})
})

describe('Shape.gradientFill — the a:path (radial/rectangular) branch', () => {
	test('a path gradient reports kind "path", its shape, a null angle, and its stops', () => {
		const grad = sp(
			`<p:spPr><a:gradFill><a:gsLst>` +
				`<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>` +
				`<a:gs><a:srgbClr val="0000FF"/></a:gs>` +
				`</a:gsLst><a:path path="circle"/></a:gradFill></p:spPr>`
		).gradientFill
		assert(grad, 'an a:gradFill surfaces a gradientFill')
		assertEqual(grad.kind, 'path', 'an a:path gradient is kind "path"')
		assertEqual(grad.path, 'circle', 'the path shape is surfaced')
		assertEqual(grad.angleDeg, null, 'a path gradient has no linear angle')
		assertEqual(grad.stops.length, 2, 'both stops surfaced')
		assertEqual(grad.stops[0].position, 0, 'first stop at 0%')
		assertEqual(grad.stops[1].position, null, 'a stop with no @pos reports a null position')
	})

	test('a non-gradient fill has no gradientFill', () => {
		assertEqual(
			sp(`<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`).gradientFill,
			null,
			'a solid fill → null'
		)
	})
})
