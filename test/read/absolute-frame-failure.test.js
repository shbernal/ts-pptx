// Read-model coverage for `Shape.absoluteFrameFailure` (src/read/api/shapes/base.ts):
// the three reasons `absoluteFrame` reports one `null` for. The distinction is not
// recoverable from the `null`, and `src/inspect.ts` used to re-derive it by walking
// the ancestry a second time — these fix the vocabulary that replaced that walk.
//
// Hand-authored shape trees rather than a fixture: two of the three states are
// malformed decks PowerPoint does not author, and the third (an inherited
// placeholder box) needs no theme to read.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { AutoShape } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const XFRM = '<a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>'

/** A `p:sp` carrying `spPrInner`, wrapped in `wrap` (a function nesting it in groups). */
function shapeIn(spPrInner, wrap = (inner) => inner) {
	const sp = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${spPrInner}</p:spPr></p:sp>`
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${wrap(sp)}</p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	return new AutoShape(el, /** @type {any} */ ({}))
}

/** Nest `inner` in a `p:grpSp` whose `p:grpSpPr` holds `grpSpPrInner`. */
const group = (grpSpPrInner) => (inner) =>
	`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${grpSpPrInner}</p:grpSpPr>${inner}</p:grpSp>`

/** A group transform mapping a `chExt` of `cx`×`cy` onto a 1000×1000 slide box. */
const groupXfrm = (cx, cy) =>
	`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm>`

describe('Shape.absoluteFrameFailure', () => {
	test('is null when the frame resolves', () => {
		const shape = shapeIn(XFRM)
		assert(shape.absoluteFrame, 'expected a resolvable frame at slide level')
		assertEqual(shape.absoluteFrameFailure, null, 'a resolved frame has no failure to report')
	})

	test('is null for a group child whose whole chain resolves', () => {
		const shape = shapeIn(XFRM, group(groupXfrm(1000, 1000)))
		assert(shape.absoluteFrame, 'expected a resolvable frame inside a well-formed group')
		assertEqual(shape.absoluteFrameFailure, null, 'a complete group chain reports no failure')
	})

	test('reports no-own-transform when the shape states no a:xfrm', () => {
		// The ordinary case: a placeholder inheriting its box from the layout. Nothing is
		// wrong with the deck, which is why `inspect` stays silent about this one.
		const shape = shapeIn('<a:prstGeom prst="rect"/>')
		assertEqual(shape.absoluteFrame, null, 'no own transform means no frame')
		assertEqual(shape.absoluteFrameFailure, 'no-own-transform', 'an inherited box is named, not blamed')
	})

	test('reports no-own-transform when the shape states an incomplete a:xfrm', () => {
		const shape = shapeIn('<a:xfrm><a:off x="100" y="200"/></a:xfrm>')
		assertEqual(shape.absoluteFrameFailure, 'no-own-transform', 'a half-stated transform is no transform')
	})

	test('reports group-transform-missing when an enclosing group states no a:xfrm', () => {
		const shape = shapeIn(XFRM, group(''))
		assertEqual(shape.absoluteFrame, null, 'an unmapped child space leaves no frame')
		assertEqual(shape.absoluteFrameFailure, 'group-transform-missing', 'the group, not the shape, is at fault')
	})

	test('reports group-transform-missing when an enclosing group states no child space', () => {
		const shape = shapeIn(XFRM, group('<a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>'))
		assertEqual(shape.absoluteFrameFailure, 'group-transform-missing', 'off/ext without chOff/chExt is incomplete')
	})

	test('reports group-transform-degenerate for a zero a:chExt', () => {
		const shape = shapeIn(XFRM, group(groupXfrm(1000, 0)))
		assertEqual(shape.absoluteFrame, null, 'a zero child extent has no ratio to map through')
		assertEqual(shape.absoluteFrameFailure, 'group-transform-degenerate', 'the degenerate axis is named')
	})

	test('a missing group transform outranks a degenerate one in the same chain', () => {
		// Composing needs every group's mapping, so the walk gives up at the ancestor
		// that states none — whichever of the two sits nearer the shape.
		const inner = group(groupXfrm(1000, 0))
		const shape = shapeIn(XFRM, (sp) => group('')(inner(sp)))
		assertEqual(shape.absoluteFrameFailure, 'group-transform-missing', 'missing wins over degenerate')
	})
})
