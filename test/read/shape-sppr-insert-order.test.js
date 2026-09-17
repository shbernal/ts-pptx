// Where a *created* `p:spPr` lands among its siblings.
//
// `getOrAddChild(sp, 'p:spPr', SHAPE_AFTER_SPPR)` inserts before the first sibling named in the
// successor list and appends when it matches none. So the list has to name every child that
// legally follows `p:spPr` — miss one and the new element is appended *after* it, which is a
// schema-invalid `p:sp` and reaches the user as "PowerPoint found a problem with this file",
// far from its cause and with no compile-time signal.
//
// `SHAPE_AFTER_SPPR` was the one successor list in `src/ooxml/sequence.ts` written out by hand
// rather than sliced from a declared sequence, and it was missing `p:extLst` — the last child
// of `CT_Shape` (ECMA-376 Part 1 §19.3.1.43: `p:nvSpPr`, `p:spPr`, `p:style`, `p:txBody`,
// `p:extLst`). Deriving it closed the gap in the same edit that removed the exception.
//
// A `p:sp` with no `p:spPr` is itself schema-invalid — the schema makes it required, so no deck
// PowerPoint wrote has one, and the add branch never fires on a well-formed input. That is
// precisely why it needs a test rather than a fixture: the read side is fed whatever it is
// given, and the branch that repairs a malformed shape must not replace one invalidity with
// another.

import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { AutoShape, GroupShape } from '../../dist/read.js'
import { assert } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** Empty colour maps — nothing here resolves a colour, it only places elements. */
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

/** An `AutoShape` over a hand-authored `p:sp` body, plus a serializer for the result. */
function sp(body) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp>${body}</p:sp></p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	// A host stub: the setters mark the owning part dirty on the way out, and nothing here
	// has a part behind it.
	const host = { themeContext: () => ctx(), part: { markDirty: () => {} } }
	const shape = new AutoShape(el, /** @type {any} */ (host))
	return { shape, xml: () => new XMLSerializer().serializeToString(el) }
}

/** The order of `p:sp`'s direct element children, as qnames. */
function childOrder(xml) {
	return [...xml.matchAll(/<(p:[a-zA-Z0-9]+)[\s/>]/g)].map((m) => m[1]).filter((n) => n !== 'p:sp')
}

const NV = '<p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'

describe('a created p:spPr lands in schema order', () => {
	test('before p:extLst, the child the hand-written successor list used to omit', () => {
		const { shape, xml } = sp(`${NV}<p:extLst><p:ext uri="{X}"/></p:extLst>`)
		shape.noFill()
		const order = childOrder(xml())
		assert(order.indexOf('p:spPr') < order.indexOf('p:extLst'), 'p:spPr must precede p:extLst; got ' + order.join(', '))
	})

	test('before p:style and p:txBody, which it already did', () => {
		const { shape, xml } = sp(`${NV}<p:style/><p:txBody><a:bodyPr/><a:p/></p:txBody>`)
		shape.noFill()
		const order = childOrder(xml())
		assert(order.indexOf('p:spPr') < order.indexOf('p:style'), 'p:spPr must precede p:style; got ' + order.join(', '))
	})

	test('and after p:nvSpPr, which is a predecessor rather than a successor', () => {
		// The list names successors only. `p:nvSpPr` is absent from it not by omission but
		// because naming it would insert the new `p:spPr` *before* the shape's identity.
		const { shape, xml } = sp(`${NV}<p:txBody><a:bodyPr/><a:p/></p:txBody>`)
		shape.noFill()
		const order = childOrder(xml())
		assert(order.indexOf('p:nvSpPr') < order.indexOf('p:spPr'), 'p:nvSpPr must stay first; got ' + order.join(', '))
	})
})

/** A `GroupShape` over a hand-authored `p:grpSp` body, plus a serializer for the result. */
function grpSp(body) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:grpSp>${body}</p:grpSp></p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'grpSp')[0]
	const host = { themeContext: () => ctx(), part: { markDirty: () => {} } }
	const shape = new GroupShape(el, /** @type {any} */ (host))
	return { shape, xml: () => new XMLSerializer().serializeToString(el) }
}

/** The order of `p:grpSp`'s direct element children, as qnames. */
function groupChildOrder(xml) {
	return [...xml.matchAll(/<(p:[a-zA-Z0-9]+)[\s/>]/g)].map((m) => m[1]).filter((n) => n !== 'p:grpSp')
}

const NV_GRP = '<p:nvGrpSpPr><p:cNvPr id="2" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'

// Same failure mode as above, one type along. `CT_GroupShape` (§19.3.1.22) lets `p:contentPart`
// and `p:extLst` follow `p:grpSpPr`, and the hand-written list named neither -- so a group whose
// only children were ink content parts or an extension list got its `p:grpSpPr` appended after
// them. `p:grpSpPr` is required, so only a malformed source reaches the add branch at all, which
// is exactly why this is a hand-authored DOM and not a PowerPoint fixture: there is no deck
// PowerPoint would write that exercises it.
describe('a created p:grpSpPr lands in schema order', () => {
	test('before p:contentPart, which the hand-written successor list omitted', () => {
		const { shape, xml } = grpSp(`${NV_GRP}<p:contentPart r:id="rId9" xmlns:r="${R_NS}"/>`)
		shape.noFill()
		const order = groupChildOrder(xml())
		assert(
			order.indexOf('p:grpSpPr') < order.indexOf('p:contentPart'),
			'p:grpSpPr must precede p:contentPart; got ' + order.join(', ')
		)
	})

	test('before p:extLst, which it also omitted', () => {
		const { shape, xml } = grpSp(`${NV_GRP}<p:extLst><p:ext uri="{X}"/></p:extLst>`)
		shape.noFill()
		const order = groupChildOrder(xml())
		assert(
			order.indexOf('p:grpSpPr') < order.indexOf('p:extLst'),
			'p:grpSpPr must precede p:extLst; got ' + order.join(', ')
		)
	})

	test('before the drawable children, which it already did', () => {
		const { shape, xml } = grpSp(`${NV_GRP}<p:pic><p:nvPicPr/></p:pic>`)
		shape.noFill()
		const order = groupChildOrder(xml())
		assert(order.indexOf('p:grpSpPr') < order.indexOf('p:pic'), 'p:grpSpPr must precede p:pic; got ' + order.join(', '))
	})

	test('and after p:nvGrpSpPr, which is a predecessor rather than a successor', () => {
		const { shape, xml } = grpSp(`${NV_GRP}<p:pic><p:nvPicPr/></p:pic>`)
		shape.noFill()
		const order = groupChildOrder(xml())
		assert(
			order.indexOf('p:nvGrpSpPr') < order.indexOf('p:grpSpPr'),
			'p:nvGrpSpPr must stay first; got ' + order.join(', ')
		)
	})
})
