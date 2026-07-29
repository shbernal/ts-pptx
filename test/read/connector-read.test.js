// Write→read fidelity for Connector endpoint binding (src/read/api/shapes/connector.ts).
//
// A connector authored with `startShape`/`endShape` emits
// `<p:cNvCxnSpPr><a:stCxn id idx/><a:endCxn id idx/></p:cNvCxnSpPr>`, binding each
// end to a shape by that shape's `p:cNvPr/@id` + connection-site index. The read
// side previously surfaced such a connector as geometry-only, dropping the
// attachment. `Connector.startConnection` / `.endConnection` now decode it and
// resolve `@id` back to the slide shape via `slide.shapeById`.
//
// The binding relies on the writer resolving each target's `objectName` → id at
// serialize time, so the fixture gives both target shapes an `objectName`. A
// connector authored with no binding emits a bare `<p:cNvCxnSpPr/>`; both getters
// must report null for it (faithful "unbound" ≠ a half-populated site).
//
// T1.3 extends this: a connector may bind to a shape nested inside a group (the
// writer already resolves a group child's `objectName` → id). `boundShape` now
// resolves that through `slide.shapeByIdDeep`, which descends into groups, rather
// than degrading to null the way the old top-level-only `slide.shapeById` did.

import { ShapeType } from '../../dist/node.js'
import { describe, test } from 'vitest'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** The single connector on any slide of `presentation`. */
function connectorOf(presentation) {
	const cxn = firstShape(presentation, (s) => s.shapeType === 'connector')
	assert(cxn, 'the authored connector is read back')
	return cxn
}

/** The rect autoShape named `name`. */
function rectNamed(presentation, name) {
	const rect = firstShape(presentation, (s) => s.shapeType === 'autoShape' && s.name === name)
	assert(rect, `the authored rect "${name}" is read back`)
	return rect
}

describe('Connector.startConnection / endConnection — write→read fidelity', () => {
	test('a bound connector resolves both ends to their shapes and site indexes', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'BoxA' })
			slide.addShape(ShapeType.rect, { x: 5, y: 3, w: 2, h: 1, objectName: 'BoxB' })
			slide.addConnector({
				type: 'straight',
				x1: 3,
				y1: 1.5,
				x2: 5,
				y2: 3.5,
				startShape: 'BoxA',
				startShapeIdx: 1,
				endShape: 'BoxB',
				endShapeIdx: 3,
			})
		})

		const boxA = rectNamed(presentation, 'BoxA')
		const boxB = rectNamed(presentation, 'BoxB')
		const cxn = connectorOf(presentation)

		const start = cxn.startConnection
		assert(start, 'the start end is bound')
		assertEqual(start.siteIndex, 1, 'start site index round-trips')
		assertEqual(start.shapeId, boxA.id, 'start binds BoxA by its cNvPr id')
		assert(start.boundShape, 'start resolves to a shape')
		assertEqual(start.boundShape.name, 'BoxA', 'start boundShape is BoxA')

		const end = cxn.endConnection
		assert(end, 'the end end is bound')
		assertEqual(end.siteIndex, 3, 'end site index round-trips')
		assertEqual(end.shapeId, boxB.id, 'end binds BoxB by its cNvPr id')
		assert(end.boundShape, 'end resolves to a shape')
		assertEqual(end.boundShape.name, 'BoxB', 'end boundShape is BoxB')
	})

	test('an unbound connector (bare p:cNvCxnSpPr) reports null for both ends', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addConnector({ type: 'straight', x1: 1, y1: 1, x2: 4, y2: 3 })
		})
		const cxn = connectorOf(presentation)
		assertEqual(cxn.startConnection, null, 'no start binding → null')
		assertEqual(cxn.endConnection, null, 'no end binding → null')
	})

	test('binding only one end leaves the other null', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'Only' })
			slide.addConnector({ type: 'straight', x1: 3, y1: 1.5, x2: 5, y2: 3, startShape: 'Only' })
		})
		const cxn = connectorOf(presentation)
		const start = cxn.startConnection
		assert(start, 'the bound start end reads back')
		assertEqual(start.siteIndex, 0, 'omitted startShapeIdx defaults to 0')
		assertEqual(start.boundShape?.name, 'Only', 'start boundShape resolves')
		assertEqual(cxn.endConnection, null, 'the unbound end stays null')
	})

	test('a connector bound to a group-nested shape resolves through the group (T1.3)', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			// The rect lives inside a group, so it is not a top-level shape — the very
			// case the old top-level-only resolver reported as boundShape: null.
			slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: 'Nested' } }], { objectName: 'Grp' })
			slide.addConnector({ type: 'straight', x1: 3, y1: 1.5, x2: 5, y2: 3, startShape: 'Nested', startShapeIdx: 1 })
		})
		const cxn = connectorOf(presentation)
		const start = cxn.startConnection
		assert(start, 'the start end binds the group-nested shape')
		assertEqual(start.siteIndex, 1, 'site index round-trips for a group-nested binding')
		// The `@id` is authored either way; deep resolution is what makes boundShape non-null.
		assert(start.boundShape, 'the group-nested boundShape resolves (was null before deep resolution)')
		assertEqual(start.boundShape.name, 'Nested', 'boundShape is the group-nested rect')
		assertEqual(start.boundShape.id, start.shapeId, 'the resolved shape carries the bound id')
	})

	test.skipIf(!validatorInstalled)('the authored bound-connector deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addShape(ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'BoxA' })
			slide.addShape(ShapeType.rect, { x: 5, y: 3, w: 2, h: 1, objectName: 'BoxB' })
			slide.addConnector({
				type: 'straight',
				x1: 3,
				y1: 1.5,
				x2: 5,
				y2: 3.5,
				startShape: 'BoxA',
				startShapeIdx: 1,
				endShape: 'BoxB',
				endShapeIdx: 3,
			})
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'bound-connector deck validates')
	})

	test.skipIf(!validatorInstalled)('the authored group-nested-binding deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: 'Nested' } }], { objectName: 'Grp' })
			slide.addConnector({ type: 'straight', x1: 3, y1: 1.5, x2: 5, y2: 3, startShape: 'Nested', startShapeIdx: 1 })
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'group-nested-binding deck validates')
	})
})
