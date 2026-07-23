// Write→read fidelity for Connector endpoint binding (src/read/api/shapes.ts).
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
			slide.addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'BoxA' })
			slide.addShape(pres.ShapeType.rect, { x: 5, y: 3, w: 2, h: 1, objectName: 'BoxB' })
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
			slide.addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'Only' })
			slide.addConnector({ type: 'straight', x1: 3, y1: 1.5, x2: 5, y2: 3, startShape: 'Only' })
		})
		const cxn = connectorOf(presentation)
		const start = cxn.startConnection
		assert(start, 'the bound start end reads back')
		assertEqual(start.siteIndex, 0, 'omitted startShapeIdx defaults to 0')
		assertEqual(start.boundShape?.name, 'Only', 'start boundShape resolves')
		assertEqual(cxn.endConnection, null, 'the unbound end stays null')
	})

	test.skipIf(!validatorInstalled)('the authored bound-connector deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, objectName: 'BoxA' })
			slide.addShape(pres.ShapeType.rect, { x: 5, y: 3, w: 2, h: 1, objectName: 'BoxB' })
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
})
