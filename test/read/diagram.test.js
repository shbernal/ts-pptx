// The SmartArt (`dgm:`) reader added in src/read/api/diagram.ts.
//
// The oracle is `mixed.pptx` slide 2, a PowerPoint-authored `hList1` diagram whose data
// model carries eleven nodes with real run-split text. It is genuine Office output, which
// this reader needs: nothing in the library writes SmartArt, so a write→read round trip
// could only prove the reader agrees with a fixture the reader's own author invented.
//
// The gap this closes: a `p:graphicFrame` holding a diagram answered `false` to all three
// host predicates and `null` to all three accessors, so a slide whose whole message was a
// SmartArt graphic flattened to the empty string with no signal that anything was lost.
// `Slide.text` is therefore asserted here too, not only the new getters.

import { describe, test } from 'vitest'
import { OpcPackage, Presentation, isGraphicFrame } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture, readFixture } from './corpus.js'

const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'

/** The SmartArt frame on `mixed.pptx` slide 2 (index 1). */
async function smartArtFrame() {
	const presentation = await openFixture('mixed')
	const slide = presentation.slides[1]
	const frame = slide.shapes.find(isGraphicFrame)
	assert(frame, 'slide 2 of mixed.pptx has a graphic frame')
	return { presentation, slide, frame }
}

/** The eleven node strings PowerPoint stored, in `dgm:ptLst` order. */
const NODE_TEXT = [
	'Understand Data Modelling Principles',
	'Uncontrollable Inputs (environmental factors)',
	'Mathematical model',
	'Be able to create complex models using spreadsheets',
	'Create models, simulate changes of uncontrollable inputs, see the impact on results',
	'Graph the results',
	'Be able to use Pivot Tables',
	'Use probabilities in models',
	'Simulate uncontrollable inputs using probability distribution',
	'Conduct simulations',
	'Controllable Inputs (decision variables)',
]

describe('GraphicFrame — diagram host', () => {
	test('a SmartArt frame reports hasDiagram and its graphicDataUri', async () => {
		const { frame } = await smartArtFrame()
		assertEqual(frame.hasDiagram, true, 'the frame reports a diagram host')
		assertEqual(frame.hasTable, false, 'and is not a table')
		assertEqual(frame.hasChart, false, 'and is not a classic chart')
		assertEqual(frame.hasChartEx, false, 'and is not a chartEx chart')
		assertEqual(frame.graphicDataUri, DIAGRAM_URI, 'the raw a:graphicData/@uri is the diagram namespace')
		assert(frame.table === null && frame.chart === null && frame.chartEx === null, 'the other accessors stay null')
		assert(frame.diagram !== null, 'the diagram accessor resolves the dgm: data part')
	})

	test('a table or chart frame reports no diagram', async () => {
		const slides = (await openFixture('mixed')).slides
		const table = slides[6].shapes.find(isGraphicFrame) // slide7: a:tbl
		const chart = slides[7].shapes.find(isGraphicFrame) // slide8: c:chart
		for (const label of ['table', 'chart']) {
			const frame = label === 'table' ? table : chart
			assert(frame, `slide for the ${label} frame resolves`)
			assertEqual(frame.hasDiagram, false, `the ${label} frame is not a diagram`)
			assert(frame.diagram === null, `the ${label} frame's diagram accessor is null`)
			assert(frame.graphicDataUri !== DIAGRAM_URI, `the ${label} frame carries a different graphicData uri`)
		}
	})

	test('a diagram frame with no dgm:relIds reads as null rather than throwing', async () => {
		// The other half of "the predicate is about the frame": a frame can announce the
		// diagram uri and carry no reference at all. It must still say what it is.
		const { frame } = await smartArtFrame()
		const relIds = frame.element_.getElementsByTagName('dgm:relIds')[0]
		relIds.parentNode.removeChild(relIds)
		assertEqual(frame.hasDiagram, true, 'the uri still names a diagram')
		assert(frame.diagram === null, 'but there is no data part to reach')
	})

	test('a diagram frame whose data part is absent reads as null rather than throwing', async () => {
		// The predicate is about the frame; the accessor is about the package. A deck that
		// lost its data part must still say "this is a diagram" so a consumer can warn.
		const pkg = await OpcPackage.load(await readFixture('mixed'))
		assertEqual(pkg.removePart('/ppt/diagrams/data1.xml'), true, 'the data part was there to remove')
		const presentation = await Presentation.load(await pkg.save())
		const frame = presentation.slides[1].shapes.find(isGraphicFrame)
		assertEqual(frame.hasDiagram, true, 'the frame still reports a diagram host')
		assert(frame.diagram === null, 'but the accessor is null with no data part to resolve')
	})
})

describe('Diagram — the data model', () => {
	test('exposes every point in document order, typed', async () => {
		const { frame } = await smartArtFrame()
		const points = frame.diagram.points
		const byType = {}
		for (const point of points) byType[point.type] = (byType[point.type] ?? 0) + 1
		// A saved hList1 with eleven nodes: one doc root, a parTrans/sibTrans pair per edge,
		// and the layout engine's own presentation points.
		assertEqual(
			JSON.stringify(byType),
			JSON.stringify({ doc: 1, node: 11, parTrans: 11, sibTrans: 11, pres: 12 }),
			'point types and their counts'
		)
		assert(
			points.every((point) => /^\{[0-9A-F-]+\}$/.test(point.modelId ?? '')),
			'every point carries an ST_ModelId GUID'
		)
	})

	test('node text extracts in document order and joins as Diagram.text', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		assertEqual(
			diagram.points
				.filter((point) => point.type === 'node')
				.map((point) => point.text)
				.join('|'),
			NODE_TEXT.join('|'),
			'the node points carry the authored text'
		)
		assertEqual(diagram.text, NODE_TEXT.join('\n'), 'Diagram.text is those blocks joined by newlines')
	})

	test('text skips the generated pres points, the doc root, and empty transitions', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		// Every excluded class is present in this fixture, so the filter is actually exercised:
		// were any of them folded in, `text` would gain blank lines or duplicate content.
		assert(
			diagram.points.some((point) => point.type === 'pres'),
			'the fixture has pres points to exclude'
		)
		const doc = diagram.points.find((point) => point.type === 'doc')
		assertEqual(doc.isPlaceholder, true, 'the doc root is flagged as an unfilled placeholder')
		assertEqual(doc.text, '', 'and carries no text')
		assert(!diagram.text.includes('\n\n'), 'no empty block reaches the joined text')
		assertEqual(diagram.text.split('\n').length, NODE_TEXT.length, 'exactly the node blocks are joined')
	})

	test('a point reads through the ordinary TextFrame run model', async () => {
		const { frame } = await smartArtFrame()
		const point = frame.diagram.points.find((candidate) => candidate.text === NODE_TEXT[0])
		assert(point, 'the first node point is located by its text')
		const paragraphs = point.textFrame.paragraphs
		assertEqual(paragraphs.length, 1, 'the node holds one paragraph')
		// PowerPoint split this node across five runs on spell-check boundaries (`@err`).
		assertEqual(
			paragraphs[0].runs.map((run) => run.text).join(''),
			NODE_TEXT[0],
			'the runs concatenate to the node text'
		)
		assert(paragraphs[0].runs.length > 1, 'and the split survives rather than flattening to one run')
	})

	test('connections give the points their tree', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		const ids = new Set(diagram.points.map((point) => point.modelId))
		const parOf = diagram.connections.filter((connection) => connection.type === 'parOf')
		assert(parOf.length > 0, 'the data model has parent/child edges')
		for (const connection of parOf) {
			assert(ids.has(connection.sourceId), `parOf srcId ${connection.sourceId} names a point`)
			assert(ids.has(connection.destinationId), `parOf destId ${connection.destinationId} names a point`)
			assert(typeof connection.sourceOrder === 'number', 'srcOrd decodes as a number')
		}
		// The root's children are the three top-level nodes, ordered by srcOrd.
		const root = diagram.points.find((point) => point.type === 'doc').modelId
		const children = parOf
			.filter((connection) => connection.sourceId === root)
			.sort((a, b) => a.sourceOrder - b.sourceOrder)
			.map((connection) => diagram.points.find((point) => point.modelId === connection.destinationId).text)
		assertEqual(
			children.join('|'),
			[NODE_TEXT[0], NODE_TEXT[3], NODE_TEXT[7]].join('|'),
			'the root has three children in srcOrd order'
		)
		// `presOf` edges are the layout engine's bookkeeping, and must not be mistaken for structure.
		assert(
			diagram.connections.some((connection) => connection.type === 'presOf'),
			'the generated presOf edges are surfaced too, typed apart from the tree'
		)
	})

	test('a transition point names the edge it labels; a content point does not', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		const edges = new Set(diagram.connections.map((connection) => connection.modelId))
		const transitions = diagram.points.filter((point) => point.type === 'parTrans' || point.type === 'sibTrans')
		assert(transitions.length > 0, 'the fixture has transition points')
		for (const point of transitions) {
			assert(edges.has(point.connectionId), `@cxnId ${point.connectionId} names a connection`)
		}
		// `@cxnId` defaults to the literal 0, which means "none" rather than naming a
		// point, so a node must read null rather than the string '0'.
		for (const point of diagram.points.filter((candidate) => candidate.type === 'node')) {
			assert(point.connectionId === null, 'a content node labels no edge')
		}
	})

	test('an empty data model reads as no points, no connections, and no text', async () => {
		// The floor under every list getter above: they must return [] rather than throw
		// when the element they read is absent. Reached by emptying the model in place,
		// since no fixture ships a diagram with nothing in it.
		const { presentation, frame } = await smartArtFrame()
		const diagram = frame.diagram
		const root = diagram.element_
		for (const name of ['dgm:ptLst', 'dgm:cxnLst', 'dgm:extLst']) {
			const child = root.getElementsByTagName(name)[0]
			child.parentNode.removeChild(child)
		}
		diagram.markDirty()
		const reopened = await Presentation.load(await presentation.save())
		const emptied = reopened.slides[1].shapes.find(isGraphicFrame).diagram
		assertEqual(emptied.points.length, 0, 'no points')
		assertEqual(emptied.connections.length, 0, 'no connections')
		assertEqual(emptied.text, '', 'no text')
		assert(emptied.layoutTypeId === null, 'no doc point, so no layout preset id')
		assert(emptied.drawingPart === null, 'no extension, so no fallback drawing')
	})

	test('resolves the four sidecar parts and the fallback drawing', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		assertEqual(diagram.partName, '/ppt/diagrams/data1.xml', 'the data part')
		assertEqual(diagram.layoutPart.partName, '/ppt/diagrams/layout1.xml', 'the layout part (r:lo)')
		assertEqual(diagram.quickStylePart.partName, '/ppt/diagrams/quickStyle1.xml', 'the quick-style part (r:qs)')
		assertEqual(diagram.colorsPart.partName, '/ppt/diagrams/colors1.xml', 'the colours part (r:cs)')
		// Named by the MS extension on the data model, not by dgm:relIds.
		assertEqual(diagram.drawingPart.partName, '/ppt/diagrams/drawing1.xml', 'the fallback drawing part')
	})

	test('the drawing lookup skips extensions that are not the diagram one', async () => {
		// `dgm:extLst` is a list, and a data model can carry extensions this reader knows
		// nothing about. Matching the first `a:ext` rather than the one whose `@uri` is the
		// diagram extension would resolve the wrong relationship id, or none.
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		const extLst = diagram.element_.getElementsByTagName('dgm:extLst')[0]
		const real = extLst.getElementsByTagName('a:ext')[0]
		const decoy = extLst.ownerDocument.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:ext')
		decoy.setAttribute('uri', '{00000000-0000-0000-0000-000000000000}')
		extLst.insertBefore(decoy, real)
		assertEqual(diagram.drawingPart.partName, '/ppt/diagrams/drawing1.xml', 'the decoy is skipped')

		extLst.removeChild(real)
		assert(diagram.drawingPart === null, 'with only the decoy left, no drawing resolves')
	})

	test('names the SmartArt kind through the doc point preset ids', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		assertEqual(
			diagram.layoutTypeId,
			'urn:microsoft.com/office/officeart/2005/8/layout/hList1',
			'the layout preset names the SmartArt kind'
		)
		assertEqual(
			diagram.quickStyleTypeId,
			'urn:microsoft.com/office/officeart/2005/8/quickstyle/3d4',
			'the quick-style preset'
		)
		assertEqual(diagram.colorsTypeId, 'urn:microsoft.com/office/officeart/2005/8/colors/accent0_3', 'the colour preset')
	})

	test('a node text edit marks the data part, not the slide, and survives a save', async () => {
		// The wiring this pins: `DiagramPoint` holds the diagram data part. Threading the
		// slide's part instead would mark the wrong part dirty, and the edit would vanish on
		// save while every in-memory read kept reporting it.
		const { presentation, frame } = await smartArtFrame()
		const point = frame.diagram.points.find((candidate) => candidate.text === NODE_TEXT[0])
		point.textFrame.text = 'Rewritten node'
		const reopened = await Presentation.load(await presentation.save())
		const diagram = reopened.slides[1].shapes.find(isGraphicFrame).diagram
		assert(diagram.text.startsWith('Rewritten node\n'), 'the edit is in the saved data part')
		assert(!diagram.text.includes(NODE_TEXT[0]), 'and the original node text is gone')
	})
})

describe('Slide.text — diagram contribution', () => {
	test('a SmartArt slide no longer flattens to its title alone', async () => {
		const { slide } = await smartArtFrame()
		const text = slide.text
		assert(text.startsWith('Course objectives\n'), 'the title placeholder still leads')
		for (const node of NODE_TEXT) assert(text.includes(node), `slide text includes the node "${node}"`)
		assertEqual(
			text,
			['Course objectives', ...NODE_TEXT].join('\n'),
			'the diagram contributes one block per node, in shape-tree order'
		)
	})
})
