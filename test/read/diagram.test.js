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
import { assert, assertEqual, assertUnchangedExcept, captureDiagnostics, partBodies } from '../helpers.js'
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

/**
 * One diagram of `smartart-families.pptx`, the four-family fixture: slide 1 `orgChart1`
 * (multi-level, with an `asst`), 2 `process1` (labelled arrows), 3 `cycle2`, 4 `pList1`.
 *
 * @param {number} slideNumber 1-based
 */
async function familyDiagram(slideNumber) {
	const presentation = await openFixture('smartart-families')
	const frame = presentation.slides[slideNumber - 1].shapes.find(isGraphicFrame)
	assert(frame?.diagram, `slide ${slideNumber} of smartart-families.pptx holds a diagram`)
	return { presentation, diagram: frame.diagram }
}

/** A node tree as `level:text` lines, so a whole shape asserts as one string. */
function outline(nodes, depth = 0) {
	return nodes.flatMap((node) => [
		`${'  '.repeat(depth)}${node.point.type}:${node.point.text}`,
		...outline(node.children, depth + 1),
	])
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

	test('a textFrame edit marks the data part, not the slide, and leaves the cache stale', async () => {
		// The wiring this pins: `DiagramPoint` holds the diagram data part. Threading the
		// slide's part instead would mark the wrong part dirty, and the edit would vanish on
		// save while every in-memory read kept reporting it.
		//
		// It also pins the *documented* limit of this path, which used to be the only path.
		// `textFrame` is the escape hatch for run-level work, so it edits the data model and
		// nothing else; the drawing cache keeps the old string, and every renderer without a
		// SmartArt layout engine keeps painting it. `DiagramPoint.text` is what mirrors.
		const { presentation, frame } = await smartArtFrame()
		const point = frame.diagram.points.find((candidate) => candidate.text === NODE_TEXT[0])
		point.textFrame.text = 'Rewritten node'
		const saved = await presentation.save()
		const reopened = await Presentation.load(saved)
		const diagram = reopened.slides[1].shapes.find(isGraphicFrame).diagram
		assert(diagram.text.startsWith('Rewritten node\n'), 'the edit is in the saved data part')
		assert(!diagram.text.includes(NODE_TEXT[0]), 'and the original node text is gone')

		const drawing = new TextDecoder().decode((await partBodies(saved)).get('ppt/diagrams/drawing1.xml'))
		assert(!drawing.includes('Rewritten node'), 'the drawing cache did not follow')
		assert(drawing.includes('Understand'), 'it still holds the run-split original')
	})
})

describe('Diagram — the authored tree', () => {
	test('nodes give hList1 its three roots and their children in srcOrd order', async () => {
		const { frame } = await smartArtFrame()
		const nodes = frame.diagram.nodes
		assertEqual(nodes.length, 3, 'the doc root has three top-level nodes')
		assertEqual(
			outline(nodes).join('\n'),
			[
				'node:Understand Data Modelling Principles',
				'  node:Uncontrollable Inputs (environmental factors)',
				'  node:Controllable Inputs (decision variables)',
				'  node:Mathematical model',
				'node:Be able to create complex models using spreadsheets',
				'  node:Create models, simulate changes of uncontrollable inputs, see the impact on results',
				'  node:Graph the results',
				'  node:Be able to use Pivot Tables',
				'node:Use probabilities in models',
				'  node:Simulate uncontrollable inputs using probability distribution',
				'  node:Conduct simulations',
			].join('\n'),
			'the tree, with children in srcOrd order'
		)
		// The assertion that distinguishes srcOrd from document order: `Controllable Inputs` is
		// the *last* node in `dgm:ptLst` and the *second* child of the first branch.
		assertEqual(nodes[0].children[1].point.text, NODE_TEXT[10], 'srcOrd wins over document order')
		assertEqual(nodes[0].level, 0, 'a root is level 0')
		assertEqual(nodes[0].children[0].level, 1, 'its children are level 1')
		assert(nodes[0].children[0].parent === nodes[0], 'and point back at it')
		assert(nodes[0].parent === null, 'a root has no parent')
	})

	test('an org chart nests, and its assistant is a child like any other', async () => {
		const { diagram } = await familyDiagram(1)
		assertEqual(
			outline(diagram.nodes).join('\n'),
			[
				'node:org-root',
				'  asst:org-asst',
				'  node:org-child-1',
				'    node:org-grandchild',
				'  node:org-child-2',
				'  node:org-child-3',
			].join('\n'),
			'three levels, with the asst in the tree and typed apart from the nodes'
		)
		const asst = diagram.nodes[0].children[0]
		assertEqual(asst.point.type, 'asst', 'an assistant keeps its own type rather than folding into node')
		assertEqual(asst.level, 1, 'and sits at the level its parOf edge puts it at')
	})

	test('transition points stay out of the tree and are reached through their edge', async () => {
		// `process1` labels three of its arrows, so the fixture has transition points that carry
		// real text. Putting them in `children` would make `nodes` disagree with what
		// PowerPoint's own text pane shows, which is the whole reason the tree is nodes-only.
		const { diagram } = await familyDiagram(2)
		assertEqual(outline(diagram.nodes).join('|'), 'node:proc-1|node:proc-2|node:proc-3|node:proc-4', 'four flat nodes')
		const labels = diagram.points.filter((point) => point.type === 'sibTrans' && point.text !== '')
		assertEqual(labels.length, 3, 'and three labelled sibling transitions beside them')
		for (const label of labels) {
			const edge = diagram.connections.find((connection) => connection.siblingTransitionId === label.modelId)
			assert(edge, `the label ${label.text} is reachable from the edge it labels`)
			assertEqual(label.connectionId, edge.modelId, 'and names that edge back')
		}
	})

	test('point() resolves a connection id, and misses cleanly', async () => {
		const { frame } = await smartArtFrame()
		const diagram = frame.diagram
		const edge = diagram.connections.find((connection) => connection.type === 'parOf' && connection.sourceOrder === 0)
		assertEqual(diagram.point(edge.destinationId).modelId, edge.destinationId, 'a connection end resolves to its point')
		assert(diagram.point('{00000000-0000-0000-0000-000000000000}') === null, 'an id naming no point reads null')
	})

	test('a parOf cycle raises rather than walking forever', async () => {
		// A corrupt model, not a diagram shape: PowerPoint cannot author one and the walk
		// cannot terminate on one. Built by re-pointing a child's parOf edge at its own
		// descendant, which is the smallest cycle a real data model could acquire.
		const { diagram } = await familyDiagram(1)
		const child = diagram.nodes[0].children[1]
		const grandchild = child.children[0]
		assert(grandchild, 'org-child-1 has the grandchild the cycle is built through')
		// `parOf` is the schema default, so a real edge carries no `@type` at all — filtering
		// on the string would have matched nothing and left the graph acyclic.
		let repointed = 0
		for (const cxn of diagram.element_.getElementsByTagName('dgm:cxn')) {
			if (cxn.getAttribute('type')) continue
			if (cxn.getAttribute('destId') !== child.point.modelId) continue
			cxn.setAttribute('srcId', grandchild.point.modelId)
			repointed++
		}
		assertEqual(repointed, 1, 'exactly one parOf edge was re-pointed into the cycle')
		let raised = null
		try {
			void diagram.nodes
		} catch (error) {
			raised = error
		}
		assert(raised, 'walking a cyclic parOf graph raises rather than hanging')
		assertEqual(raised.code, 'diagram/parent-edge-cycle', 'and names the condition')
	})
})

describe('DiagramPoint — the link to what is drawn', () => {
	test('several nodes resolve to one drawn shape, at the paragraph destOrd names', async () => {
		// The measured many-to-one case: one `dsp:sp` presents three nodes, and `@destOrd`
		// orders them *against* document order. A mapping that walked the drawing in document
		// order would agree on the first branch and be wrong on this one.
		const { frame } = await smartArtFrame()
		const branch = frame.diagram.nodes[0].children
		const drawn = branch.map((node) => node.point.drawnShape)
		assert(drawn.every(Boolean), 'all three children resolve to a drawn shape')
		assertEqual(new Set(drawn.map((shape) => shape.modelId)).size, 1, 'and it is one and the same shape')
		assertEqual(drawn.map((shape) => shape.paragraphIndex).join(','), '0,1,2', 'at consecutive paragraphs')
		for (const [index, node] of branch.entries()) {
			assertEqual(
				drawn[index].textFrame.paragraphs[drawn[index].paragraphIndex].text,
				node.point.text,
				`paragraph ${index} draws the node that resolved to it`
			)
		}
		assertEqual(drawn[0].part.partName, '/ppt/diagrams/drawing1.xml', 'bound to the drawing part, not the data part')
		assertEqual(drawn[0].modelId, branch[0].point.presentationId, 'keyed by the pres point presentationId names')
		assert(
			frame.diagram.points.every((point) => point.drawnShape?.modelId !== point.modelId),
			'and never by the authored point own modelId, which draws nothing'
		)
	})

	test('a point with nothing drawn for it reads null rather than guessing', async () => {
		// Three of the measured ways to resolve to nothing, all in one deck: no `presOf` edge,
		// a `pres` point that draws no `dsp:sp`, and a generated point that presents nothing.
		const { diagram } = await familyDiagram(4)
		const parTrans = diagram.points.filter((point) => point.type === 'parTrans')
		assert(parTrans.length > 0, 'pList1 has parent transitions')
		for (const point of parTrans) {
			assert(point.presentationId === null, 'an unlabelled parTrans has no presOf edge at all')
			assert(point.drawnShape === null, 'so nothing is drawn for it')
		}
		const sibTrans = diagram.points.filter((point) => point.type === 'sibTrans' && point.presentationId !== null)
		assert(sibTrans.length > 0, 'while its sibling transitions do have one')
		for (const point of sibTrans) assert(point.drawnShape === null, 'whose pres point draws no dsp:sp')

		const pres = diagram.points.filter((point) => point.type === 'pres')
		assert(pres.length > 0, 'and a generated pres point presents nothing itself')
		for (const point of pres) assert(point.drawnShape === null, 'so it too reads null')
	})

	test('a diagram whose drawing part is absent resolves the pres point and nothing more', async () => {
		const pkg = await OpcPackage.load(await readFixture('mixed'))
		assertEqual(pkg.removePart('/ppt/diagrams/drawing1.xml'), true, 'the drawing part was there to remove')
		const presentation = await Presentation.load(await pkg.save())
		const diagram = presentation.slides[1].shapes.find(isGraphicFrame).diagram
		assert(diagram.drawingPart === null, 'the fallback drawing is gone')
		const node = diagram.nodes[0].point
		assert(node.presentationId !== null, 'the presOf edge still names the pres point it always did')
		assert(node.drawnShape === null, 'but there is no shape to reach')
	})
})

describe('DiagramPoint.text — re-texting a node, cache and all', () => {
	/** The `<a:t>` payloads of a saved drawing part, in document order. */
	async function drawnStrings(saved) {
		return [...(await drawingXml(saved)).matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((match) => match[1])
	}

	async function drawingXml(saved) {
		return new TextDecoder().decode((await partBodies(saved)).get('ppt/diagrams/drawing1.xml'))
	}

	/**
	 * The verbatim `<a:p>` blocks of one drawn shape, so "untouched" can be asserted as byte
	 * identity rather than as equal text. A `dsp:sp` never nests, so scanning to the first
	 * close tag is exact.
	 */
	async function drawnParagraphs(saved, modelId) {
		const xml = await drawingXml(saved)
		const start = xml.indexOf(`<dsp:sp modelId="${modelId}"`)
		assert(start >= 0, `the drawing part holds a dsp:sp for ${modelId}`)
		const sp = xml.slice(start, xml.indexOf('</dsp:sp>', start))
		return [...sp.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((match) => match[0])
	}

	test('writes the data model and the drawing cache, and both survive a save', async () => {
		const { presentation, frame } = await smartArtFrame()
		const { diagnostics } = await captureDiagnostics(async () => {
			frame.diagram.nodes[0].point.text = 'Rewritten node'
		})
		assertEqual(diagnostics.length, 0, 'a node that resolves needs no diagnostic')

		const saved = await presentation.save()
		const reopened = await Presentation.load(saved)
		const diagram = reopened.slides[1].shapes.find(isGraphicFrame).diagram
		assertEqual(diagram.nodes[0].point.text, 'Rewritten node', 'the data model holds the new text')
		assertEqual(
			diagram.nodes[0].point.drawnShape.textFrame.paragraphs[0].text,
			'Rewritten node',
			'and so does the paragraph the drawing cache draws for it'
		)
		assert(!(await drawnStrings(saved)).includes('Understand'), 'with no fragment of the old string left behind')
	})

	test('editing one of three nodes sharing a drawn shape leaves the other two byte-identical', async () => {
		// The assertion that catches a mirror written through `TextFrame.text`: that setter
		// collapses the whole body to one paragraph, which would delete the siblings' text
		// from the cache while leaving the data model perfectly correct.
		const { presentation, frame } = await smartArtFrame()
		const branch = frame.diagram.nodes[0].children
		const shared = branch[0].point.drawnShape.modelId
		const before = await drawnParagraphs(await (await openFixture('mixed')).save(), shared)
		assertEqual(before.length, 3, 'one drawn shape, three paragraphs, one per node')

		branch[1].point.text = 'Only the middle one'
		const saved = await presentation.save()
		const after = await drawnParagraphs(saved, shared)
		assertEqual(after.length, 3, 'still three paragraphs afterwards')
		assertEqual(after[0], before[0], 'the paragraph above the edit is byte-identical')
		assertEqual(after[2], before[2], 'and so is the one below it')
		assert(after[1] !== before[1], 'while the edited one changed')
		assert(after[1].includes('<a:t>Only the middle one</a:t>'), 'to a single run holding the new string')

		const reopened = await Presentation.load(saved)
		const drawn = reopened.slides[1].shapes
			.find(isGraphicFrame)
			.diagram.nodes[0].children.map((node) => node.point.drawnShape)
		assertEqual(drawn.map((shape) => shape.paragraphIndex).join(','), '0,1,2', 'and the three still map as they did')
	})

	test('marks the data part and the drawing part, and nothing else', async () => {
		const input = await readFixture('mixed')
		const presentation = await Presentation.load(input)
		presentation.slides[1].shapes.find(isGraphicFrame).diagram.nodes[0].point.text = 'Rewritten node'
		const before = await partBodies(input)
		const after = await partBodies(await presentation.save())
		assertUnchangedExcept(before, after, ['ppt/diagrams/data1.xml', 'ppt/diagrams/drawing1.xml'])
	})

	test('a point that resolves to no drawn paragraph still edits the data model, and says so once', async () => {
		// Every way a node can fail to reach the cache, each measured on a real deck. The edit
		// PowerPoint reads always lands; the diagnostic is what tells the caller the deck will
		// look unchanged in a renderer that does not recompute the drawing.
		/** @type {{ label: string, resolve: () => Promise<import('../../dist/read.js').DiagramPoint> }[]} */
		const cases = [
			{
				label: 'no drawing part in the package',
				async resolve() {
					const pkg = await OpcPackage.load(await readFixture('mixed'))
					pkg.removePart('/ppt/diagrams/drawing1.xml')
					const presentation = await Presentation.load(await pkg.save())
					return presentation.slides[1].shapes.find(isGraphicFrame).diagram.nodes[0].point
				},
			},
			{
				label: 'a pres point that draws no dsp:sp',
				async resolve() {
					const { diagram } = await familyDiagram(4)
					const point = diagram.points.find(
						(candidate) => candidate.type === 'sibTrans' && candidate.presentationId !== null
					)
					assert(point?.textFrame, 'pList1 has a sibTrans with a dgm:t and nothing drawn for it')
					return point
				},
			},
		]
		for (const { label, resolve } of cases) {
			const point = await resolve()
			const { diagnostics } = await captureDiagnostics(async () => {
				point.text = 'Mirrored nowhere'
			})
			assertEqual(point.text, 'Mirrored nowhere', `${label}: the data-model edit applies regardless`)
			assertEqual(diagnostics.length, 1, `${label}: exactly one diagnostic`)
			assertEqual(diagnostics[0].code, 'diagram/drawing-cache-not-updated', `${label}: naming the condition`)
		}
	})

	test('a point with no dgm:t is left alone rather than given one', async () => {
		// The one case measurement says not to paper over: an unlabelled transition point in a
		// layout with no place for a label is stored exactly like this, and PowerPoint strips
		// text put on one at the next save. Synthesizing a body would produce an edit that
		// reads back correctly here and vanishes the first time the deck is opened.
		const { diagram } = await familyDiagram(2)
		const point = diagram.points.find((candidate) => candidate.type === 'parTrans' && candidate.textFrame === null)
		assert(point, 'process1 has a transition point with no text body at all')
		const { diagnostics } = await captureDiagnostics(async () => {
			point.text = 'Not a label this layout has room for'
		})
		assertEqual(point.text, '', 'nothing was written')
		assertEqual(diagnostics.length, 1, 'and one diagnostic explains why')
		assertEqual(diagnostics[0].code, 'diagram/point-has-no-text-body', 'named apart from the cache-mirror one')
	})

	test('a labelled transition re-texts through exactly the same path a node does', async () => {
		// `process1` gives a surviving arrow label its own `dsp:sp`, so the mapping needs no
		// separate arm for it. Editing one is the proof.
		const { presentation, diagram } = await familyDiagram(2)
		const label = diagram.points.find((point) => point.text === 'sibTrans-4')
		assert(label, 'the fixture has the injected arrow label')
		const { diagnostics } = await captureDiagnostics(async () => {
			label.text = 'relabelled arrow'
		})
		assertEqual(diagnostics.length, 0, 'no diagnostic: the label resolves like any node')

		const reopened = await Presentation.load(await presentation.save())
		const saved = reopened.slides[1].shapes.find(isGraphicFrame).diagram
		const relabelled = saved.points.find((point) => point.text === 'relabelled arrow')
		assert(relabelled, 'the data model carries the new label')
		assertEqual(relabelled.drawnShape.textFrame.paragraphs[0].text, 'relabelled arrow', 'and so does the cache')
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
