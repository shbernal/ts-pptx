// What a copy of a page shares with the page it came from, and what it must own.
//
// PowerPoint refuses to open a package where two slides resolve to one chart or
// one SmartArt diagram — `0x80070570`, the whole deck rejected, and the schema
// validator accepts the file, so only the application says so. The library's rule
// (`src/read/api/ops/page-owned.ts`) is therefore: a page copy shares deck-wide
// assets (layout, master, theme, media) and takes its own copy of everything the
// page owns, with the subtree under each owned part copied along with it.
//
// Contract under test, across the three ways a page gets duplicated —
// `cloneSlide`, `importSlide` of one source page twice, and `importSlides` naming
// it twice:
//   - the copy's chart / diagram / notes parts are parts of its own;
//   - the subtree under an owned part comes with it (a chart's embedded workbook
//     and user-shapes drawing), while media stay shared;
//   - a relationship back to the page (a notes slide names its slide) follows the
//     copy rather than pointing at the original;
//   - shared deck furniture is still shared, and the result is schema-valid.

import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { openFixture } from './corpus.js'
import { assertNoDanglingRels } from './opc.js'

/** Index of the `mixed` fixture pages this file leans on. */
const NOTES_PAGE = 0 // a page with speaker notes
const DIAGRAM_PAGE = 1 // SmartArt: five diagram parts
const CHART_PAGE = 7 // a chart with an embedded workbook and user shapes

/** Internal relationship targets of one part, by rel-type suffix. */
function targetsByType(opc, partName, typeSuffix) {
	const rels = opc.relationshipsFor(partName)
	return [...rels]
		.filter((rel) => rel.targetMode !== 'External' && rel.type.endsWith(`/${typeSuffix}`))
		.map((rel) => rels.resolveTarget(rel.id))
}

/** The one internal target of `typeSuffix`, asserting there is exactly one. */
function targetByType(opc, partName, typeSuffix) {
	const found = targetsByType(opc, partName, typeSuffix)
	assertEqual(found.length, 1, `${partName} has exactly one ${typeSuffix} relationship`)
	return found[0]
}

/** Every internal partname reachable from `partName`, itself excluded. */
function reachable(opc, partName, seen = new Set()) {
	const rels = opc.relationshipsFor(partName)
	for (const rel of rels) {
		if (rel.targetMode === 'External') continue
		const target = rels.resolveTarget(rel.id)
		if (seen.has(target)) continue
		seen.add(target)
		reachable(opc, target, seen)
	}
	return seen
}

describe('a page copy owns what the page owned', () => {
	test('cloneSlide gives the clone its own chart, workbook and user shapes', async () => {
		const deck = await openFixture('mixed')
		const source = deck.slides[CHART_PAGE]
		const clone = deck.cloneSlide(CHART_PAGE)

		const reopened = await Presentation.load(await deck.save())
		assertNoDanglingRels(reopened.opc)
		const sourceChart = targetByType(reopened.opc, source.partName, 'chart')
		const cloneChart = targetByType(reopened.opc, clone.partName, 'chart')
		assert(sourceChart !== cloneChart, 'the clone points at a chart part of its own')

		// The chart's own subtree comes with it. Sharing the workbook or the
		// user-shapes drawing is refused by PowerPoint just as sharing the chart is.
		const sourceUnder = reachable(reopened.opc, sourceChart)
		const cloneUnder = reachable(reopened.opc, cloneChart)
		assert(sourceUnder.size > 0, 'the fixture chart has a subtree to copy')
		assertEqual(cloneUnder.size, sourceUnder.size, 'the copy carries the same shape of subtree')
		for (const partName of cloneUnder) {
			assert(!sourceUnder.has(partName), `${partName} is the copy's own, not the source's`)
		}

		// Deck furniture is still shared: this is a copy, not a second deck.
		assertEqual(
			targetByType(reopened.opc, clone.partName, 'slideLayout'),
			targetByType(reopened.opc, source.partName, 'slideLayout'),
			'both pages still use the one layout'
		)
	})

	test('cloneSlide gives the clone its own notes slide, wired back to the clone', async () => {
		const deck = await openFixture('mixed')
		const source = deck.slides[NOTES_PAGE]
		const clone = deck.cloneSlide(NOTES_PAGE)

		const reopened = await Presentation.load(await deck.save())
		assertNoDanglingRels(reopened.opc)
		const sourceNotes = targetByType(reopened.opc, source.partName, 'notesSlide')
		const cloneNotes = targetByType(reopened.opc, clone.partName, 'notesSlide')
		assert(sourceNotes !== cloneNotes, 'the clone has notes of its own')
		assertEqual(
			targetByType(reopened.opc, cloneNotes, 'slide'),
			clone.partName,
			'the copied notes slide names the clone, not the page it was copied from'
		)
		assertEqual(
			targetByType(reopened.opc, cloneNotes, 'notesMaster'),
			targetByType(reopened.opc, sourceNotes, 'notesMaster'),
			'and both notes slides share the one notes master'
		)
	})

	test('cloneSlide keeps shared media shared', async () => {
		// The rule is not "copy everything under the page": PowerPoint stores one
		// image and points every shape that shows it at that copy, and so does this.
		const deck = await openFixture('image')
		const source = deck.slides[1]
		const before = [...deck.opc.parts.keys()].filter((name) => name.startsWith('/ppt/media/')).length
		const clone = deck.cloneSlide(1)

		const reopened = await Presentation.load(await deck.save())
		const after = [...reopened.opc.parts.keys()].filter((name) => name.startsWith('/ppt/media/')).length
		assertEqual(after, before, 'no media part was duplicated')
		assertEqual(
			JSON.stringify(targetsByType(reopened.opc, clone.partName, 'image').sort()),
			JSON.stringify(targetsByType(reopened.opc, source.partName, 'image').sort()),
			'the clone shows the same image parts as the page it copies'
		)
	})

	test('importSlide twice gives each imported page its own diagram parts', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const first = target.importSlide(source, DIAGRAM_PAGE)
		const second = target.importSlide(source, DIAGRAM_PAGE)

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		for (const type of ['diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors', 'diagramDrawing']) {
			const a = targetByType(reopened.opc, first.partName, type)
			const b = targetByType(reopened.opc, second.partName, type)
			assert(a !== b, `the two imported pages have their own ${type} part`)
		}
		// The subgraph that is genuinely shared still is — one master, one theme.
		assertEqual(
			targetByType(reopened.opc, first.partName, 'slideLayout'),
			targetByType(reopened.opc, second.partName, 'slideLayout'),
			'both imported pages share the one imported layout'
		)
	})

	test('importSlides duplicating a chart page copies the chart per requested page', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const [first, second] = target.importSlides([
			{ source, sourceIndex: CHART_PAGE, outputIndex: 0 },
			{ source, sourceIndex: CHART_PAGE, outputIndex: 1 },
		])

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		const firstChart = targetByType(reopened.opc, first.partName, 'chart')
		const secondChart = targetByType(reopened.opc, second.partName, 'chart')
		assert(firstChart !== secondChart, 'each requested page got a chart part of its own')
		const firstUnder = reachable(reopened.opc, firstChart)
		for (const partName of reachable(reopened.opc, secondChart)) {
			assert(!firstUnder.has(partName), `${partName} belongs to one chart only`)
		}
	})

	test('two different pages of one source still share what they legitimately share', async () => {
		// Ownership is per page, not per import: two *different* pages coming across
		// from one source must not each drag a private layout/master/theme.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const [a, b] = target.importSlides([
			{ source, sourceIndex: 2, outputIndex: 0 },
			{ source, sourceIndex: 3, outputIndex: 1 },
		])

		const reopened = await Presentation.load(await target.save())
		const layoutA = targetByType(reopened.opc, a.partName, 'slideLayout')
		const layoutB = targetByType(reopened.opc, b.partName, 'slideLayout')
		const masterOf = (layout) => targetByType(reopened.opc, layout, 'slideMaster')
		assertEqual(masterOf(layoutA), masterOf(layoutB), 'both pages hang off the one imported master')
	})

	test('importShape carrying one chart shape twice gives each frame its own chart', async () => {
		// The same rule one level down: two graphic frames resolving to one chart part
		// is the package PowerPoint refuses, whether the frames arrived as pages or as
		// shapes. Media stay shared through the batch's rel-id cache either way.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const chartSlide = source.slides[CHART_PAGE]
		const chartIndex = chartSlide.shapes.findIndex((shape) => shape.shapeType === 'graphicFrame' && shape.chart)
		assert(chartIndex >= 0, 'the fixture page has a chart shape to carry')

		target.importShape(target.slides[3], chartSlide, chartIndex)
		target.importShape(target.slides[4], chartSlide, chartIndex)

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		const first = targetByType(reopened.opc, reopened.slides[3].partName, 'chart')
		const second = targetByType(reopened.opc, reopened.slides[4].partName, 'chart')
		assert(first !== second, 'each carried frame got a chart part of its own')
	})

	test.skipIf(!validatorInstalled)('a deck of duplicated chart and diagram pages stays schema-valid', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		target.importSlides([
			{ source, sourceIndex: CHART_PAGE, outputIndex: 0 },
			{ source, sourceIndex: CHART_PAGE, outputIndex: 1 },
			{ source, sourceIndex: DIAGRAM_PAGE, outputIndex: 2 },
			{ source, sourceIndex: DIAGRAM_PAGE, outputIndex: 3 },
		])
		target.cloneSlide(0)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
