// Cross-slide shape-import tests for `pptxgenjs/read`.
//
// Contract under test: Presentation.importShape(target, source, index, opts)
// copies one shape — autoshape, picture, table/chart graphic frame, or group —
// from a slide of any open package onto a slide of THIS package. It drags the
// shape's media/chart/embedding parts across (deduped via the copy registry),
// rewrites their relationship references to fresh host-slide rels, reassigns the
// shape's (and any group children's) drawing ids so they cannot collide with the
// host, optionally bakes the source theme to literals (`preserve`, the default)
// or leaves it symbolic (`restyle`), positions it (`left`/`top`/`width`/`height`,
// `at`), survives a save→reopen round-trip with no dangling rels, and keeps the
// package schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

/** Every internal relationship of every part resolves to a part that exists (no dangling rels). */
function assertNoDanglingRels(opc) {
	for (const partName of opc.parts.keys()) {
		if (partName.endsWith('.rels')) continue
		for (const rel of opc.relationshipsFor(partName)) {
			if (rel.targetMode === 'External') continue
			const target = opc.relationshipsFor(partName).resolveTarget(rel.id)
			assert(opc.part(target), `${partName} → ${rel.id} targets an existing part (${target})`)
		}
	}
}

/** Count the package parts whose name matches `re`. */
function countParts(opc, re) {
	return [...opc.parts.keys()].filter((n) => re.test(n)).length
}

/** All `p:cNvPr/@id` values within a shape subtree (inclusive of nested group children). */
function cNvPrIds(element) {
	const live = element.getElementsByTagNameNS(P_NS, 'cNvPr')
	const out = []
	for (let i = 0; i < live.length; i++) {
		const id = live[i].getAttribute('id')
		if (id != null) out.push(Number(id))
	}
	return out
}

/** Geometry tuples (`a:off`/`a:ext`/`a:chOff`/`a:chExt`) within a subtree, document order. */
function geometryTuples(element) {
	const out = []
	for (const tag of ['off', 'ext', 'chOff', 'chExt']) {
		const live = element.getElementsByTagNameNS(A_NS, tag)
		for (let i = 0; i < live.length; i++) {
			const el = live[i]
			out.push(
				`${tag}:${el.getAttribute('x') ?? el.getAttribute('cx')},${el.getAttribute('y') ?? el.getAttribute('cy')}`
			)
		}
	}
	return out
}

/** Table column widths (`a:gridCol@w`, EMU) within a subtree, document order. */
function gridColWidths(element) {
	const live = element.getElementsByTagNameNS(A_NS, 'gridCol')
	const out = []
	for (let i = 0; i < live.length; i++) {
		const w = live[i].getAttribute('w')
		if (w != null) out.push(Number(w))
	}
	return out
}

function schemeClrCount(element) {
	return element.getElementsByTagNameNS(A_NS, 'schemeClr').length
}

function srgbClrCount(element) {
	return element.getElementsByTagNameNS(A_NS, 'srgbClr').length
}

/** Index of the first shape on `slide` matching `pred`. */
function findShapeIndex(slide, pred) {
	return slide.shapes.findIndex(pred)
}

describe('Presentation.importShape', () => {
	test('lifts a picture onto a foreign host; media copied once, survives a round-trip', async () => {
		const target = await open('empty') // 16:9, slide[0] holds one autoShape
		const source = await open('image') // 16:9
		const targetSlide = target.slides[0]
		const picIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		assert(picIndex >= 0, 'source slide has a picture to lift')
		const hostMediaBefore = countParts(target.opc, /ppt\/media\//)
		const hostShapesBefore = targetSlide.shapes.length

		const shape = target.importShape(targetSlide, source.slides[0], picIndex)
		assertEqual(shape.shapeType, 'picture', 'returns a Picture proxy')
		assertEqual(targetSlide.shapes.length, hostShapesBefore + 1, 'one shape was appended to the host slide')

		// Exactly one media part was copied in, and the blip points at it.
		assertEqual(countParts(target.opc, /ppt\/media\//), hostMediaBefore + 1, 'one media part copied')
		assert(shape.imagePartName && target.opc.part(shape.imagePartName), 'blip resolves to a copied media part')

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		const pic = reopened.slides[0].shapes.find((s) => s.shapeType === 'picture')
		assert(pic && pic.imagePartName && reopened.opc.part(pic.imagePartName), 'imported picture survives the round-trip')
	})

	test('lifts a table; cells intact on a host with a different theme', async () => {
		const target = await open('empty')
		const source = await open('table')
		const tableIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'graphicFrame' && s.table)
		assert(tableIndex >= 0, 'source slide has a table')
		const srcTable = source.slides[0].shapes[tableIndex].table
		const srcRows = srcTable.rowCount
		const srcFirstCell = srcTable.rows[0].cells[0].text

		const shape = target.importShape(target.slides[0], source.slides[0], tableIndex)
		assertEqual(shape.shapeType, 'graphicFrame', 'returns the graphic frame')
		assert(shape.table, 'the lifted frame still hosts a table')
		assertEqual(shape.table.rowCount, srcRows, 'row count preserved')
		assertEqual(shape.table.rows[0].cells[0].text, srcFirstCell, 'first cell text preserved')

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		const table = reopened.slides[0].shapes.find((s) => s.shapeType === 'graphicFrame' && s.table)?.table
		assert(table && table.rowCount === srcRows, 'table survives the round-trip')
	})

	test('lifts a chart with its part and embedded workbook', async () => {
		const target = await open('mixed') // 4:3
		const source = await open('mixed')
		const chartSlide = source.slides.findIndex((s) =>
			s.shapes.some((sh) => sh.shapeType === 'graphicFrame' && sh.chart)
		)
		assert(chartSlide >= 0, 'source deck has a chart slide')
		const chartIndex = findShapeIndex(source.slides[chartSlide], (s) => s.shapeType === 'graphicFrame' && s.chart)
		const chartsBefore = countParts(target.opc, /\/charts\/chart\d+\.xml$/)

		const shape = target.importShape(target.slides[0], source.slides[chartSlide], chartIndex)
		assert(shape.chart, 'the lifted frame still hosts a chart')

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
		assert(countParts(reopened.opc, /\/charts\/chart\d+\.xml$/) > chartsBefore, 'a chart part was copied in')
		assert(countParts(reopened.opc, /\/embeddings\//) > 0, 'the chart embedding workbook is present')
	})

	test('lifts a group; child ids reassigned, child offsets intact', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const grpSlide = source.slides.findIndex((s) => s.shapes.some((sh) => sh.shapeType === 'group'))
		assert(grpSlide >= 0, 'source deck has a group')
		const grpIndex = findShapeIndex(source.slides[grpSlide], (s) => s.shapeType === 'group')
		const srcGroup = source.slides[grpSlide].shapes[grpIndex].element_
		const srcIds = new Set(cNvPrIds(srcGroup))
		const srcGeom = geometryTuples(srcGroup)

		// `copy` keeps the subtree verbatim so geometry is a clean byte-for-byte check.
		const shape = target.importShape(target.slides[0], source.slides[grpSlide], grpIndex, { theme: 'copy' })
		assertEqual(shape.shapeType, 'group', 'returns a GroupShape')
		const importedIds = cNvPrIds(shape.element_)

		// Every id (group + children) was reassigned away from the source ids and is
		// unique within the host slide.
		for (const id of importedIds) assert(!srcIds.has(id), `child id ${id} was reassigned off the source ids`)
		const hostIds = cNvPrIds(target.slides[0].shapeTree())
		assertEqual(new Set(hostIds).size, hostIds.length, 'all host drawing ids are unique')

		// No rescale: every off/ext/chOff/chExt matches the source verbatim.
		assertEqual(
			JSON.stringify(geometryTuples(shape.element_)),
			JSON.stringify(srcGeom),
			'child offsets/extents unchanged'
		)
	})

	test('preserve bakes scheme colours to literals; restyle leaves them symbolic', async () => {
		const source = await open('mixed')
		const schemeSlide = source.slides.findIndex((s) =>
			s.shapes.some((sh) => sh.shapeType === 'autoShape' && schemeClrCount(sh.element_) > 0)
		)
		assert(schemeSlide >= 0, 'source deck has a scheme-coloured autoshape')
		const schemeIndex = findShapeIndex(
			source.slides[schemeSlide],
			(s) => s.shapeType === 'autoShape' && schemeClrCount(s.element_) > 0
		)
		const sourceSchemeClrs = schemeClrCount(source.slides[schemeSlide].shapes[schemeIndex].element_)

		const preserveTarget = await open('mixed')
		const preserved = preserveTarget.importShape(preserveTarget.slides[0], source.slides[schemeSlide], schemeIndex, {
			theme: 'preserve',
		})
		assert(schemeClrCount(preserved.element_) < sourceSchemeClrs, 'preserve resolved scheme colours to literals')
		assert(srgbClrCount(preserved.element_) > 0, 'preserve emitted literal srgbClr')

		const restyleTarget = await open('mixed')
		const restyled = restyleTarget.importShape(restyleTarget.slides[0], source.slides[schemeSlide], schemeIndex, {
			theme: 'restyle',
		})
		assertEqual(schemeClrCount(restyled.element_), sourceSchemeClrs, 'restyle left every scheme colour symbolic')
	})

	test('reassigns the lifted shape id off every host id (no collision)', async () => {
		const target = await open('empty')
		const source = await open('image')
		const hostIdsBefore = new Set(cNvPrIds(target.slides[0].shapeTree()))
		const picIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')

		const shape = target.importShape(target.slides[0], source.slides[0], picIndex)
		assert(!hostIdsBefore.has(shape.id), `imported id ${shape.id} differs from every host id`)
		const hostIds = cNvPrIds(target.slides[0].shapeTree())
		assertEqual(new Set(hostIds).size, hostIds.length, 'all host drawing ids remain unique')
	})

	test('honours placement overrides and z-order', async () => {
		const target = await open('empty')
		const source = await open('image')
		const picIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')

		const shape = target.importShape(target.slides[0], source.slides[0], picIndex, {
			left: 100000,
			top: 200000,
			width: 300000,
			height: 400000,
			at: 0,
		})
		assertEqual(shape.left, 100000, 'left override applied')
		assertEqual(shape.top, 200000, 'top override applied')
		assertEqual(shape.width, 300000, 'width override applied')
		assertEqual(shape.height, 400000, 'height override applied')
		assertEqual(target.slides[0].shapes[0].id, shape.id, 'at:0 placed the shape backmost (first in document order)')
	})

	test('batch imports several shapes in order, with unique ids', async () => {
		const target = await open('image')
		const source = await open('image')
		const targetSlide = target.slides[0]
		const indices = [0, 1] // picture + autoShape on image slide[0]
		const before = targetSlide.shapes.length

		const shapes = target.importShapes(targetSlide, source.slides[0], indices)
		assertEqual(shapes.length, 2, 'two shapes returned')
		assertEqual(shapes[0].shapeType, source.slides[0].shapes[0].shapeType, 'first lifted shape matches first index')
		assertEqual(shapes[1].shapeType, source.slides[0].shapes[1].shapeType, 'second lifted shape matches second index')
		assertEqual(targetSlide.shapes.length, before + 2, 'both shapes appended')
		const hostIds = cNvPrIds(targetSlide.shapeTree())
		assertEqual(new Set(hostIds).size, hostIds.length, 'all ids unique after the batch')

		const reopened = await Presentation.load(await target.save())
		assertNoDanglingRels(reopened.opc)
	})

	test('dedupes shared media across repeated imports from the same source', async () => {
		const target = await open('empty')
		const source = await open('image')
		const picIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		const before = countParts(target.opc, /ppt\/media\//)

		const a = target.importShape(target.slides[0], source.slides[0], picIndex)
		const b = target.importShape(target.slides[0], source.slides[0], picIndex)

		// The same source image was copied exactly once (registry dedupe), even though
		// the two pictures get distinct host relationships.
		assertEqual(countParts(target.opc, /ppt\/media\//), before + 1, 'shared media copied once')
		assertEqual(a.imagePartName, b.imagePartName, 'both pictures resolve to the one copied media part')
	})

	test('rejects a slide-size mismatch between source and target', async () => {
		const target = await open('empty') // 16:9
		const source = await open('mixed') // 4:3
		assert(
			throws(() => target.importShape(target.slides[0], source.slides[0], 0)),
			'lifting a shape across mismatched slide sizes throws'
		)
	})

	test('rejects an out-of-range shape index', async () => {
		const target = await open('empty')
		const source = await open('image')
		assert(
			throws(() => target.importShape(target.slides[0], source.slides[0], 99)),
			'lifting a missing shape throws'
		)
	})

	test('rejects a target slide from another presentation', async () => {
		const target = await open('empty')
		const other = await open('empty')
		const source = await open('image')
		assert(
			throws(() => target.importShape(other.slides[0], source.slides[0], 0)),
			'a target slide not owned by this presentation throws'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with a lifted picture stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('image')
		const picIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		target.importShape(target.slides[0], source.slides[0], picIndex)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test.skipIf(!validatorInstalled)('a deck with a lifted table stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('table')
		const tableIndex = findShapeIndex(source.slides[0], (s) => s.shapeType === 'graphicFrame' && s.table)
		target.importShape(target.slides[0], source.slides[0], tableIndex)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test.skipIf(!validatorInstalled)('a deck with a lifted chart and group stays schema-valid', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const chartSlide = source.slides.findIndex((s) =>
			s.shapes.some((sh) => sh.shapeType === 'graphicFrame' && sh.chart)
		)
		const chartIndex = findShapeIndex(source.slides[chartSlide], (s) => s.shapeType === 'graphicFrame' && s.chart)
		target.importShape(target.slides[0], source.slides[chartSlide], chartIndex)
		const grpSlide = source.slides.findIndex((s) => s.shapes.some((sh) => sh.shapeType === 'group'))
		const grpIndex = findShapeIndex(source.slides[grpSlide], (s) => s.shapeType === 'group')
		target.importShape(target.slides[0], source.slides[grpSlide], grpIndex)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('Presentation.importShape({ rescale })', () => {
	// image/table (16:9, 12192000×6858000) → mixed (4:3, 9144000×6858000):
	// fit scale = min(0.75, 1.0) = 0.75; centering dx = 0, dy = 857250.
	const near = (got, want, label) => assert(Math.abs(got - want) <= 2, `${label}: ${got} ≈ ${want}`)

	test("'fit' scales the lifted shape uniformly and centers the slack", async () => {
		const target = await open('mixed') // 4:3
		const source = await open('image') // 16:9
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		const src = source.slides[0].shapes[idx].absoluteFrame

		const shape = target.importShape(target.slides[0], source.slides[0], idx, { rescale: 'fit' })
		near(shape.left, Math.round(src.left * 0.75), 'left scaled by 0.75')
		near(shape.top, Math.round(src.top * 0.75 + 857250), 'top scaled + centered')
		near(shape.width, Math.round(src.width * 0.75), 'width scaled by 0.75')
		near(shape.height, Math.round(src.height * 0.75), 'height scaled by 0.75')
	})

	test("'stretch' scales each axis independently (height holds when only width differs)", async () => {
		const target = await open('mixed')
		const source = await open('image')
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		const src = source.slides[0].shapes[idx].absoluteFrame

		const shape = target.importShape(target.slides[0], source.slides[0], idx, { rescale: 'stretch' })
		near(shape.width, Math.round(src.width * 0.75), 'width scaled by sx (0.75)')
		near(shape.height, src.height, 'height unchanged (sy = 1.0)')
		near(shape.top, src.top, 'top unchanged (no centering, sy = 1.0)')
	})

	test('true is an alias for fit', async () => {
		const target = await open('mixed')
		const source = await open('image')
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		const src = source.slides[0].shapes[idx].absoluteFrame

		const shape = target.importShape(target.slides[0], source.slides[0], idx, { rescale: true })
		near(shape.width, Math.round(src.width * 0.75), 'rescale:true scales like fit')
		near(shape.top, Math.round(src.top * 0.75 + 857250), 'rescale:true centers like fit')
	})

	test('explicit left/width overrides win over rescale', async () => {
		const target = await open('mixed')
		const source = await open('image')
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')

		const shape = target.importShape(target.slides[0], source.slides[0], idx, {
			rescale: 'fit',
			left: 123456,
			width: 654321,
		})
		assertEqual(shape.left, 123456, 'left override beats rescale')
		assertEqual(shape.width, 654321, 'width override beats rescale')
	})

	test('scales a lifted table grid (gridCol@w, tr@h)', async () => {
		const target = await open('mixed') // 4:3
		const source = await open('table') // 16:9
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'graphicFrame' && s.table)
		assert(idx >= 0, 'source slide has a table')
		const srcCols = gridColWidths(source.slides[0].shapes[idx].element_)
		assert(srcCols.length > 0, 'source table has grid columns')

		const shape = target.importShape(target.slides[0], source.slides[0], idx, { rescale: 'fit' })
		const gotCols = gridColWidths(shape.element_)
		assertEqual(gotCols.length, srcCols.length, 'column count preserved')
		for (let i = 0; i < srcCols.length; i++) {
			near(gotCols[i], Math.round(srcCols[i] * 0.75), `col ${i} width scaled by 0.75`)
		}
	})

	test('still throws on a size mismatch when rescale is not requested', async () => {
		const target = await open('mixed')
		const source = await open('image')
		assert(
			throws(() => target.importShape(target.slides[0], source.slides[0], 0)),
			'a size mismatch without rescale throws'
		)
	})

	test.skipIf(!validatorInstalled)('a rescaled lifted shape stays schema-valid', async () => {
		const target = await open('mixed')
		const source = await open('image')
		const idx = findShapeIndex(source.slides[0], (s) => s.shapeType === 'picture')
		target.importShape(target.slides[0], source.slides[0], idx, { rescale: 'fit' })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
