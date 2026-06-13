// Phase 4 cross-package slide-import tests for `pptxgenjs/read`.
//
// Contract under test: Presentation.importSlide(source, index) appends a copy of
// a slide from a *different* open package, bringing its layout → master → theme
// and any media/chart/embedding parts under fresh partnames; prunes the imported
// master's p:sldLayoutIdLst to the layout(s) actually used; survives a save →
// reopen round-trip with no dangling relationships; leaves untouched parts of the
// target package byte-identical; and keeps the package schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

function bytesEqual(a, b) {
	return a && b && a.length === b.length && a.every((value, index) => value === b[index])
}

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

/** Resolve the single relationship of `type` owned by `partName` to a target partname, or null. */
function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	if (matches.length === 0) return null
	return rels.resolveTarget(matches[0].id)
}

/** Walk a slide's layout → master → theme chain, asserting every hop targets an existing part. */
function assertGraphResolves(opc, slidePartName) {
	const layout = resolveSingle(opc, slidePartName, SLIDE_LAYOUT_REL)
	assert(layout && opc.part(layout), `slide ${slidePartName} resolves to an existing layout (${layout})`)
	const master = resolveSingle(opc, layout, SLIDE_MASTER_REL)
	assert(master && opc.part(master), `layout ${layout} resolves to an existing master (${master})`)
	const theme = resolveSingle(opc, master, THEME_REL)
	assert(theme && opc.part(theme), `master ${master} resolves to an existing theme (${theme})`)
	return { layout, master, theme }
}

/** Partnames the master lists in p:sldLayoutIdLst, resolved via the master's rels. */
function masterLayoutList(opc, masterPartName) {
	const part = opc.part(masterPartName)
	const root = part.dom.documentElement
	const rels = opc.relationshipsFor(masterPartName)
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'sldLayoutIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType !== 1 || e.localName !== 'sldLayoutId') continue
			const relId = e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
			out.push(rels.resolveTarget(relId))
		}
	}
	return out
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

describe('Presentation.importSlide', () => {
	test('imports an image slide into another deck; the graph + media survive a round-trip', async () => {
		const target = await open('empty')
		const source = await open('image')
		const beforeCount = target.slides.length
		const sourcePicCount = source.slides[0].shapes.filter((s) => s.shapeType === 'picture').length
		assert(sourcePicCount > 0, 'source slide actually has a picture to carry')

		const imported = target.importSlide(source, 0)
		assertEqual(target.slides.length, beforeCount + 1, 'a slide was appended in-memory')
		assertEqual(imported.index, beforeCount, 'imported slide is last')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, beforeCount + 1, 'slide count grew after reload')
		const ids = reopened.slides.map((s) => s.slideId)
		assertEqual(new Set(ids).size, ids.length, 'slide ids are unique')

		const last = reopened.slides[reopened.slides.length - 1]
		assertEqual(
			last.shapes.filter((s) => s.shapeType === 'picture').length,
			sourcePicCount,
			'imported pictures survive'
		)

		const opc = reopened.opc
		const { layout, master } = assertGraphResolves(opc, last.partName)
		assertNoDanglingRels(opc)

		// The imported picture's image rel resolves to a media part that exists.
		const pic = last.shapes.find((s) => s.shapeType === 'picture')
		assert(pic.imagePartName && opc.part(pic.imagePartName), `imported media present (${pic.imagePartName})`)

		// The imported master lists exactly the one copied layout.
		const listed = masterLayoutList(opc, master)
		assertEqual(listed.length, 1, 'imported master lists exactly one layout')
		assertEqual(listed[0], layout, 'and it is the layout the imported slide uses')
	})

	test('imports a table slide into another deck; the table survives a round-trip', async () => {
		const target = await open('empty')
		const source = await open('table')
		const sourceTableSlide = source.slides.findIndex((s) =>
			s.shapes.some((sh) => sh.shapeType === 'graphicFrame' && sh.table)
		)
		assert(sourceTableSlide >= 0, 'source deck has a table slide')

		const imported = target.importSlide(source, sourceTableSlide)
		const reopened = await Presentation.load(await target.save())
		const last = reopened.slides[reopened.slides.length - 1]
		const table = last.shapes.find((s) => s.shapeType === 'graphicFrame' && s.table)?.table
		assert(table, 'imported slide still has a table')
		assertGraphResolves(reopened.opc, last.partName)
		assertNoDanglingRels(reopened.opc)
		assert(imported.index === reopened.slides.length - 1, 'imported slide is last')
	})

	test('only the presentation part + its rels + content types change; imported parts are added', async () => {
		const input = await readFile(fixturePath('empty'))
		const target = await Presentation.load(input)
		const source = await open('image')
		target.importSlide(source, 0)
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await target.save())

		const allowedToChange = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, body] of inputBodies) {
			if (allowedToChange.has(name)) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
		const added = [...outputBodies.keys()].filter((name) => !inputBodies.has(name))
		// A slide, a layout, a master, a theme, and the image media (+ their rels) are added.
		assert(
			added.some((n) => /ppt\/slides\/slide\d+\.xml$/.test(n)),
			`a slide part was added: ${JSON.stringify(added)}`
		)
		assert(
			added.some((n) => /ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)),
			'a layout part was added'
		)
		assert(
			added.some((n) => /ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n)),
			'a master part was added'
		)
		assert(
			added.some((n) => /ppt\/theme\/theme\d+\.xml$/.test(n)),
			'a theme part was added'
		)
		assert(
			added.some((n) => /ppt\/media\//.test(n)),
			'a media part was added'
		)
	})

	test('dedups a shared master across imports and accumulates its used layouts', async () => {
		// mixed slide 1 uses layout1, slide 2 uses layout2 — both on the same master.
		const target = await open('mixed')
		const source = await open('mixed')
		const beforeCount = target.slides.length

		target.importSlide(source, 0)
		target.importSlide(source, 1)

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, beforeCount + 2, 'two slides were appended')
		assertNoDanglingRels(reopened.opc)

		const opc = reopened.opc
		const a = reopened.slides[reopened.slides.length - 2]
		const b = reopened.slides[reopened.slides.length - 1]
		const ga = assertGraphResolves(opc, a.partName)
		const gb = assertGraphResolves(opc, b.partName)

		assertEqual(ga.master, gb.master, 'the two imported slides share one copied master (deduped)')
		assert(ga.layout !== gb.layout, 'but they use two distinct copied layouts')

		const listed = masterLayoutList(opc, ga.master).sort()
		assertEqual(listed.length, 2, 'the shared master lists exactly the two used layouts')
		assertEqual(
			JSON.stringify(listed),
			JSON.stringify([ga.layout, gb.layout].sort()),
			'and they are precisely those two'
		)
	})

	test('imports a chart slide with its embedded data', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const chartSlide = source.slides.findIndex((s) =>
			s.shapes.some((sh) => sh.shapeType === 'graphicFrame' && sh.chart)
		)
		assert(chartSlide >= 0, 'mixed has a chart slide')

		target.importSlide(source, chartSlide)
		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc
		assertNoDanglingRels(opc)

		// A chart part and an embedded workbook were copied in.
		const partNames = [...opc.parts.keys()]
		const chartsBefore = [...(await open('mixed')).opc.parts.keys()].filter((n) =>
			/\/charts\/chart\d+\.xml$/.test(n)
		).length
		const chartsAfter = partNames.filter((n) => /\/charts\/chart\d+\.xml$/.test(n)).length
		assert(chartsAfter > chartsBefore, 'a chart part was added by the import')
		assert(
			partNames.some((n) => /\/embeddings\//.test(n)),
			'the chart embedding part is present'
		)
	})

	test('rejects a slide-size mismatch between source and target', async () => {
		const target = await open('empty') // 16:9
		const source = await open('mixed') // 4:3
		assert(
			throws(() => target.importSlide(source, 0)),
			'importing across mismatched slide sizes throws'
		)
	})

	test('rejects an out-of-range index', async () => {
		const target = await open('empty')
		const source = await open('image')
		assert(
			throws(() => target.importSlide(source, 99)),
			'importing a missing slide throws'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with an imported slide stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('image')
		target.importSlide(source, 0)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test.skipIf(!validatorInstalled)('a deck with two imported same-master slides stays schema-valid', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		target.importSlide(source, 0)
		target.importSlide(source, 1)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
