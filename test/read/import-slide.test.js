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
import PptxGenJS from '../../dist/node.js'
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

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

/** Master partnames registered in presentation.xml's p:sldMasterIdLst (resolved via rels). */
function registeredMasters(opc) {
	const rootRels = opc.relationshipsFor('/')
	const officeDoc = [...rootRels].find((rel) => rel.type === OFFICE_DOCUMENT_REL)
	const presName = rootRels.resolveTarget(officeDoc.id)
	const root = opc.part(presName).dom.documentElement
	const rels = opc.relationshipsFor(presName)
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'sldMasterIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType !== 1 || e.localName !== 'sldMasterId') continue
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

		// The copied master must be registered in presentation.xml's
		// p:sldMasterIdLst — exactly once for a master shared across imports — or
		// renderers treat it as inactive and never paint its background/graphics.
		const masters = registeredMasters(opc)
		assert(masters.includes(ga.master), 'the copied master is registered in p:sldMasterIdLst')
		assertEqual(
			masters.filter((m) => m === ga.master).length,
			1,
			'the shared master is registered exactly once (idempotent)'
		)
	})

	test('registers each copied master in p:sldMasterIdLst (so master graphics render)', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const before = registeredMasters(target.opc).length

		target.importSlide(source, 0)
		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc

		const { master } = assertGraphResolves(opc, reopened.slides[reopened.slides.length - 1].partName)
		const masters = registeredMasters(opc)
		assertEqual(masters.length, before + 1, 'importing a slide on a new master registers one more master')
		assert(masters.includes(master), 'and the registered master is the slide’s own copied master')
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

// Slide position control (`at`): bookends need the cover first and the closer
// last regardless of import order, so `importSlide`/`cloneSlide` accept an
// `at?: number` insert position in p:sldIdLst (deck order). 0 = first; an
// out-of-range or omitted `at` appends (the prior behaviour).
describe('Presentation.importSlide({ at })', () => {
	// Tag the target's existing slides so we can recognise where the import landed.
	async function targetWithMarkedSlides() {
		const target = await open('mixed')
		const ids = target.slides.map((s) => s.slideId)
		assert(ids.length >= 1, 'mixed has at least one slide to anchor against')
		return { target, ids }
	}

	test('at: 0 inserts the imported slide first', async () => {
		const { target, ids } = await targetWithMarkedSlides()
		const source = await open('mixed')

		const imported = target.importSlide(source, 0, { at: 0 })
		assertEqual(imported.index, 0, 'imported slide reports index 0 in-memory')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, ids.length + 1, 'slide count grew by one')
		assertEqual(reopened.slides[0].slideId, imported.slideId, 'imported slide is first after round-trip')
		assertEqual(
			JSON.stringify(reopened.slides.slice(1).map((s) => s.slideId)),
			JSON.stringify(ids),
			'the original slides keep their order, shifted back by one'
		)
		assertNoDanglingRels(reopened.opc)
	})

	test('omitting at appends (unchanged behaviour)', async () => {
		const { target, ids } = await targetWithMarkedSlides()
		const source = await open('mixed')

		const imported = target.importSlide(source, 0)
		assertEqual(imported.index, ids.length, 'imported slide reports the last index')

		const reopened = await Presentation.load(await target.save())
		const reIds = reopened.slides.map((s) => s.slideId)
		assertEqual(JSON.stringify(reIds.slice(0, ids.length)), JSON.stringify(ids), 'originals stay first in order')
		assertEqual(reIds[reIds.length - 1], imported.slideId, 'imported slide is last')
	})

	test('an out-of-range at appends rather than throwing', async () => {
		const { target, ids } = await targetWithMarkedSlides()
		const source = await open('mixed')

		const imported = target.importSlide(source, 0, { at: 999 })
		assertEqual(imported.index, ids.length, 'an at past the end appends')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides[reopened.slides.length - 1].slideId, imported.slideId, 'imported slide is last')
		assertNoDanglingRels(reopened.opc)
	})

	test('cover-first + closer-last bookend placement around an interior', async () => {
		// One source deck supplies both bookends; place cover at 0 and append closer.
		const { target, ids } = await targetWithMarkedSlides()
		const source = await open('mixed')

		const cover = target.importSlide(source, 0, { at: 0 })
		const closer = target.importSlide(source, 1) // append

		const reopened = await Presentation.load(await target.save())
		const reIds = reopened.slides.map((s) => s.slideId)
		assertEqual(reIds[0], cover.slideId, 'cover is first')
		assertEqual(reIds[reIds.length - 1], closer.slideId, 'closer is last')
		assertEqual(
			JSON.stringify(reIds.slice(1, 1 + ids.length)),
			JSON.stringify(ids),
			'interior slides sit between the bookends, in order'
		)
		assertNoDanglingRels(reopened.opc)
	})

	test('cloneSlide accepts at to place the duplicate', async () => {
		const target = await open('mixed')
		const firstId = target.slides[0].slideId

		const clone = target.cloneSlide(target.slides.length - 1, { at: 0 })
		assertEqual(clone.index, 0, 'clone reports index 0')

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides[0].slideId, clone.slideId, 'clone is first after round-trip')
		assertEqual(reopened.slides[1].slideId, firstId, 'the former first slide shifted back by one')
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('an at-inserted import stays schema-valid', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		target.importSlide(source, 0, { at: 0 })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// Generate → read bridge: interior slides are authored with the generate API
// (`new PptxGenJS()`), then bookends are imported on the read/import model
// (`Presentation`). This pins that the two APIs compose: a pptxgen-generated
// package loads into Presentation, accepts an importSlide from a fixture, and
// re-saves without dangling relationships or schema errors.
describe('generate → read import bridge', () => {
	async function generatedDeckBytes() {
		// LAYOUT_WIDE is 12192000×6858000 EMU, matching the `image` read fixture so
		// importSlide's equal-size pre-flight passes (the pptxgen default is the
		// narrower LAYOUT_16x9, 9144000×5143500).
		const pres = new PptxGenJS()
		pres.layout = 'LAYOUT_WIDE'
		pres.addSlide().addText('interior slide one', { x: 1, y: 1, w: 6, h: 1 })
		pres.addSlide().addText('interior slide two', { x: 1, y: 1, w: 6, h: 1 })
		const out = await pres.stream()
		return out instanceof Uint8Array ? out : new Uint8Array(out)
	}

	test('a pptxgen-generated deck loads and accepts an imported bookend', async () => {
		const deck = await Presentation.load(await generatedDeckBytes())
		const interiorCount = deck.slides.length
		assertEqual(interiorCount, 2, 'the generated interior has two slides')

		const source = await open('image')
		const cover = deck.importSlide(source, 0, { at: 0 })

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides.length, interiorCount + 1, 'the bookend was added to the generated deck')
		assertEqual(reopened.slides[0].slideId, cover.slideId, 'imported cover is first')
		assertGraphResolves(reopened.opc, reopened.slides[0].partName)
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('the bridged deck stays schema-valid', async () => {
		const deck = await Presentation.load(await generatedDeckBytes())
		const source = await open('image')
		deck.importSlide(source, 0, { at: 0 })
		const errors = await validateBuf(Buffer.from(await deck.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
