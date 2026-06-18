// Cross-package slide-master graft tests for `pptxgenjs/read`.
//
// Contract under test: Presentation.importSlideMasters(source) copies master(s)
// from a *different* open package together with their WHOLE layout family (not
// just the layouts some slide uses, as importSlide does), attaches them to no
// slide, registers each in p:sldMasterIdLst, rebuilds each master's
// p:sldLayoutIdLst to exactly the copied layouts, brings the theme/media across
// under fresh partnames, survives a save → reopen with no dangling rels, leaves
// untouched parts byte-identical, and stays schema-valid. The masters/layouts
// filters narrow what is grafted; re-calls are idempotent; a slide-size mismatch
// is rejected unless explicitly overridden.

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

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const THEME_REL = `${R_NS}/theme`
const OFFICE_DOCUMENT_REL = `${R_NS}/officeDocument`

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

function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	return matches.length === 0 ? null : rels.resolveTarget(matches[0].id)
}

function presentationPartName(opc) {
	const rootRels = opc.relationshipsFor('/')
	const officeDoc = [...rootRels].find((rel) => rel.type === OFFICE_DOCUMENT_REL)
	return rootRels.resolveTarget(officeDoc.id)
}

/** Master partnames registered in presentation.xml's p:sldMasterIdLst (resolved via rels). */
function registeredMasters(opc) {
	const presName = presentationPartName(opc)
	const root = opc.part(presName).dom.documentElement
	const rels = opc.relationshipsFor(presName)
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'sldMasterIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType !== 1 || e.localName !== 'sldMasterId') continue
			out.push(rels.resolveTarget(e.getAttributeNS(R_NS, 'id')))
		}
	}
	return out
}

/** Partnames the master lists in p:sldLayoutIdLst, resolved via the master's rels. */
function masterLayoutList(opc, masterPartName) {
	const root = opc.part(masterPartName).dom.documentElement
	const rels = opc.relationshipsFor(masterPartName)
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'sldLayoutIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType !== 1 || e.localName !== 'sldLayoutId') continue
			out.push(rels.resolveTarget(e.getAttributeNS(R_NS, 'id')))
		}
	}
	return out
}

/** Count source layouts on the first registered master of a package. */
function sourceLayoutCount(opc) {
	return masterLayoutList(opc, registeredMasters(opc)[0]).length
}

/** Every internal relationship of every part resolves to a part that exists. */
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

describe('Presentation.importSlideMasters', () => {
	test('grafts a master with its WHOLE layout family, attached to no slide', async () => {
		const target = await open('empty')
		const source = await open('image')
		const slidesBefore = target.slides.length
		const mastersBefore = registeredMasters(target.opc).length
		const familySize = sourceLayoutCount(source.opc)
		assert(familySize > 1, 'source master has a multi-layout family to graft')

		const result = target.importSlideMasters(source)
		assertEqual(result.length, 1, 'one master was grafted')
		assertEqual(result[0].layoutPartNames.length, familySize, 'all source layouts came across (not just used ones)')

		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc
		assertEqual(reopened.slides.length, slidesBefore, 'no slide was added — the master is gallery-only')

		const masters = registeredMasters(opc)
		assertEqual(masters.length, mastersBefore + 1, 'the grafted master is registered in p:sldMasterIdLst')
		const grafted = masters[masters.length - 1] // registerMaster appends
		const listed = masterLayoutList(opc, grafted)
		assertEqual(listed.length, familySize, 'the grafted master lists its full layout family')
		assertEqual(new Set(listed).size, listed.length, 'with no duplicate layout entries')

		// The grafted master resolves to a theme, and every listed layout exists.
		assert(resolveSingle(opc, grafted, THEME_REL), 'grafted master carries a theme')
		for (const layout of listed) assert(opc.part(layout), `listed layout exists (${layout})`)
		assertNoDanglingRels(opc)
	})

	test('grafted master + layouts + theme are added; existing parts stay byte-identical', async () => {
		const input = await readFile(fixturePath('empty'))
		const target = await Presentation.load(input)
		const source = await open('image')
		target.importSlideMasters(source)

		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await target.save())
		const allowedToChange = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, body] of inputBodies) {
			if (allowedToChange.has(name)) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
		const added = [...outputBodies.keys()].filter((name) => !inputBodies.has(name))
		assert(
			added.some((n) => /ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n)),
			'a master part was added'
		)
		assert(
			added.filter((n) => /ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).length > 1,
			'multiple layout parts were added'
		)
		assert(
			added.some((n) => /ppt\/theme\/theme\d+\.xml$/.test(n)),
			'a theme part was added'
		)
		assert(!added.some((n) => /ppt\/slides\/slide\d+\.xml$/.test(n)), 'no slide part was added')
	})

	test('layouts filter grafts only the chosen subset', async () => {
		const target = await open('empty')
		const source = await open('image')
		const result = target.importSlideMasters(source, { layouts: (_name, index) => index < 3 })
		assertEqual(result[0].layoutPartNames.length, 3, 'only the first three layouts were grafted')

		const reopened = await Presentation.load(await target.save())
		const grafted = registeredMasters(reopened.opc).pop()
		assertEqual(masterLayoutList(reopened.opc, grafted).length, 3, 'the grafted master lists exactly the subset')
		assertNoDanglingRels(reopened.opc)
	})

	test('masters filter selects which masters to graft', async () => {
		const target = await open('empty')
		const source = await open('image')
		const none = target.importSlideMasters(source, { masters: () => false })
		assertEqual(none.length, 0, 'no master matched, nothing grafted')
		assertEqual(registeredMasters(target.opc).length, 1, 'destination master count unchanged')
	})

	test('re-grafting the same source is idempotent (no duplicate layouts/masters)', async () => {
		const target = await open('empty')
		const source = await open('image')
		target.importSlideMasters(source)
		const afterFirst = registeredMasters(target.opc).length
		const familySize = sourceLayoutCount(source.opc)

		target.importSlideMasters(source)
		const reopened = await Presentation.load(await target.save())
		assertEqual(registeredMasters(reopened.opc).length, afterFirst, 'a second graft adds no new master')
		const grafted = registeredMasters(reopened.opc).pop()
		assertEqual(masterLayoutList(reopened.opc, grafted).length, familySize, 'and no duplicate layout entries')
		assertNoDanglingRels(reopened.opc)
	})

	test('rejects a slide-size mismatch unless overridden', async () => {
		const target = await open('empty') // 16:9
		const source = await open('mixed') // 4:3
		assert(
			throws(() => target.importSlideMasters(source)),
			'mismatched sizes throw by default'
		)
		const result = target.importSlideMasters(source, { requireEqualSize: false })
		assert(result.length >= 1, 'override grafts despite the size mismatch')
		assertNoDanglingRels(target.opc)
	})

	test.skipIf(!validatorInstalled)('a deck with a grafted master stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('image')
		target.importSlideMasters(source)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// Generate → read bridge: the real use case. Interior slides are authored with
// the generate API; the brand master is then grafted in on the read/import model
// so the generated deck ships the template's layout gallery without applying it.
describe('generate → read slide-master graft bridge', () => {
	async function generatedDeckBytes() {
		// LAYOUT_WIDE (12192000×6858000 EMU) matches the `image` fixture so the
		// equal-size guard passes (pptxgen's default is the narrower LAYOUT_16x9).
		const pres = new PptxGenJS()
		pres.layout = 'LAYOUT_WIDE'
		pres.addSlide().addText('interior one', { x: 1, y: 1, w: 6, h: 1 })
		pres.addSlide().addText('interior two', { x: 1, y: 1, w: 6, h: 1 })
		const out = await pres.stream()
		return out instanceof Uint8Array ? out : new Uint8Array(out)
	}

	test('a generated deck ships a grafted master without changing its slides', async () => {
		const deck = await Presentation.load(await generatedDeckBytes())
		const slidesBefore = deck.slides.length
		const mastersBefore = registeredMasters(deck.opc).length
		const source = await open('image')

		deck.importSlideMasters(source)
		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides.length, slidesBefore, 'generated slides are untouched')
		assertEqual(registeredMasters(reopened.opc).length, mastersBefore + 1, 'the brand master was added to the gallery')
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('the bridged deck stays schema-valid', async () => {
		const deck = await Presentation.load(await generatedDeckBytes())
		const source = await open('image')
		deck.importSlideMasters(source)
		const errors = await validateBuf(Buffer.from(await deck.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
