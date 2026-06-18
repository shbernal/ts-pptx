// Slide-removal tests for `pptxgenjs/read`.
//
// Contract under test: Presentation.removeSlide(index) drops the p:sldId entry,
// the presentation→slide relationship, the slide part + its .rels, and any part
// the slide privately owned (notes/media) that nothing else references — but
// never shared chrome (layout/master/theme). Removing every slide yields a valid
// master/layout-only template shell. Round-trips with no dangling rels and stays
// schema-valid; untouched parts stay byte-identical.

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

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const SLIDE_MASTER_REL = `${R_NS}/slideMaster`
const SLIDE_LAYOUT_REL = `${R_NS}/slideLayout`

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
function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const matches = [...rels].filter((rel) => rel.type === type)
	return matches.length === 0 ? null : rels.resolveTarget(matches[0].id)
}
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
function partNames(opc) {
	return new Set([...opc.parts.keys()])
}

describe('Presentation.removeSlide', () => {
	test('removes a slide, its part, and its rel; survives a round-trip', async () => {
		const deck = await open('mixed')
		const before = deck.slides.length
		assert(before > 1, 'fixture has multiple slides')
		const removedId = deck.slides[0].slideId
		const removedPart = deck.removeSlide(0)

		assertEqual(deck.slides.length, before - 1, 'one fewer slide in-memory')
		assert(!deck.slides.some((s) => s.slideId === removedId), 'the removed slide id is gone')

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides.length, before - 1, 'slide count stays dropped after reload')
		assert(!reopened.opc.part(removedPart), 'the slide part is gone from the package')
		assert(!reopened.opc.part(`${removedPart.replace(/\/([^/]+)$/, '/_rels/$1')}.rels`), 'its .rels is gone too')
		assertNoDanglingRels(reopened.opc)
	})

	test('keeps shared chrome (layout/master/theme) when a slide is removed', async () => {
		const deck = await open('mixed')
		const layout = resolveSingle(deck.opc, deck.slides[0].partName, SLIDE_LAYOUT_REL)
		const master = resolveSingle(deck.opc, layout, SLIDE_MASTER_REL)
		deck.removeSlide(0)
		const reopened = await Presentation.load(await deck.save())
		assert(reopened.opc.part(layout), 'the slide layout is preserved')
		assert(reopened.opc.part(master), 'the slide master is preserved')
		assertNoDanglingRels(reopened.opc)
	})

	test('removing every slide yields a valid master/layout-only shell', async () => {
		const deck = await open('image')
		const layoutCount = [...partNames(deck.opc)].filter((n) => /slideLayouts\/slideLayout\d+\.xml$/.test(n)).length
		while (deck.slides.length) deck.removeSlide(0)
		assertEqual(deck.slides.length, 0, 'no slides remain in-memory')

		const reopened = await Presentation.load(await deck.save())
		assertEqual(reopened.slides.length, 0, 'the saved shell has zero slides')
		const remaining = [...partNames(reopened.opc)]
		assertEqual(
			remaining.filter((n) => /slideLayouts\/slideLayout\d+\.xml$/.test(n)).length,
			layoutCount,
			'all layouts are retained in the shell'
		)
		assert(
			remaining.some((n) => /slideMasters\/slideMaster\d+\.xml$/.test(n)),
			'the master is retained'
		)
		assert(!remaining.some((n) => /ppt\/slides\/slide\d+\.xml$/.test(n)), 'no slide parts remain')
		assertNoDanglingRels(reopened.opc)
	})

	test('untouched parts stay byte-identical after a removal', async () => {
		const input = await readFile(fixturePath('mixed'))
		const deck = await Presentation.load(input)
		deck.removeSlide(0)

		const inBodies = new Map()
		for (const e of Object.values((await JSZip.loadAsync(input)).files))
			if (!e.dir) inBodies.set(e.name, await e.async('uint8array'))
		const outBodies = new Map()
		for (const e of Object.values((await JSZip.loadAsync(await deck.save())).files))
			if (!e.dir) outBodies.set(e.name, await e.async('uint8array'))

		const allowedToChange = new Set(['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels', '[Content_Types].xml'])
		for (const [name, body] of inBodies) {
			if (allowedToChange.has(name) || !outBodies.has(name)) continue // removed parts are expected gone
			const out = outBodies.get(name)
			assert(out && out.length === body.length && out.every((v, i) => v === body[i]), `${name} should be untouched`)
		}
	})

	test('rejects an out-of-range index', async () => {
		const deck = await open('mixed')
		assert(
			throws(() => deck.removeSlide(999)),
			'removing a missing slide throws'
		)
	})

	test.skipIf(!validatorInstalled)('a master/layout-only shell stays schema-valid', async () => {
		const deck = await open('image')
		while (deck.slides.length) deck.removeSlide(0)
		const errors = await validateBuf(Buffer.from(await deck.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
