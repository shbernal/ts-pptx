// Cross-package slide-master graft tests for `ts-pptx/read`.
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
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { throws, bytesEqual, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { fixturePath } from './corpus.js'
import { assertNoDanglingRels, resolveSingle } from './opc.js'

const validatorInstalled = await validatorAvailable()

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const THEME_REL = `${R_NS}/theme`
const OFFICE_DOCUMENT_REL = `${R_NS}/officeDocument`

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
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

/** ST_SlideMasterId / ST_SlideLayoutId floor (ECMA-376); both ids share this space. */
const ST_MASTER_LAYOUT_ID_MIN = 2147483648

/**
 * Every p:sldMasterId/@id (in presentation.xml) plus every p:sldLayoutId/@id (across
 * all masters), as raw numbers. These draw from ONE presentation-wide id space and
 * must all be unique — a duplicate makes PowerPoint report the file as corrupt.
 */
function allMasterAndLayoutIds(opc) {
	const ids = []
	const presRoot = opc.part(presentationPartName(opc)).dom.documentElement
	for (let n = presRoot.firstChild; n; n = n.nextSibling) {
		if (n.nodeType !== 1 || n.localName !== 'sldMasterIdLst') continue
		for (let e = n.firstChild; e; e = e.nextSibling) {
			if (e.nodeType === 1 && e.localName === 'sldMasterId') ids.push(Number(e.getAttribute('id')))
		}
	}
	for (const master of registeredMasters(opc)) {
		const root = opc.part(master).dom.documentElement
		for (let n = root.firstChild; n; n = n.nextSibling) {
			if (n.nodeType !== 1 || n.localName !== 'sldLayoutIdLst') continue
			for (let e = n.firstChild; e; e = e.nextSibling) {
				if (e.nodeType === 1 && e.localName === 'sldLayoutId') ids.push(Number(e.getAttribute('id')))
			}
		}
	}
	return ids
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

	test('grafted master + layout ids stay unique across the whole presentation id space', async () => {
		// Regression: p:sldMasterId/@id and every master's p:sldLayoutId/@id share ONE
		// presentation-wide id space. The destination already owns a master + layouts
		// numbered from the floor, and the source master's family is numbered the same
		// way — so a grafted master/layout that reuses its own list's floor (rather than
		// offsetting past the destination's ids) collides. PowerPoint rejects the saved
		// file as corrupt on such a duplicate; LibreOffice silently tolerates it, so this
		// invariant has to be asserted directly.
		const target = await open('empty')
		const source = await open('image')
		assert(sourceLayoutCount(source.opc) > 1, 'source has a multi-layout family (collisions are possible)')
		assert(allMasterAndLayoutIds(target.opc).length > 1, 'destination already owns master+layout ids to collide with')

		target.importSlideMasters(source)
		const opc = (await Presentation.load(await target.save())).opc

		const ids = allMasterAndLayoutIds(opc)
		assert(ids.length >= 4, `two masters with layout families contribute several ids (got ${ids.length})`)
		assertEqual(
			new Set(ids).size,
			ids.length,
			`all sldMasterId/sldLayoutId ids are unique presentation-wide: ${[...ids].sort((a, b) => a - b).join(',')}`
		)
		assert(
			ids.every((id) => Number.isInteger(id) && id >= ST_MASTER_LAYOUT_ID_MIN),
			'every id respects the ST_SlideMasterId/ST_SlideLayoutId floor'
		)
	})

	test('grafted master + layouts + theme are added; existing parts stay byte-identical', async () => {
		const input = await readFile(fixturePath('empty'))
		const target = await Presentation.load(input)
		const source = await open('image')
		target.importSlideMasters(source)

		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await target.save())
		assertUnchangedExcept(inputBodies, outputBodies, [
			'ppt/presentation.xml',
			'ppt/_rels/presentation.xml.rels',
			'[Content_Types].xml',
		])
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

// Embedded-font carry on the graft path. A master whose layouts render with an
// embedded face is only self-sufficient if the face travels with it, so
// importSlideMasters({ embedFonts: true }) carries the source deck's
// presentation-level fonts — same semantics as importSlide({ embedFonts }), which
// this shares an implementation with. Off by default: fonts live on the
// presentation, not the master, and copying them can add megabytes.
// Oracle: test/read/fixtures/embedded-fonts.pptx (PowerPoint-authored, SIL OFL
// 'Silkscreen', regular + bold) + embedded-fonts.oracle.json.
describe('Presentation.importSlideMasters({ embedFonts })', () => {
	async function graft(options) {
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlideMasters(source, options)
		return JSZip.loadAsync(await target.save())
	}

	test('carries font parts, content-type Default, rels, and a merged embeddedFontLst', async () => {
		const zip = await graft({ embedFonts: true })

		const fontParts = Object.keys(zip.files)
			.filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n))
			.sort()
		assertEqual(fontParts.length, 2, `both faces carried (got ${JSON.stringify(fontParts)})`)

		const ct = await zip.file('[Content_Types].xml').async('string')
		assert(/<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/.test(ct), 'fntdata Default added')
		assertEqual((ct.match(/x-fontdata/g) || []).length, 1, 'content type registered once (Default, no Override)')

		const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
		assertEqual(
			[...rels.matchAll(/<Relationship[^>]*\/relationships\/font"[^>]*\/>/g)].length,
			2,
			'two font relationships on presentation.xml'
		)

		const pres = await zip.file('ppt/presentation.xml').async('string')
		const lst = pres.match(/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/)?.[0]
		assert(lst, 'embeddedFontLst present')
		assert(
			/<p:font typeface="Silkscreen" pitchFamily="2" charset="0"\/>/.test(lst),
			`p:font identity carried; got ${lst}`
		)
		assert(
			/<p:regular r:id="[^"]+"\/>/.test(lst) && /<p:bold r:id="[^"]+"\/>/.test(lst),
			'regular + bold faces carried'
		)
		// embeddedFontLst sits before defaultTextStyle (CT_Presentation index 7).
		assert(
			pres.indexOf('<p:embeddedFontLst>') < pres.indexOf('<p:defaultTextStyle'),
			'embeddedFontLst precedes defaultTextStyle'
		)
	})

	test('default (flag off) carries no fonts — the graft alone is unchanged', async () => {
		const zip = await graft(undefined)
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts without embedFonts')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst without embedFonts')
	})

	test('is idempotent: grafting the same source twice carries each face once', async () => {
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlideMasters(source, { embedFonts: true })
		target.importSlideMasters(source, { embedFonts: true })

		const zip = await JSZip.loadAsync(await target.save())
		const fontParts = Object.keys(zip.files).filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n))
		assertEqual(fontParts.length, 2, 'each face copied exactly once across repeated grafts')

		const pres = await zip.file('ppt/presentation.xml').async('string')
		assertEqual(
			(pres.match(/<p:embeddedFont>/g) || []).length,
			1,
			'a single embeddedFont entry for the shared typeface'
		)
		assertEqual((pres.match(/<p:regular /g) || []).length, 1, 'regular face not duplicated')
		assertEqual((pres.match(/<p:bold /g) || []).length, 1, 'bold face not duplicated')
	})

	test('the carry adds no slide — the master stays gallery-only', async () => {
		const target = await open('empty')
		const source = await open('embedded-fonts')
		const slidesBefore = target.slides.length
		target.importSlideMasters(source, { embedFonts: true })

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, slidesBefore, 'no slide was added alongside the fonts')
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('a graft with carried embedded fonts stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlideMasters(source, { embedFonts: true })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// Table-style carry on the graft path. A grafted master's layouts arrive but the
// deck's table styling does not, so a table inserted on a grafted layout resolves
// against the destination's tableStyles.xml — for a generated deck, a stub defining
// zero styles whose default is the standard "Medium Style 2 - Accent 1". The same
// table then renders in a different accent than it would in the source deck.
// importSlideMasters({ tableStyles: true }) carries the source's whole list.
// Fixture: test/read/fixtures/table-styles.pptx (PowerPoint-authored; three tables
// styled with Microsoft built-in styles, so PowerPoint materialised four real
// a:tblStyle definitions, and a default of Medium Style 2 - Accent 3 — deliberately
// NOT the standard default, so the def carry is observable).
describe('Presentation.importSlideMasters({ tableStyles })', () => {
	const ACCENT3 = '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}' // fixture's default
	const ACCENT1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' // the standard default

	async function tableStylesXmlOf(pptxBytes) {
		const zip = await JSZip.loadAsync(pptxBytes)
		const file = zip.file('ppt/tableStyles.xml')
		return file ? file.async('string') : null
	}
	function styleIds(xml) {
		return [...xml.matchAll(/<a:tblStyle[^>]*styleId="([^"]+)"/g)].map((m) => m[1])
	}
	function defOf(xml) {
		return xml.match(/<a:tblStyleLst[^>]*\bdef="([^"]+)"/)?.[1]
	}

	test('carries every source style and the source default', async () => {
		const target = await open('empty')
		const source = await open('table-styles')

		const before = await tableStylesXmlOf(await target.save())
		assertEqual(styleIds(before).length, 0, 'precondition: destination defines no table styles')
		assertEqual(defOf(before), ACCENT1, 'precondition: destination default is the standard one')

		target.importSlideMasters(source, { tableStyles: true })
		const after = await tableStylesXmlOf(await target.save())

		const sourceIds = styleIds(await tableStylesXmlOf(await readFile(fixturePath('table-styles'))))
		assertEqual(sourceIds.length, 4, `precondition: source defines four styles (got ${sourceIds.length})`)
		for (const id of sourceIds) assert(styleIds(after).includes(id), `source style ${id} carried`)

		// The def carry is the load-bearing half: ACCENT1 is a style the source ALSO
		// defines, so keeping the destination's def would silently resolve a new table
		// to the wrong style rather than visibly to none.
		assertEqual(defOf(after), ACCENT3, 'the source default table style won')
	})

	test('default (flag off) leaves the destination table styles untouched', async () => {
		const target = await open('empty')
		const source = await open('table-styles')
		const before = await tableStylesXmlOf(await target.save())

		target.importSlideMasters(source)
		const after = await tableStylesXmlOf(await target.save())

		assertEqual(styleIds(after).length, 0, 'no styles carried without the flag')
		assertEqual(defOf(after), defOf(before), 'default table style unchanged')
	})

	test('is idempotent: grafting twice defines each style once', async () => {
		const target = await open('empty')
		const source = await open('table-styles')
		target.importSlideMasters(source, { tableStyles: true })
		const once = await tableStylesXmlOf(await target.save())

		target.importSlideMasters(source, { tableStyles: true })
		const twice = await tableStylesXmlOf(await target.save())

		const ids = styleIds(twice)
		assertEqual(new Set(ids).size, ids.length, 'no duplicate styleId across repeated grafts')
		assertEqual(ids.length, styleIds(once).length, 'a re-graft adds no styles')
		assertEqual(defOf(twice), ACCENT3, 'def stays the source default')
	})

	test('a style the destination already defines wins over the source', async () => {
		// Union is destination-wins per id, matching the embedded-font carry's de-dupe.
		const target = await open('empty')
		const source = await open('table-styles')
		target.importSlideMasters(source, { tableStyles: true })
		const first = await tableStylesXmlOf(await target.save())
		const marker = first.match(new RegExp(`<a:tblStyle[^>]*styleId="\\${ACCENT3}"[\\s\\S]*?</a:tblStyle>`))?.[0]
		assert(marker, 'the carried Accent 3 style is a full definition, not an empty stub')

		// Re-carrying must not append a second copy of an id already present.
		target.importSlideMasters(source, { tableStyles: true })
		const second = await tableStylesXmlOf(await target.save())
		assertEqual(
			(second.match(new RegExp(`styleId="\\${ACCENT3}"`, 'g')) || []).length,
			1,
			'the already-defined style was not duplicated'
		)
	})

	test('the carry adds no slide — the master stays gallery-only', async () => {
		const target = await open('empty')
		const source = await open('table-styles')
		const slidesBefore = target.slides.length
		target.importSlideMasters(source, { tableStyles: true })

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, slidesBefore, 'no slide was added alongside the table styles')
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('a graft with carried table styles stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('table-styles')
		target.importSlideMasters(source, { tableStyles: true })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// importSlideMasters({ primary: true }) moves the grafted masters to the front of
// p:sldMasterIdLst, so the deck presents as their theme (PowerPoint's Designs(1)).
// This reorders the id list only — it changes no existing slide's appearance, since
// a slide resolves its theme through its own layout's master, not through list order.
describe('Presentation.importSlideMasters({ primary })', () => {
	test('grafted master leads p:sldMasterIdLst; without the flag it trails', async () => {
		const graftedFirst = await open('empty')
		const sourceA = await open('image')
		const beforeCount = registeredMasters(graftedFirst.opc).length
		const result = graftedFirst.importSlideMasters(sourceA, { primary: true })
		assertEqual(result.length, 1, 'one master grafted')

		const withFlag = registeredMasters((await Presentation.load(await graftedFirst.save())).opc)
		assertEqual(withFlag.length, beforeCount + 1, 'the grafted master is registered')
		assertEqual(withFlag[0], result[0].partName, 'the grafted master now leads the list')

		// Same graft without the flag: the grafted master appends after the original.
		const appended = await open('empty')
		const sourceB = await open('image')
		const trailing = appended.importSlideMasters(sourceB)
		const withoutFlag = registeredMasters((await Presentation.load(await appended.save())).opc)
		assertEqual(
			withoutFlag[withoutFlag.length - 1],
			trailing[0].partName,
			'without the flag the grafted master trails, confirming the flag is what moved it'
		)
	})

	test('reorders the id list only — every other part stays byte-identical', async () => {
		const input = await readFile(fixturePath('empty'))
		const withPrimary = await Presentation.load(input)
		const source = await open('image')
		withPrimary.importSlideMasters(source, { primary: true })

		// A plain graft (append) and a primary graft from the same inputs differ in
		// exactly one part: presentation.xml, where the id list is reordered.
		const plain = await Presentation.load(input)
		plain.importSlideMasters(await open('image'))

		const primaryBodies = await partBodies(await withPrimary.save())
		const plainBodies = await partBodies(await plain.save())
		const differing = [...primaryBodies.keys()].filter(
			(name) => !bytesEqual(primaryBodies.get(name), plainBodies.get(name))
		)
		assertEqual(
			differing.join(','),
			'ppt/presentation.xml',
			'primary vs append differ in presentation.xml alone (id-list order)'
		)
	})

	test('is idempotent: a second primary graft does not reshuffle', async () => {
		const target = await open('empty')
		const source = await open('image')
		target.importSlideMasters(source, { primary: true })
		const once = registeredMasters((await Presentation.load(await target.save())).opc)

		// Re-carrying the SAME source is a copy-registry no-op; the promotion must not
		// reorder the already-leading master, and it must not double-register.
		target.importSlideMasters(source, { primary: true })
		const twice = registeredMasters((await Presentation.load(await target.save())).opc)
		assertEqual(twice.join('|'), once.join('|'), 'the master list is unchanged by a repeat primary graft')
	})

	test('the promotion adds no slide and leaves no dangling rels', async () => {
		const target = await open('empty')
		const source = await open('image')
		const slidesBefore = target.slides.length
		target.importSlideMasters(source, { primary: true })

		const reopened = await Presentation.load(await target.save())
		assertEqual(reopened.slides.length, slidesBefore, 'no slide was added by promoting the master')
		assertNoDanglingRels(reopened.opc)
	})

	test.skipIf(!validatorInstalled)('a primary graft stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('image')
		target.importSlideMasters(source, { primary: true })
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
		// equal-size guard passes (ts-pptx's default is the narrower LAYOUT_16x9).
		const pres = new TsPptx()
		pres.layout = 'LAYOUT_WIDE'
		pres.addSlide().addText('interior one', { x: 1, y: 1, w: 6, h: 1 })
		pres.addSlide().addText('interior two', { x: 1, y: 1, w: 6, h: 1 })
		const out = await pres.stream()
		return out instanceof Uint8Array ? out : new Uint8Array(/** @type {ArrayBuffer} */ (out))
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
