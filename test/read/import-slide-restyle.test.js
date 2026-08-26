// Tests for `Presentation.importSlide(source, index, { theme: 'restyle' })`.
//
// Contract under test: `restyle` rebinds the imported slide to *this* deck's
// master/layout exactly like `preserve`, but **skips the flatten** — every
// `a:schemeClr`, style-matrix ref, and `p:bg` `bgRef` is left symbolic so it
// re-resolves against the *destination* theme and the slide re-brands. The one
// mutation is dropping the slide's own `p:clrMapOvr` (so the destination master's
// `clrMap` governs the re-brand). Literal `a:srgbClr` colours have no theme
// reference and stay byte-identical — the load-bearing limitation.
//
// `mixed` is the only 4×3 fixture, so the honest "different theme" destination is
// a synthetic recolour of `mixed`'s own theme1: slide5 binds (after rebind) to
// destination layout1 → master1 → theme1, whose accent1 we move off the source's
// 00E4A8 to a sentinel. The slide keeps its symbolic `accent1` and so adopts the
// sentinel — proof the colour re-brands rather than baking to the source RGB.

import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'
import { resolveSingle } from './opc.js'

const validatorInstalled = await validatorAvailable()

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

/** The serialized XML of a part, by partname, from saved package bytes. */
async function slideXml(bytes, partName) {
	const zip = await JSZip.loadAsync(bytes)
	const zipPath = partName.replace(/^\//, '')
	return zip.file(zipPath).async('string')
}

function countParts(opc, re) {
	return [...opc.parts.keys()].filter((n) => re.test(n)).length
}

/** The XML body of a part as a string (decoded from its bytes). */
function partText(part) {
	return new TextDecoder('utf-8').decode(part.bytes)
}

/** The index of the first `mixed` slide that uses scheme colours + a p:style + a clrMapOvr. */
const THEMED_SLIDE_INDEX = 4 // slide5: 27 schemeClr accent1, fillRef idx="1", a clrMapOvr, a literal FFFF00

/** The Fusion (theme1) accent1 RGB the source slide's `accent1` resolves to before re-brand. */
const SOURCE_ACCENT1 = '00E4A8'
/** A sentinel accent1 the destination theme1 is recoloured to, so re-brand is observable. */
const DEST_ACCENT1 = 'AABBCC'

/**
 * `mixed.pptx` with theme1's accent1 moved from the Fusion 00E4A8 to a sentinel,
 * so a slide rebound onto this deck's first master/layout (→ theme1) re-brands its
 * symbolic `accent1` to the sentinel. Returns the rebuilt package bytes.
 */
async function deckMixedRecoloredAccent1() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const theme = (await zip.file('ppt/theme/theme1.xml').async('string')).replaceAll(
		`<a:accent1><a:srgbClr val="${SOURCE_ACCENT1}"/></a:accent1>`,
		`<a:accent1><a:srgbClr val="${DEST_ACCENT1}"/></a:accent1>`
	)
	zip.file('ppt/theme/theme1.xml', theme)
	return zip.generateAsync({ type: 'uint8array' })
}

describe("Presentation.importSlide({ theme: 'restyle' })", () => {
	test('keeps a non-default PowerPoint theme source symbolic when importing into a default-theme deck', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)

		const imported = target.importSlide(source, 0, { theme: 'restyle' })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(/<a:schemeClr val="accent1"/.test(xml), 'accent1 remains symbolic for destination-theme rebinding')
		assert(/<a:schemeClr val="accent2"/.test(xml), 'accent2 remains symbolic for destination-theme rebinding')
		assert(/<a:schemeClr val="accent5"/.test(xml), 'run colour remains symbolic for destination-theme rebinding')
		assert(/<a:fillRef idx="1"><a:schemeClr val="accent1"\/><\/a:fillRef>/.test(xml), 'style fillRef remains symbolic')
		assert(!/val="B01513"/.test(xml), 'source Ion accent1 was not baked as a literal')
		assert(!/val="54849A"/.test(xml), 'source Ion accent5 was not baked as a literal')
		assertEqual(countParts(target.opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'restyle copies no source theme part')
	})

	test('re-brands a symbolic schemeClr to the destination theme (no flatten, no source theme copied)', async () => {
		// Source slide5 fills with schemeClr accent1; the destination's theme1 accent1
		// is the sentinel. restyle keeps accent1 symbolic and binds it to the
		// destination theme, so the slide adopts the sentinel — not the source 00E4A8.
		const target = await Presentation.load(await deckMixedRecoloredAccent1())
		const source = await openFixture('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)

		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' })
		const bytes = await target.save()
		const xml = await slideXml(bytes, imported.partName)

		// The scheme colour is left symbolic (the defining difference from preserve).
		assert(/<a:schemeClr val="accent1"/.test(xml), 'restyle leaves schemeClr accent1 symbolic')
		assert(!new RegExp(`val="${SOURCE_ACCENT1}"`).test(xml), 'accent1 was not baked to the source RGB')

		// No source theme was copied; the slide can only resolve against this deck's
		// theme1, whose accent1 is the sentinel.
		assertEqual(countParts(target.opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'restyle copies no source theme part')
		const reopened = await Presentation.load(bytes)
		const noSourceAccent1 = [...reopened.opc.parts.keys()]
			.filter((n) => /\/theme\/theme\d+\.xml$/.test(n))
			.every((n) => !partText(reopened.opc.part(n)).includes(SOURCE_ACCENT1))
		assert(noSourceAccent1, 'no theme in the package still carries the source accent1')
		const destTheme1 = reopened.opc.part('/ppt/theme/theme1.xml')
		assert(
			partText(destTheme1).includes(DEST_ACCENT1),
			'the destination theme1 accent1 is the sentinel the slide re-brands to'
		)
	})

	test('leaves the slide bound to the destination master/layout with style refs intact', async () => {
		// Unlike preserve (which neutralizes p:style refs to idx="0"), restyle leaves
		// fillRef/lnRef indices alone so they re-resolve against the destination fmtScheme.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		const mastersBefore = countParts(target.opc, /\/slideMasters\/slideMaster\d+\.xml$/)

		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' })
		const bytes = await target.save()
		const xml = await slideXml(bytes, imported.partName)

		assert(/<a:fillRef idx="1"/.test(xml), 'a non-zero fillRef survives (refs are not neutralized)')
		assert(/<a:lnRef idx="[12]"/.test(xml), 'a non-zero lnRef survives')

		const reopened = await Presentation.load(bytes)
		const opc = reopened.opc
		assertEqual(countParts(opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'restyle adds no new theme part')
		assertEqual(
			countParts(opc, /\/slideMasters\/slideMaster\d+\.xml$/),
			mastersBefore,
			'restyle adds no new master part'
		)

		// The imported slide binds to a layout that already exists in the target deck.
		const last = reopened.slides[reopened.slides.length - 1]
		const layout = resolveSingle(opc, last.partName, SLIDE_LAYOUT_REL)
		assert(layout && opc.part(layout), `imported slide binds to an existing destination layout (${layout})`)

		// No dangling internal relationships anywhere in the package.
		for (const partName of opc.parts.keys()) {
			if (partName.endsWith('.rels')) continue
			for (const rel of opc.relationshipsFor(partName)) {
				if (rel.targetMode === 'External') continue
				const t = opc.relationshipsFor(partName).resolveTarget(rel.id)
				assert(opc.part(t), `${partName} → ${rel.id} resolves to an existing part (${t})`)
			}
		}
		assertEqual(imported.index, reopened.slides.length - 1, 'imported slide is last')
	})

	test('drops the source slide clrMapOvr so the destination clrMap governs the re-brand', async () => {
		// slide5 carries a p:clrMapOvr/a:overrideClrMapping; restyle must remove it.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		assert(
			partText(source.opc.part(source.slides[THEMED_SLIDE_INDEX].partName)).includes('clrMapOvr'),
			'precondition: the source slide carries a clrMapOvr'
		)

		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' })
		const xml = await slideXml(await target.save(), imported.partName)
		assert(!/clrMapOvr/.test(xml), 'the imported slide has no p:clrMapOvr')
		assert(!/overrideClrMapping/.test(xml), 'the override colour mapping was removed')
	})

	test('leaves a literal srgbClr byte-identical (the limitation, as a guarantee)', async () => {
		// slide5 has a literal yellow (srgbClr FFFF00) with no theme reference. restyle
		// can only recolour symbolic colours, so this literal must survive untouched.
		const target = await Presentation.load(await deckMixedRecoloredAccent1())
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' })
		const xml = await slideXml(await target.save(), imported.partName)
		assert(/<a:srgbClr val="FFFF00"/.test(xml), 'the source literal colour is unchanged by restyle')
	})

	test('does not bake an inherited background (leaves it symbolic to re-brand)', async () => {
		// preserve bakes the slide's effective master/layout background onto the slide;
		// restyle must NOT — leaving the background to re-resolve against the destination.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'restyle' }) // slide1: no own p:bg
		const xml = await slideXml(await target.save(), imported.partName)

		const sourceHasNoBg = !/<p:bg>/.test(partText(source.opc.part(source.slides[0].partName)))
		assert(sourceHasNoBg, 'precondition: the source slide defines no own p:bg')
		assert(!/<p:bg>/.test(xml), 'restyle does not bake an inherited p:bg onto the slide')
	})

	test('carryMasterGraphics composes with restyle, leaving carried decorations symbolic', async () => {
		// mixed's slideMaster1/slideLayout1 carry non-placeholder decorations. carry bakes
		// them onto the slide; restyle leaves them (and the slide) symbolic, not flattened.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle', carryMasterGraphics: true })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(xml.includes('name="Rectangle 2"'), 'a source-master decoration was baked onto the slide')
		assert(
			xml.indexOf('name="Rectangle 2"') < xml.indexOf('<a:schemeClr val="accent1"'),
			'carried decoration precedes the slide content'
		)
		assert(/<a:schemeClr/.test(xml), 'restyle left scheme colours symbolic (carried shapes were not flattened)')
	})

	// Across two *different* decks, so the copy has something to copy: importing from a
	// deck the destination already holds byte-for-byte binds to what is there instead
	// (`ops/part-reuse.ts`), which would make a same-file import prove nothing here.
	test('the default (no option) still copies the source theme subgraph', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('theme-colors')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		target.importSlide(source, 0) // default: copy
		const after = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		assert(after > themesBefore, 'the default copy mode brings a source theme across')
	})

	test.skipIf(!validatorInstalled)('a restyle-imported deck stays schema-valid', async () => {
		const target = await Presentation.load(await deckMixedRecoloredAccent1())
		const source = await openFixture('mixed')
		target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' }) // slide5
		target.importSlide(source, 5, { theme: 'restyle' }) // slide6: also themed
		target.importSlide(source, 0, { theme: 'restyle', carryMasterGraphics: true }) // slide1 + carry
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// `remapLiterals` pushes the re-brand past what plain restyle reaches: it rewrites
// source-theme *literal* colours back to symbolic scheme colours, and copies a
// referenced source table style into the destination so a restyled table keeps it.
//
// Fixture: `multi-theme.pptx` slide 3 (PowerPoint-authored, Ion theme) carries the
// two constructs this needs — a rectangle with a literal `srgbClr B01513` equal to
// Ion accent1 (the force-remap match), a `123456` literal matching no slot (the
// negative control), and a table whose `@tableStyleId` resolves to a `<a:tblStyle>`
// in the source `tableStyles.xml`. The destination `empty.pptx` defines no such
// table style, so the copy is observable.
describe("Presentation.importSlide({ theme: 'restyle', remapLiterals: true })", () => {
	const SLIDE3 = 2
	const ION_ACCENT1 = 'B01513' // a slide3 literal equal to Ion accent1 — the force-remap match
	const NON_SLOT_LITERAL = '123456' // a slide3 literal matching no Ion slot — the negative control
	const TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' // slide3 table @tableStyleId (Medium Style 2 - Accent 1)

	test('remaps a source-theme literal to a symbolic schemeClr, leaving a non-slot literal alone', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const imported = target.importSlide(source, SLIDE3, { theme: 'restyle', remapLiterals: true })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(!new RegExp(`<a:srgbClr val="${ION_ACCENT1}"`).test(xml), 'the Ion-accent1 literal is no longer a literal')
		assert(/<a:schemeClr val="accent1"/.test(xml), 'it was rewritten to a symbolic accent1 that re-brands to this deck')
		assert(
			new RegExp(`<a:srgbClr val="${NON_SLOT_LITERAL}"`).test(xml),
			'a literal matching no source slot is left untouched'
		)
	})

	test('without the flag, plain restyle leaves the source-theme literal byte-identical', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const imported = target.importSlide(source, SLIDE3, { theme: 'restyle' })
		const xml = await slideXml(await target.save(), imported.partName)
		assert(
			new RegExp(`<a:srgbClr val="${ION_ACCENT1}"`).test(xml),
			'the literal stays put (the limitation, as a guarantee)'
		)
		assert(!/<a:schemeClr val="accent1"\/>[\s\S]*B01513/.test(xml), 'no remap happened')
	})

	test('copies the referenced source table style into this deck, without duplicating on repeat', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')

		const before = await slideXml(await target.save(), '/ppt/tableStyles.xml')
		assert(!before.includes(`styleId="${TABLE_STYLE_ID}"`), 'precondition: the destination defines no such table style')

		target.importSlide(source, SLIDE3, { theme: 'restyle', remapLiterals: true })
		target.importSlide(source, SLIDE3, { theme: 'restyle', remapLiterals: true })
		const after = await slideXml(await target.save(), '/ppt/tableStyles.xml')

		const count = after.split(`styleId="${TABLE_STYLE_ID}"`).length - 1
		assertEqual(count, 1, 'the source table style is copied exactly once across two imports')
		assert(after.includes('Medium Style 2 - Accent 1'), 'the copied definition is the source style, verbatim')
	})

	test('does not copy a table style without remapLiterals (the table falls back)', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		target.importSlide(source, SLIDE3, { theme: 'restyle' })
		const after = await slideXml(await target.save(), '/ppt/tableStyles.xml')
		assert(!after.includes(`styleId="${TABLE_STYLE_ID}"`), 'plain restyle leaves the destination tableStyles untouched')
	})

	test.skipIf(!validatorInstalled)('a remap + table-copy restyle import stays schema-valid', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		target.importSlide(source, SLIDE3, { theme: 'restyle', remapLiterals: true })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	// The remap builds two reverse indexes — RGB → slot and slot → token — and both
	// are many-to-one in the general case, because nothing makes a theme's twelve
	// `clrScheme` slots hold distinct RGBs or its twelve `clrMap` attributes name
	// distinct slots. `mixed`'s own Fusion theme already collides on the first
	// (lt1 and accent3 are both FFFFFF; dk1 and accent4 both 000000); the clrMap
	// collision is spliced in.
	const LITERAL_SHAPE =
		'<p:sp><p:nvSpPr><p:cNvPr id="88" name="Literals"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
		'<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm>' +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
		'<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>' +
		'<a:gs pos="100000"><a:srgbClr val="333399"/></a:gs></a:gsLst>' +
		'<a:lin ang="0" scaled="0"/></a:gradFill></p:spPr></p:sp>'

	/**
	 * mixed.pptx with slide1 given a shape holding two literals that match source
	 * `clrScheme` slots (FFFFFF = lt1 *and* accent3; 333399 = dk2), and its master
	 * `clrMap` re-pointed so `tx2` names `lt1` alongside `bg1` — which leaves the
	 * `dk2` slot named by no token at all. Returns package bytes.
	 */
	async function deckMixedCollidingSlotsAndTokens() {
		const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
		const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
			'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"',
			'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="lt1"'
		)
		zip.file('ppt/slideMasters/slideMaster1.xml', master)
		const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
			'</p:spTree>',
			`${LITERAL_SHAPE}</p:spTree>`
		)
		zip.file('ppt/slides/slide1.xml', slide)
		return zip.generateAsync({ type: 'uint8array' })
	}

	/**
	 * mixed.pptx with every `clrScheme` slot rewritten to `a:prstClr` — a legal
	 * `a:CT_Color` child that carries no literal RGB to read back, so the theme
	 * defines no slot → RGB mapping the remap could match a literal against.
	 * Returns package bytes.
	 */
	async function deckMixedUnreadableColorModels() {
		const slots = ['dk1', 'lt1', 'dk2', 'lt2']
			.concat([1, 2, 3, 4, 5, 6].map((n) => `accent${n}`))
			.concat(['hlink', 'folHlink'])
		const scheme = `<a:clrScheme name="Preset">${slots.map((s) => `<a:${s}><a:prstClr val="black"/></a:${s}>`).join('')}</a:clrScheme>`
		const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
		const theme = (await zip.file('ppt/theme/theme1.xml').async('string')).replace(
			/<a:clrScheme[\s\S]*?<\/a:clrScheme>/,
			scheme
		)
		zip.file('ppt/theme/theme1.xml', theme)
		const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
			'</p:spTree>',
			`${LITERAL_SHAPE}</p:spTree>`
		)
		zip.file('ppt/slides/slide1.xml', slide)
		return zip.generateAsync({ type: 'uint8array' })
	}

	test('picks the first slot on a shared RGB and the first token on a shared slot', async () => {
		const source = await Presentation.load(await deckMixedCollidingSlotsAndTokens())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'restyle', remapLiterals: true })
		const xml = await slideXml(await target.save(), imported.partName)

		// FFFFFF is both lt1 and accent3; lt1 comes first in clrScheme order. lt1 is then
		// named by both bg1 and tx2; bg1 comes first in clrMap order.
		assert(
			/<a:gs pos="0"><a:schemeClr val="bg1"\/><\/a:gs>/.test(xml),
			'FFFFFF took the first slot (lt1) and that slot the first token (bg1)'
		)
		assert(!/val="accent3"/.test(xml) && !/val="tx2"/.test(xml), 'neither runner-up won')
	})

	test('falls back to the raw slot name for a slot no clrMap token names', async () => {
		// With `tx2` re-pointed at lt1 nothing names `dk2`, so the remap has no token to
		// emit for a literal matching that slot. `dk1`/`lt1`/`dk2`/`lt2` are themselves
		// members of `ST_SchemeColorVal` and address the slot directly, bypassing the
		// clrMap — so emitting the slot name is a correct answer, not a fallback to junk.
		const source = await Presentation.load(await deckMixedCollidingSlotsAndTokens())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'restyle', remapLiterals: true })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(
			/<a:gs pos="100000"><a:schemeClr val="dk2"\/><\/a:gs>/.test(xml),
			'333399 was emitted as the direct-slot token dk2'
		)
	})

	test('remaps nothing when the source theme states its colours in a model with no literal RGB', async () => {
		// `a:prstClr` (like `a:scrgbClr`/`a:hslClr`) is a legal `a:CT_Color` child that
		// names no 6-hex RGB. A theme built from those defines no slot → RGB index for a
		// literal to match, so every literal must stay exactly as authored — the same
		// outcome as plain restyle, not a crash or a wrong match.
		const source = await Presentation.load(await deckMixedUnreadableColorModels())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'restyle', remapLiterals: true })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(/<a:gs pos="0"><a:srgbClr val="FFFFFF"\/><\/a:gs>/.test(xml), 'the FFFFFF literal is untouched')
		assert(/<a:gs pos="100000"><a:srgbClr val="333399"\/><\/a:gs>/.test(xml), 'the 333399 literal is untouched')
	})

	test.skipIf(!validatorInstalled)(
		'the colliding-slot and preset-colour sources, and their imports, stay schema-valid',
		async () => {
			for (const [name, build] of Object.entries({
				'colliding slots/tokens': deckMixedCollidingSlotsAndTokens,
				'preset-colour scheme': deckMixedUnreadableColorModels,
			})) {
				const errors = await validateBuf(Buffer.from(await build()))
				assertEqual(errors.length, 0, `${name} source: ${JSON.stringify(errors).slice(0, 2000)}`)
			}

			const target = await openFixture('mixed')
			for (const build of [deckMixedCollidingSlotsAndTokens, deckMixedUnreadableColorModels]) {
				target.importSlide(await Presentation.load(await build()), 0, { theme: 'restyle', remapLiterals: true })
			}
			const errors = await validateBuf(Buffer.from(await target.save()))
			assertEqual(errors.length, 0, `import: ${JSON.stringify(errors).slice(0, 2000)}`)
		}
	)
})
