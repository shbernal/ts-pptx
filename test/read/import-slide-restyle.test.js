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

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

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

/** Resolve the single relationship of `type` owned by `partName`, or null. */
function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const match = [...rels].find((r) => r.type === type)
	return match ? rels.resolveTarget(match.id) : null
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
	test('re-brands a symbolic schemeClr to the destination theme (no flatten, no source theme copied)', async () => {
		// Source slide5 fills with schemeClr accent1; the destination's theme1 accent1
		// is the sentinel. restyle keeps accent1 symbolic and binds it to the
		// destination theme, so the slide adopts the sentinel — not the source 00E4A8.
		const target = await Presentation.load(await deckMixedRecoloredAccent1())
		const source = await open('mixed')
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
		const target = await open('mixed')
		const source = await open('mixed')
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
		const target = await open('mixed')
		const source = await open('mixed')
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
		const source = await open('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' })
		const xml = await slideXml(await target.save(), imported.partName)
		assert(/<a:srgbClr val="FFFF00"/.test(xml), 'the source literal colour is unchanged by restyle')
	})

	test('does not bake an inherited background (leaves it symbolic to re-brand)', async () => {
		// preserve bakes the slide's effective master/layout background onto the slide;
		// restyle must NOT — leaving the background to re-resolve against the destination.
		const target = await open('mixed')
		const source = await open('mixed')
		const imported = target.importSlide(source, 0, { theme: 'restyle' }) // slide1: no own p:bg
		const xml = await slideXml(await target.save(), imported.partName)

		const sourceHasNoBg = !/<p:bg>/.test(partText(source.opc.part(source.slides[0].partName)))
		assert(sourceHasNoBg, 'precondition: the source slide defines no own p:bg')
		assert(!/<p:bg>/.test(xml), 'restyle does not bake an inherited p:bg onto the slide')
	})

	test('carryMasterGraphics composes with restyle, leaving carried decorations symbolic', async () => {
		// mixed's slideMaster1/slideLayout1 carry non-placeholder decorations. carry bakes
		// them onto the slide; restyle leaves them (and the slide) symbolic, not flattened.
		const target = await open('mixed')
		const source = await open('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle', carryMasterGraphics: true })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(xml.includes('name="Rectangle 2"'), 'a source-master decoration was baked onto the slide')
		assert(
			xml.indexOf('name="Rectangle 2"') < xml.indexOf('<a:schemeClr val="accent1"'),
			'carried decoration precedes the slide content'
		)
		assert(/<a:schemeClr/.test(xml), 'restyle left scheme colours symbolic (carried shapes were not flattened)')
	})

	test('the default (no option) still copies the source theme subgraph', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		target.importSlide(source, THEMED_SLIDE_INDEX) // default: copy
		const after = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		assert(after > themesBefore, 'the default copy mode brings a source theme across')
	})

	test.skipIf(!validatorInstalled)('a restyle-imported deck stays schema-valid', async () => {
		const target = await Presentation.load(await deckMixedRecoloredAccent1())
		const source = await open('mixed')
		target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'restyle' }) // slide5
		target.importSlide(source, 5, { theme: 'restyle' }) // slide6: also themed
		target.importSlide(source, 0, { theme: 'restyle', carryMasterGraphics: true }) // slide1 + carry
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
