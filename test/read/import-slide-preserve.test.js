// Tests for `Presentation.importSlide(source, index, { theme: 'preserve' })`.
//
// Contract under test: `preserve` flattens the imported slide's *source* theme
// into the slide XML (scheme colours + style-matrix fills baked to literal
// srgbClr, with colour transforms carried through) and binds the slide to this
// deck's existing master/layout — so the output is a single-theme file whose
// imported slides render with their original colours and no longer depend on
// which theme they resolve against. The default (no option) is unchanged: the
// whole source theme subgraph is copied across.

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

/** Resolve the single relationship of `type` owned by `partName`, or null. */
function resolveSingle(opc, partName, type) {
	const rels = opc.relationshipsFor(partName)
	const match = [...rels].find((r) => r.type === type)
	return match ? rels.resolveTarget(match.id) : null
}

/** The index of the first `mixed` slide that uses scheme colours + a p:style. */
const THEMED_SLIDE_INDEX = 4 // slide5: 69 schemeClr, 18 p:style (Fusion theme: accent1=00E4A8, dk2=333399)

describe("Presentation.importSlide({ theme: 'preserve' })", () => {
	test('flattens scheme colours to the resolved source RGB, carrying transforms', async () => {
		// Import mixed→mixed (equal slide size); slide5 binds to theme1 "Fusion".
		const target = await open('mixed')
		const source = await open('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await slideXml(await target.save(), imported.partName)

		assert(!/schemeClr/.test(xml), 'no a:schemeClr token remains in the flattened slide')
		assert(!/phClr/.test(xml), 'no phClr placeholder leaked from the style matrix')
		// accent1 → 00E4A8; tx2 → (clrMap tx2=dk2) → dk2 = 333399.
		assert(xml.includes('val="00E4A8"'), 'accent1 resolved to its Fusion RGB (00E4A8)')
		assert(xml.includes('val="333399"'), 'tx2 mapped through clrMap to dk2 RGB (333399)')
		// A scheme colour with a child transform keeps the transform on the literal.
		assert(/<a:srgbClr val="[0-9A-Fa-f]{6}"><a:lumMod/.test(xml), 'lumMod transform carried onto the resolved srgbClr')
	})

	test('materializes p:style fill/line/effect into spPr and neutralizes the refs', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await slideXml(await target.save(), imported.partName)

		// Every fill/line/effect ref is neutralized to idx="0" with no colour child…
		assert(/<a:fillRef idx="0"\/>/.test(xml), 'fillRef neutralized to idx="0" with no colour')
		assert(/<a:lnRef idx="0"\/>/.test(xml), 'lnRef neutralized to idx="0" with no colour')
		// …but the fontRef is left intact so its font re-binds to the destination theme.
		assert(/<a:fontRef idx="(major|minor|none)"/.test(xml), 'fontRef is preserved for the destination theme')
		// The shapes that were styled by reference now carry an explicit spPr fill.
		assert(
			/<p:spPr[ >][\s\S]*?<a:solidFill>/.test(xml) || /<a:gradFill/.test(xml),
			'an explicit spPr fill was materialized'
		)
	})

	test("carries the slide's effective background from the source master onto the slide", async () => {
		// mixed slides define no own p:bg; the master does (bgPr/solidFill schemeClr
		// bg1 → clrMap bg1=lt1 → lt1 = FFFFFF). Rebinding to the destination master
		// would otherwise drop it, so preserve must bake it onto the slide.
		const target = await open('mixed')
		const source = await open('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await slideXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(bg, 'the imported slide carries an explicit p:cSld/p:bg')
		assert(!/schemeClr/.test(bg), 'the carried background holds no scheme colour')
		assert(/<a:srgbClr val="FFFFFF"\/>/.test(bg), 'the background is the resolved literal (bg1 → lt1 = FFFFFF)')
		// p:bg must sit before p:spTree inside p:cSld.
		assert(
			/<p:cSld[^>]*>\s*<p:bg>/.test(xml) || /<p:bg>[\s\S]*?<\/p:bg>\s*<p:spTree/.test(xml),
			'p:bg precedes p:spTree'
		)
	})

	test('materializes an inherited bgRef background into an explicit fill', async () => {
		// The empty deck's master uses a theme-indexed background (bgRef idx="1001").
		// preserve must resolve it through fmtScheme into a literal bgPr fill.
		const target = await open('empty')
		const source = await open('empty')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await slideXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(bg, 'the imported slide carries an explicit background')
		assert(!/bgRef/.test(bg), 'the theme-indexed bgRef was resolved away')
		assert(!/schemeClr/.test(bg) && !/phClr/.test(bg), 'the background fill is a literal, no scheme/placeholder colour')
		assert(/<p:bgPr>[\s\S]*<a:srgbClr /.test(bg), 'bgRef became a bgPr with a literal srgbClr fill')
	})

	test('attaches to the destination master/layout without importing a new theme', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		const mastersBefore = countParts(target.opc, /\/slideMasters\/slideMaster\d+\.xml$/)

		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc

		assertEqual(countParts(opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'preserve adds no new theme part')
		assertEqual(
			countParts(opc, /\/slideMasters\/slideMaster\d+\.xml$/),
			mastersBefore,
			'preserve adds no new master part'
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

	test('preserve carries the slide media across (image slide stays intact)', async () => {
		// The image fixture's slide carries a picture; import it with preserve.
		const source = await open('image')
		const picSlide = source.slides.findIndex((s) => s.shapes.some((sh) => sh.shapeType === 'picture'))
		assert(picSlide >= 0, 'image fixture has a slide with a picture')

		const target = await open('image')
		const imported = target.importSlide(source, picSlide, { theme: 'preserve' })
		const reopened = await Presentation.load(await target.save())
		const last = reopened.slides[reopened.slides.length - 1]
		const pic = last.shapes.find((s) => s.shapeType === 'picture')
		assert(pic, 'imported slide still has its picture')
		assert(
			pic.imagePartName && reopened.opc.part(pic.imagePartName),
			`the picture's media part survives (${pic.imagePartName})`
		)
		assertEqual(imported.partName, last.partName, 'imported slide is the appended one')
	})

	test('the default (no option) still copies the source theme subgraph', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		target.importSlide(source, THEMED_SLIDE_INDEX) // default: copy
		const after = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		assert(after > themesBefore, 'the default copy mode brings a source theme across')
	})

	test.skipIf(!validatorInstalled)('a preserve-imported deck stays schema-valid', async () => {
		const target = await open('mixed')
		const source = await open('mixed')
		target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		target.importSlide(source, 5, { theme: 'preserve' }) // slide6: also themed (schemeClr + p:style)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
