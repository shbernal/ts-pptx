// Integration regression for the read model against a single deck that combines,
// in one PowerPoint-authored package, the stress dimensions that individually
// live in narrower fixtures but rarely co-occur: two live slide masters each with
// its own theme, multi-typeface embedded fonts, blip recolor, an SVG picture,
// nested groups, table-style cell-fill resolution + inline scheme fill with lumMod,
// threaded modern comments, and speaker notes.
//
// read-stress.pptx is authored by desktop Microsoft PowerPoint (see
// .tmp/author-read-stress.ps1 and fixtures/README.md). Its job is to prove the
// resolution chains keep working *together* on a non-trivial deck — a guard that
// the single-dimension fixtures cannot give.
//
// Deliberately out of this fixture (real PowerPoint COM cannot author them
// head­less; each is covered off-fixture): duotone/clrChange/alphaModFix recolor
// (picture-recolor.test.js), the 'both' raster+SVG mediaKind (style-accessors),
// and hdphoto/.wdp artistic-effect layers.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

let cached
async function pres() {
	if (!cached) cached = Presentation.load(await readFile(path.join(__dirname, 'fixtures', 'read-stress.pptx')))
	return cached
}

/** Flatten a shape list, descending into groups. */
function allShapes(shapes) {
	return shapes.flatMap((s) => (s.shapeType === 'group' ? [s, ...allShapes(s.shapes)] : [s]))
}
async function everyShape() {
	const p = await pres()
	return p.slides.flatMap((s) => allShapes(s.shapes))
}
function named(shapes, name) {
	const s = shapes.find((x) => x.name === name)
	assert(s, `expected a shape named ${name}`)
	return s
}

describe('read-stress.pptx — combined read-model integration', () => {
	test('two slide masters, each with its own theme, both live', async () => {
		const p = await pres()
		assertEqual(p.masters().length, 2, 'exactly two slide masters')
		const themeNames = p.masters().map((m) => m.theme?.name)
		assert(new Set(themeNames).size === 2, `master themes should be distinct, got ${themeNames}`)

		// Each slide resolves to a different master + theme (both masters are used).
		const [s1, s2] = p.slides
		assert(s1.master && s2.master, 'both slides resolve a master')
		assert(
			s1.master.partName !== s2.master.partName,
			`slides should resolve to different masters, both ${s1.master.partName}`
		)
		assert(s1.theme?.name !== s2.theme?.name, `slides should resolve to different themes, both ${s1.theme?.name}`)
	})

	test('multi-typeface embedded fonts with faces', async () => {
		const p = await pres()
		const byName = new Map(p.embeddedFonts.map((f) => [f.typeface, f]))
		for (const tf of ['Georgia', 'Consolas', 'Trebuchet MS']) {
			const f = byName.get(tf)
			assert(f, `expected ${tf} among embedded fonts (${[...byName.keys()]})`)
			assert(f.faces.length >= 1, `${tf} should carry at least one face`)
			for (const face of f.faces) {
				assert(face.slot, `${tf} face should name a slot`)
				assert(/\/ppt\/fonts\/.+\.fntdata$/.test(face.partName), `${tf} face should point at a .fntdata part`)
			}
		}
		assert(p.embeddedFonts.length >= 3, 'at least three embedded typefaces')
	})

	test('blip recolor: grayscale and biLevel both resolve', async () => {
		const shapes = await everyShape()
		const gray = named(shapes, 'GrayPic')
		assertEqual(gray.recolor?.kind, 'grayscale', 'ColorType grayscale -> grayscale recolor')
		const bi = named(shapes, 'BiLevelPic')
		assertEqual(bi.recolor?.kind, 'biLevel', 'ColorType black&white -> biLevel recolor')
	})

	test('SVG picture resolves as svg-only with an svg part', async () => {
		const shapes = await everyShape()
		const svg = named(shapes, 'SvgIcon')
		// Real PowerPoint COM writes an svg-only blip (empty a:blip + asvg:svgBlip).
		assertEqual(svg.mediaKind, 'svg', 'inserted .svg is svg-only')
		assert(/\.svg$/.test(svg.svgPartName ?? ''), `svgPartName should resolve, got ${svg.svgPartName}`)
	})

	test('nested groups: a group inside a group', async () => {
		const p = await pres()
		const outer = allShapes(p.slides[0].shapes).find((s) => s.name === 'OuterGroup')
		assert(outer && outer.shapeType === 'group', 'OuterGroup is a group')
		const inner = outer.shapes.find((s) => s.shapeType === 'group')
		assert(inner, 'OuterGroup should directly contain a nested group')
		assert(inner.shapes.length >= 2, 'the inner group should contain its child shapes')
	})

	test('styled table: cells with no own fill resolve fill from the table style', async () => {
		const shapes = await everyShape()
		const gf = named(shapes, 'StyledTable')
		assert(gf.hasTable, 'StyledTable is a table')
		const tbl = gf.table
		assert(tbl.styleId, 'table carries a style id')
		let resolvedFromStyle = 0
		for (const row of tbl.rows) {
			for (const c of row.cells) {
				const ownFill = c.fillColor || c.fillSchemeColor
				if (!ownFill && c.resolvedFill?.effectiveHex) resolvedFromStyle++
			}
		}
		assert(resolvedFromStyle > 0, 'at least one no-own-fill cell resolves a fill from the table style')
	})

	test('inline scheme fill with lumMod resolves through the transform chain', async () => {
		const shapes = await everyShape()
		const gf = named(shapes, 'LumModTable')
		let found = false
		for (const row of gf.table.rows) {
			for (const c of row.cells) {
				if (c.fillSchemeColor === 'accent4' && c.resolvedFill?.transforms?.some((t) => t.name === 'lumMod')) {
					found = true
					assert(c.resolvedFill.effectiveHex, 'the lumMod cell resolves an effective hex')
				}
			}
		}
		assert(found, 'a cell has an accent4 scheme fill carrying a lumMod transform')
	})

	test('threaded modern comments with two authors', async () => {
		const p = await pres()
		const authors = p.modernCommentAuthors.map((a) => a.name)
		assert(authors.length >= 2, `expected >=2 modern comment authors, got ${authors}`)

		const comments = p.slides.flatMap((s) => s.modernComments)
		assert(comments.length >= 2, `expected >=2 top-level modern comments, got ${comments.length}`)
		const threaded = comments.find((c) => (c.replies?.length ?? 0) >= 1)
		assert(threaded, 'at least one comment has replies')

		const inThread = new Set()
		inThread.add(threaded.author)
		for (const r of threaded.replies) inThread.add(r.author)
		assert(inThread.size >= 2, `a reply thread should involve >=2 distinct authors, got ${[...inThread]}`)
	})

	test('speaker notes are extracted', async () => {
		const p = await pres()
		const notes = p.slides.map((s) => s.notesText).filter((n) => n && n.trim())
		assert(notes.length >= 1, 'at least one slide has speaker notes')
		assert(
			notes.some((n) => /Speaker note/i.test(n)),
			'the authored note text is read back'
		)
	})

	test('lossless round-trip: save() succeeds', async () => {
		const bytes = await (await pres()).save()
		assert((bytes.length ?? bytes.byteLength) > 0, 'save() produces bytes')
	})
})
