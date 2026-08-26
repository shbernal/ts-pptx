#!/usr/bin/env node
/**
 * Emit *edited* decks from the read fixtures so each editing capability can be
 * opened in PowerPoint and confirmed to render without a repair prompt. Unlike
 * `read-emit-roundtrip.mjs` (which saves an unmodified load → save), every deck
 * here exercises a mutation that reserializes parts, which is what PowerPoint's
 * stricter desktop validation actually reacts to.
 *
 * Each output is named for the edit it performs (added-textbox, added-picture,
 * deleted-shape, cloned-slide, edited-table-cells, imported-*-slide). Output goes to
 * .tmp/read-edits/ (gitignored) by default; override with the first CLI arg or
 * TSPPTX_READ_EDITS_DIR. Assumes a current build — the
 * test:read:emit:edits script ensures `dist/` is current first.
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT, parseCliOrExit } from './script-utils.mjs'

const { positionals } = parseCliOrExit(process.argv.slice(2), {
	usage: `Emit *edited* decks from the read fixtures, for the manual PowerPoint check.

  pnpm run test:read:emit:edits
  pnpm run test:read:emit:edits -- <out-dir>

This is a GENERATOR, not a gate: every deck exercises one editing capability and is
written for a human to open in desktop PowerPoint. Nothing here asserts on the output.

Arguments:
  <out-dir>   where to write (default .tmp/read-edits, or $TSPPTX_READ_EDITS_DIR)

Options:
  -h, --help  show this message`,
	allowPositionals: true,
	options: {},
})

const fixturesDir = path.join(ROOT, 'test', 'read', 'fixtures')
const outDir = positionals[0] || process.env.TSPPTX_READ_EDITS_DIR || path.join(ROOT, '.tmp', 'read-edits')

const readEntry = path.join(ROOT, 'dist', 'read.js')
try {
	await fs.access(readEntry)
} catch {
	console.error(
		`Missing ${path.relative(ROOT, readEntry)}. Run \`pnpm run build\` first (or use \`pnpm run test:read:emit:edits\`).`
	)
	process.exit(1)
}
// `readEntry` is always `dist/read.js`; the import is dynamic only so the check above can
// print a build-first message instead of an unresolved-specifier stack.
const { Presentation, isGraphicFrame } = /** @type {typeof import('../dist/read.js')} */ (
	await import(pathToFileURL(readEntry).href)
)

/**
 * Open a fixture by base name (without extension) as a Presentation.
 * @param {string} name
 * @returns {Promise<import('../dist/read.js').Presentation>}
 */
async function open(name) {
	return Presentation.load(await fs.readFile(path.join(fixturesDir, `${name}.pptx`)))
}

/**
 * The deck's first slide. Every fixture opened here has one, so an empty deck is a broken
 * fixture rather than a case to edit around.
 * @param {import('../dist/read.js').Presentation} presentation
 * @returns {import('../dist/read.js').Slide}
 */
function firstSlide(presentation) {
	const slide = presentation.slides[0]
	if (!slide) throw new Error('fixture has no slides')
	return slide
}

/** A real raster image (bytes + type) borrowed from the image fixture's media. */
async function sampleImage() {
	const presentation = await open('image')
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType !== 'picture' || !shape.imagePartName) continue
			if (!/\.(png|jpe?g)$/i.test(shape.imagePartName)) continue
			const part = presentation.opc.part(shape.imagePartName)
			if (part) return { bytes: part.bytes, contentType: part.contentType, partName: shape.imagePartName }
		}
	}
	throw new Error('No raster image found in image.pptx to sample')
}

const inch = 914400

/**
 * Each case names the edit it exercises and returns the saved bytes. The
 * comment on each is the thing to verify visually in PowerPoint.
 */
const cases = [
	{
		out: 'empty.added-textbox.pptx',
		// Verify: a text box reading "Added via addTextBox" appears on the slide.
		async build() {
			const presentation = await open('empty')
			firstSlide(presentation).addTextBox({
				text: 'Added via addTextBox',
				left: 1 * inch,
				top: 1 * inch,
				width: 5 * inch,
				height: 1 * inch,
				name: 'EmittedTextBox',
			})
			return presentation.save()
		},
	},
	{
		out: 'empty.added-picture.pptx',
		// Verify: the borrowed raster image renders (not a missing-image placeholder).
		async build() {
			const presentation = await open('empty')
			const image = await sampleImage()
			firstSlide(presentation).addPicture(image.bytes, {
				left: 1 * inch,
				top: 1 * inch,
				width: 3 * inch,
				height: 3 * inch,
				name: 'EmittedPicture',
				contentType: image.contentType,
			})
			return presentation.save()
		},
	},
	{
		out: 'textbox.deleted-shape.pptx',
		// Verify: the "replaceText" shape is gone; the rest of the slide is intact.
		async build() {
			const presentation = await open('textbox')
			firstSlide(presentation)
				.shapes.find((shape) => shape.name === 'replaceText')
				?.delete()
			return presentation.save()
		},
	},
	{
		out: 'textbox.cloned-slide.pptx',
		// Verify: a duplicate of slide 1 appears as the last slide, reading "CLONED COPY".
		async build() {
			const presentation = await open('textbox')
			const clone = presentation.cloneSlide(0)
			const run = clone.shapes.find((shape) => shape.hasTextFrame)?.textFrame?.paragraphs[0]?.runs[0]
			if (run) run.text = 'CLONED COPY'
			return presentation.save()
		},
	},
	{
		out: 'table.edited-cells.pptx',
		// Verify: the first table's top-left cells read "Edited A1" / "Edited B1".
		async build() {
			const presentation = await open('table')
			const frame = presentation.slides
				.flatMap((slide) => slide.shapes)
				.find((shape) => isGraphicFrame(shape) && shape.hasTable)
			const table = frame && isGraphicFrame(frame) ? frame.table : null
			const a1 = table?.cell(0, 0)
			const b1 = table?.cell(0, 1)
			if (a1) a1.text = 'Edited A1'
			if (b1) b1.text = 'Edited B1'
			return presentation.save()
		},
	},
	{
		out: 'smartart-families.retexted-smartart.pptx',
		// Verify: every node on all four slides reads "edit-N", and each diagram still draws
		// whole — org chart with its assistant off the branch, process arrows and their labels,
		// cycle, picture list. No repair prompt.
		//
		// The point of the case is the *second* copy of the text: `DiagramPoint.text` writes
		// `ppt/diagrams/data{N}.xml` and mirrors into the `ppt/diagrams/drawing{N}.xml` cache.
		// PowerPoint regenerates that cache on open, so it is the one renderer that cannot tell
		// you whether the mirror worked — open the file in LibreOffice, or export a slide to
		// PNG before PowerPoint has re-saved it, to see what everything else draws.
		//
		// The replacements are kept about as long as the strings they replace, deliberately.
		// Geometry is not recomputed, so a much longer string overflows the cached box until
		// PowerPoint re-lays the diagram out, and that overflow would be the only thing anyone
		// looking at this deck could see.
		async build() {
			const presentation = await open('smartart-families')
			for (const slide of presentation.slides) {
				const frame = slide.shapes.find((shape) => isGraphicFrame(shape) && shape.hasDiagram)
				const diagram = frame && isGraphicFrame(frame) ? frame.diagram : null
				if (!diagram) continue
				let n = 0
				/** @param {import('../dist/read.js').DiagramNode[]} nodes */
				const walk = (nodes) => {
					for (const node of nodes) {
						node.point.text = `edit-${++n}`
						walk(node.children)
					}
				}
				walk(diagram.nodes)
			}
			return presentation.save()
		},
	},
	{
		out: 'empty.imported-image-slide.pptx',
		// Verify: a slide carrying the image fixture's picture (with its own
		// layout/master/theme) is appended to the otherwise-blank deck.
		async build() {
			const target = await open('empty')
			const source = await open('image')
			target.importSlide(source, 0)
			return target.save()
		},
	},
	{
		out: 'empty.imported-table-slide.pptx',
		// Verify: a slide carrying a table from the table fixture is appended.
		async build() {
			const target = await open('empty')
			const source = await open('table')
			const tableSlide = source.slides.findIndex((slide) =>
				slide.shapes.some((shape) => isGraphicFrame(shape) && shape.hasTable)
			)
			target.importSlide(source, tableSlide === -1 ? 0 : tableSlide)
			return target.save()
		},
	},
]

await fs.mkdir(outDir, { recursive: true })
for (const testCase of cases) {
	const bytes = await testCase.build()
	const outPath = path.join(outDir, testCase.out)
	await fs.writeFile(outPath, bytes)
	console.log(`${testCase.out}: ${bytes.length} bytes  ${path.relative(ROOT, outPath)}`)
}

console.log(
	`\nOpen the files in ${path.relative(ROOT, outDir)}/ in PowerPoint (desktop especially) and confirm no repair prompt.`
)
console.log('Record the result in test/read/fixtures/README.md (the "edited output" checklist).')
