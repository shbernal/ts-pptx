#!/usr/bin/env node
/**
 * Emit *edited* decks from the read fixtures so each editing capability can be
 * opened in PowerPoint and confirmed to render without a repair prompt. Unlike
 * `read-emit-roundtrip.mjs` (which saves an unmodified load → save), every deck
 * here exercises a mutation that reserializes parts, which is what PowerPoint's
 * stricter desktop validation actually reacts to.
 *
 * Each output is named for the edit it performs (added-textbox, added-picture,
 * deleted-shape, cloned-slide, edited-table-cells). Output goes to
 * .tmp/read-edits/ (gitignored) by default; override with the first CLI arg or
 * PPTXGENJS_READ_EDITS_DIR. Assumes a current build — the
 * test:read:emit:edits script runs `pnpm run build` first.
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT } from './script-utils.mjs'

const fixturesDir = path.join(ROOT, 'test', 'read', 'fixtures')
const outDir = process.argv[2] || process.env.PPTXGENJS_READ_EDITS_DIR || path.join(ROOT, '.tmp', 'read-edits')

const readEntry = path.join(ROOT, 'dist', 'read.js')
try {
	await fs.access(readEntry)
} catch {
	console.error(
		`Missing ${path.relative(ROOT, readEntry)}. Run \`pnpm run build\` first (or use \`pnpm run test:read:emit:edits\`).`
	)
	process.exit(1)
}
const { Presentation } = await import(pathToFileURL(readEntry).href)

/** Open a fixture by base name (without extension) as a Presentation. */
async function open(name) {
	return Presentation.load(await fs.readFile(path.join(fixturesDir, `${name}.pptx`)))
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
			presentation.slides[0].addTextBox({
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
			presentation.slides[0].addPicture(image.bytes, {
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
			presentation.slides[0].shapes.find((shape) => shape.name === 'replaceText')?.delete()
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
				.find((shape) => shape.shapeType === 'graphicFrame' && shape.table)
			const table = frame?.table
			if (table) {
				table.cell(0, 0).text = 'Edited A1'
				table.cell(0, 1).text = 'Edited B1'
			}
			return presentation.save()
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
