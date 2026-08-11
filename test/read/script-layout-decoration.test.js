// A source layout's own shapes → the `objects` of its `defineSlideMaster` call.
//
// The standalone tier rebuilds a deck's design from scratch, and until now it rebuilt the
// layout gallery as a list of named blank rectangles: the bands, rules, wordmarks and
// triangles that make a deck recognisable as somebody's template were read, declared lost, and
// dropped. They are now transcribed — by the *same* mapper the slides go through, so a
// rectangle on a layout is decided by the code that decides one on a slide.
//
// Three oracles, because no single one reaches the whole surface:
//
//  1. **PowerPoint-authored layouts** (`mixed.pptx`, `read-stress.pptx`) for the arms real
//     templates are actually made of — nested groups, preset and custom geometry, scheme-token
//     fills, connectors drawn with the line tool, decorative text boxes. Nothing here is
//     authored by this repo, so nothing here can agree with the converter by construction.
//  2. **Write-API-authored layouts** for the two arms the corpus has neither of: an image and
//     a chart on a layout. `defineSlideMaster({ objects })` is the fixture, the same technique
//     `script-standalone.test.js` uses for a slide-owned background.
//  3. **A relocated table**, for the one shape kind that has no `SlideMasterObject` variant at
//     all. The write API cannot author it (that is the point) and no fixture carries one, so a
//     genuine write-API-authored `p:graphicFrame` is moved into a layout part in the zip —
//     which is what PowerPoint does when you paste a table onto a layout.
//
// Expectations come from `ts-pptx/read`'s own accessors wherever they can, never from the
// converter: both sides of a round trip run through one mapper, so a mapper that read a
// layout's geometry wrong would produce an output that is wrong the same way and compare
// clean. The round trip is here for the other half of the claim — that what the IR describes
// actually reaches the output deck — and one deliberate perturbation pins that it would notice
// if it did not.

import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { promisify } from 'node:util'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import TsPptx, { ChartType } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import {
	canonicalDeckIr,
	diffDeckIr,
	printScript,
	printStandaloneScript,
	readModelToIr,
	LAYOUT_NOTE_PREFIX,
} from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
import { FIXTURES } from './corpus.js'

const run = promisify(execFile)
const __dirname = path.dirname(fileURLToPath(import.meta.url))
const REPO = path.join(__dirname, '..', '..')
// Inside the repo: an emitted script imports this package by its published name, which Node
// resolves by the self-reference rule only from a path under the package root.
const SCRATCH = path.join(REPO, '.tmp')

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

/** A 1×1 PNG — the blip's bytes are irrelevant here, only that they make the round trip. */
const PNG_1x1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

async function load(name) {
	return Presentation.load(await readFile(path.join(FIXTURES, name)))
}

/** Every `SlideLayout` in a deck, in the gallery order the IR's `layoutIndex` addresses. */
function layoutsOf(presentation) {
	return presentation.masters().flatMap((master) => master.layouts)
}

/**
 * The same placeholder test the converter uses, spelled out here rather than imported: a test
 * that shared the predicate with the code under test could not catch it being wrong.
 */
function isPlaceholder(shape) {
	return shape.element_.getElementsByTagNameNS(P_NS, 'ph').length > 0
}

/** A layout's decorative shapes, flattened through groups the way the converter flattens them. */
function decorationOf(shapes, out = []) {
	for (const shape of shapes) {
		if (isPlaceholder(shape)) continue
		if (shape.shapes) decorationOf(shape.shapes, out)
		else out.push(shape)
	}
	return out
}

/** The `objectName` an emitted `SlideMasterObject` carries, whichever variant it is. */
function nameOfObject(object) {
	const [body] = Object.values(object)
	return body.options?.objectName ?? body.objectName ?? null
}

/** The single-key tag of an emitted `SlideMasterObject` — `shape`, `text`, `image`, `chart`. */
function tagOf(object) {
	return Object.keys(object)[0]
}

function objectsOfLayout(ir, index) {
	return ir.chrome.masters.find((master) => master.layoutIndex === index)?.props.objects ?? []
}

/** Author a deck through the write API and return its bytes. */
async function authored(build) {
	const pptx = new TsPptx()
	build(pptx)
	return /** @type {Buffer} */ (await pptx.write({ outputType: 'nodebuffer' }))
}

/**
 * Print the standalone script for these bytes, run it with no template in reach, and read back
 * what it produced.
 */
async function runStandalone(bytes) {
	await mkdir(SCRATCH, { recursive: true })
	const dir = await mkdtemp(path.join(SCRATCH, 'layout-deco-'))
	try {
		const ir = readModelToIr(await Presentation.load(bytes))
		const printed = printStandaloneScript(ir)
		await writeFile(path.join(dir, 'script.ts'), printed.code)
		if (printed.assets.size > 0) {
			await mkdir(path.join(dir, 'assets'), { recursive: true })
			for (const [name, data] of printed.assets) await writeFile(path.join(dir, 'assets', name), data)
		}
		await run(process.execPath, ['--no-warnings', path.join(dir, 'script.ts')])
		const output = await Presentation.load(await readFile(path.join(dir, 'output.pptx')))
		const outputIr = readModelToIr(output)
		return { ir, printed, output, outputIr, report: diffDeckIr(canonicalDeckIr(ir), canonicalDeckIr(outputIr), printed.notes) } // prettier-ignore
	} finally {
		await rm(dir, { recursive: true, force: true })
	}
}

describe('layout decoration — the IR, read against the deck rather than the converter', () => {
	test('a layout’s decoration reaches its objects, one per shape, placeholders excluded', async () => {
		// `mixed.pptx`'s title layout is the shape of the problem in miniature: seven decorative
		// rectangles buried two groups deep, and five placeholders alongside them that must *not*
		// come across — the write path seeds those onto every slide from the far side, so
		// re-declaring one here would put an empty ghost on every slide bound to the layout.
		const presentation = await load('mixed.pptx')
		const layout = layoutsOf(presentation)[0]
		assertEqual(layout.name, 'Diapositive de titre', 'the fixture layout this test is about')

		const decoration = decorationOf(layout.shapes)
		const placeholders = layout.shapes.filter(isPlaceholder)
		assert(decoration.length > 0 && placeholders.length > 0, 'the layout has both kinds to tell apart')

		const objects = objectsOfLayout(readModelToIr(presentation), 0)
		assertEqual(
			objects.map(nameOfObject).join(', '),
			decoration.map((shape) => shape.name).join(', '),
			'every decorative leaf, in tree order, and nothing else'
		)
		for (const placeholder of placeholders) {
			assert(!objects.some((object) => nameOfObject(object) === placeholder.name), `${placeholder.name} stayed out`)
		}
	})

	test('a flattened group keeps every child where the group put it, and says it flattened', async () => {
		// `SlideMasterObject` has no group variant. Flattening is only honest if the children land
		// unmoved, which they do because `absoluteFrame` composes the enclosing group's offset,
		// rotation, flips and child-space scaling — so the oracle is the read model's own frame,
		// not the converter's arithmetic.
		const presentation = await load('mixed.pptx')
		const layout = layoutsOf(presentation)[0]
		const groups = layout.shapes.filter((shape) => shape.shapes)
		assert(groups.length > 0, 'the layout carries at least one group')

		const ir = readModelToIr(presentation)
		const objects = objectsOfLayout(ir, 0)
		for (const object of objects) assertEqual(tagOf(object), 'shape', 'a flattened child is a plain object')

		for (const shape of decorationOf(layout.shapes)) {
			const object = objects.find((entry) => nameOfObject(entry) === shape.name)
			assert(object, `${shape.name} is emitted`)
			const frame = shape.absoluteFrame
			const options = Object.values(object)[0].options
			assertEqual(options.x, `${frame.left}emu`, `${shape.name} x`)
			assertEqual(options.y, `${frame.top}emu`, `${shape.name} y`)
			assertEqual(options.w, `${frame.width}emu`, `${shape.name} w`)
			assertEqual(options.h, `${frame.height}emu`, `${shape.name} h`)
		}

		// Every group, at every depth, is named by a note — including the nested ones, which a
		// walk that stopped at the top level would miss.
		const flattened = ir.fidelity.filter((note) => note.construct === `${LAYOUT_NOTE_PREFIX}group`)
		const nested = layout.shapes.flatMap((shape) => (shape.shapes ?? []).filter((child) => child.shapes))
		for (const group of [...groups, ...nested]) {
			assert(
				flattened.some((note) => note.shapeName === group.name),
				`the flattening of ${group.name} is declared`
			)
		}
	})

	test('a scheme-token fill stays a token, so the rebuilt layout still follows the theme', async () => {
		// The point of transcribing decoration at all is a layout that still looks like the
		// source's. Baking `accent2` to a hex would survive the round trip and quietly stop the
		// output deck from recolouring — the round trip compares two IRs, and a hex on both sides
		// agrees with itself.
		const presentation = await load('mixed.pptx')
		const layout = layoutsOf(presentation)[0]
		// The ten `p:clrMap` slots the write path can name. The other seven degrade to a literal
		// there anyway, so the converter bakes them deliberately and says so — the leg below.
		const writable = new Set(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'bg1', 'bg2', 'tx1', 'tx2']) // prettier-ignore
		const tokened = decorationOf(layout.shapes).filter((shape) => writable.has(shape.fillSchemeColor))
		assert(tokened.length > 0, 'the fixture layout paints from the theme')

		const ir = readModelToIr(presentation)
		const objects = objectsOfLayout(ir, 0)
		for (const shape of tokened) {
			const object = objects.find((entry) => nameOfObject(entry) === shape.name)
			const fill = Object.values(object)[0].options.fill
			// A gradient carries its tokens in its stops instead; either way none may be a hex.
			const token = fill.color ?? fill.gradient?.stops?.[0]?.color
			assertEqual(token, shape.fillSchemeColor, `${shape.name} keeps its scheme token`)
		}

		// And the ones that cannot stay tokens are named rather than quietly flattened. The
		// fixture has `folHlink` on a layout rectangle, which is outside the ten.
		const baked = decorationOf(layout.shapes).filter(
			(shape) => shape.fillSchemeColor !== null && !writable.has(shape.fillSchemeColor)
		)
		for (const shape of baked) {
			assert(
				ir.fidelity.some(
					(note) => note.construct === `${LAYOUT_NOTE_PREFIX}fill.schemeToken` && note.shapeName === shape.name
				),
				`${shape.name}'s ${shape.fillSchemeColor} is declared baked rather than silently flattened`
			)
		}
	})

	test('a connector becomes a line shape carrying its stroke', async () => {
		// PowerPoint's line tool authors a `p:cxnSp`, so a rule under a title is usually a
		// connector — 18 of the 45 shapes the corpus's layouts actually draw. `objects` has no
		// connector variant, so the alternative to this arm is dropping every one of them.
		const presentation = await load('read-stress.pptx')
		const layouts = layoutsOf(presentation)
		const index = layouts.findIndex((layout) => decorationOf(layout.shapes).some((s) => s.shapeType === 'connector'))
		assert(index >= 0, 'the fixture has a layout with connectors on it')

		const connectors = decorationOf(layouts[index].shapes).filter((shape) => shape.shapeType === 'connector')
		const objects = objectsOfLayout(readModelToIr(presentation), index)
		for (const connector of connectors) {
			const object = objects.find((entry) => nameOfObject(entry) === connector.name)
			assert(object, `${connector.name} is emitted rather than dropped`)
			assertEqual(object.shape.type, 'line', `${connector.name} is a line preset`)
			const frame = connector.absoluteFrame
			assertEqual(object.shape.options.x, `${frame.left}emu`, `${connector.name} x`)
			assertEqual(object.shape.options.w, `${frame.width}emu`, `${connector.name} w`)
			assertEqual(object.shape.options.flipH ?? false, connector.flipH, `${connector.name} flipH`)
			// The stroke is the whole reason it is visible; a line shape with no `line` would be a
			// hairline black default rather than the source's rule.
			assert(object.shape.options.line?.color, `${connector.name} keeps its stroke colour`)
		}
	})

	test('a decorative text box keeps its text and its text-box-ness', async () => {
		const presentation = await load('read-stress.pptx')
		const layouts = layoutsOf(presentation)
		const index = layouts.findIndex((layout) => layout.name === 'Quote with Caption')
		assert(index >= 0, 'the fixture has the layout with decorative quote marks on it')

		const objects = objectsOfLayout(readModelToIr(presentation), index)
		const boxes = decorationOf(layouts[index].shapes)
		assertEqual(objects.length, boxes.length, 'both quote marks are emitted')
		for (const box of boxes) {
			const object = objects.find((entry) => nameOfObject(entry) === box.name)
			assertEqual(tagOf(object), 'text', `${box.name} is a text object`)
			assertEqual(object.text.options.isTextBox, true, `${box.name} stays a text box rather than an auto shape`)
			assertEqual(
				object.text.text.map((runIr) => runIr.text).join(''),
				box.textFrame.text,
				`${box.name} keeps its glyphs`
			)
		}
	})

	test('across the corpus, every decorative layout shape is emitted or noted — never silently gone', async () => {
		// The invariant the fidelity contract rests on, checked over every layout of every deck
		// rather than the two this file names. A shape that reached neither the `objects` array nor
		// a note is a silent loss, which is the failure mode the notes exist to make impossible.
		for (const name of ['mixed.pptx', 'read-stress.pptx', 'theme-colors.pptx', 'gradient-fill.pptx']) {
			const presentation = await load(name)
			const ir = readModelToIr(presentation)
			const noted = new Set(
				ir.fidelity.filter((note) => note.construct.startsWith(LAYOUT_NOTE_PREFIX)).map((note) => note.shapeName)
			)
			layoutsOf(presentation).forEach((layout, index) => {
				const emitted = new Set(objectsOfLayout(ir, index).map(nameOfObject))
				for (const shape of decorationOf(layout.shapes)) {
					assert(emitted.has(shape.name) || noted.has(shape.name), `${name} / ${layout.name}: ${shape.name} vanished`)
				}
			})
		}
	})
})

describe('layout decoration — the emitted script rebuilds it, with no template in reach', () => {
	test('the output deck’s layouts carry the source’s decoration, and only on the layouts', async () => {
		const bytes = await readFile(path.join(FIXTURES, 'read-stress.pptx'))
		const source = await Presentation.load(bytes)
		const { ir, output, report } = await runStandalone(bytes)
		assertEqual(report.undeclared.length, 0, 'no undeclared round-trip difference')

		// Matched through the IR's own title, not by name: this deck has two masters and two
		// layouts called "Title Slide", only the second of which is decorated, so a name lookup
		// would compare the decorated source layout against the bare output one and pass or fail
		// for the wrong reason. The IR's title is the deduplicated key the output layout carries.
		let decorated = 0
		layoutsOf(source).forEach((layout, index) => {
			const shapes = decorationOf(layout.shapes)
			if (shapes.length === 0) return
			decorated++
			const title = ir.chrome.masters.find((master) => master.layoutIndex === index)?.props.title
			const rebuilt = layoutsOf(output).find((candidate) => candidate.name === title)
			assert(rebuilt, `the output has a layout titled ${JSON.stringify(title)}`)
			const names = new Set(decorationOf(rebuilt.shapes).map((shape) => shape.name))
			for (const shape of shapes) {
				assert(names.has(shape.name), `${title}: ${shape.name} reached the output layout`)
			}
		})
		assert(decorated > 0, 'the fixture has decorated layouts to rebuild')

		// And it stayed there. Decoration painted onto every slide instead would render the same
		// and be a different deck — the shapes would be uneditable from the layout and duplicated
		// once per slide.
		source.slides.forEach((slide, index) => {
			const before = slide.shapes.length
			const after = output.slides[index].shapes.length
			assert(after <= before, `slide ${index + 1} gained no shapes from its layout (${before} → ${after})`)
		})
	})

	test('a layout that lost its decoration is an undeclared difference, not a clean run', async () => {
		// The oracle's half of the claim above. Without it, the round trip passes just as happily
		// against a converter that emits no `objects` at all, because nothing would then differ
		// between two sides that are both empty.
		const { ir, outputIr, printed } = await runStandalone(await readFile(path.join(FIXTURES, 'read-stress.pptx')))
		const perturbed = canonicalDeckIr(outputIr)
		const masters = /** @type {Record<string, any>[]} */ (perturbed.chrome.masters)
		const victim = masters.find((master) => Array.isArray(master.objects) && master.objects.length > 0)
		assert(victim, 'the output IR has a decorated layout to strip')
		delete victim.objects

		const dirty = diffDeckIr(canonicalDeckIr(ir), perturbed, printed.notes)
		assert(
			dirty.undeclared.some((difference) => difference.field === 'objects'),
			'a layout stripped of its decoration must be reported'
		)
	})

	test('the template-anchored tier reports none of it, because it rebuilds no layout', async () => {
		// A caveat that does not describe the output in front of you teaches the reader to skim
		// the ones that do. Tier B reuses the source deck as its template, so every layout — and
		// every loss in re-authoring one — is somebody else's problem.
		for (const name of ['mixed.pptx', 'read-stress.pptx']) {
			const ir = readModelToIr(await load(name))
			const standalone = printStandaloneScript(ir).notes.filter((n) => n.construct.startsWith(LAYOUT_NOTE_PREFIX))
			const anchored = printScript(ir).notes.filter((n) => n.construct.startsWith(LAYOUT_NOTE_PREFIX))
			assert(standalone.length > 0, `${name}: the standalone tier declares its layout losses`)
			assertEqual(anchored.length, 0, `${name}: the template-anchored tier declares none`)
		}
	})
})

describe('layout decoration — kinds the fixture corpus does not contain', () => {
	test('an image on a layout carries its bytes across', async () => {
		// No fixture has one, and a logo is the single most common thing a real template puts on a
		// layout. The write API is the fixture: `defineSlideMaster({ objects })` authors the source
		// deck, which is then read like any other.
		const bytes = await authored((pptx) => {
			pptx.defineSlideMaster({
				title: 'Branded',
				objects: [{ image: { data: PNG_1x1, x: 0.5, y: 0.25, w: 1, h: 1, objectName: 'Wordmark' } }],
			})
			pptx.addSlide({ masterTitle: 'Branded' }).addText('body', { x: 1, y: 3, w: 4, h: 1 })
		})

		const ir = readModelToIr(await Presentation.load(bytes))
		const master = ir.chrome.masters.find((entry) => entry.props.title === 'Branded')
		const [object] = /** @type {any[]} */ (master.props.objects)
		assertEqual(tagOf(object), 'image', 'the layout image is emitted as an image object')
		assertEqual(nameOfObject(object), 'Wordmark', 'and keeps its name')
		assert(object.image.data?.$asset, 'with its bytes carried as an asset rather than a path')

		const { output, report } = await runStandalone(bytes)
		assertEqual(report.undeclared.length, 0, 'no undeclared round-trip difference')
		const rebuilt = layoutsOf(output).find((layout) => layout.name === 'Branded')
		const picture = decorationOf(rebuilt.shapes).find((shape) => shape.name === 'Wordmark')
		assert(picture?.imagePartName, 'the output layout holds a picture with an embedded image part')
	})

	test('a chart on a layout carries its series across', async () => {
		const bytes = await authored((pptx) => {
			pptx.defineSlideMaster({
				title: 'Dashboard',
				objects: [
					{
						chart: {
							type: ChartType.bar,
							data: [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [3, 5] }],
							options: { x: 1, y: 1, w: 4, h: 3, objectName: 'Trend' },
						},
					},
				],
			})
			pptx.addSlide({ masterTitle: 'Dashboard' }).addText('body', { x: 1, y: 5, w: 4, h: 1 })
		})

		const ir = readModelToIr(await Presentation.load(bytes))
		const master = ir.chrome.masters.find((entry) => entry.props.title === 'Dashboard')
		const [object] = /** @type {any[]} */ (master.props.objects)
		assertEqual(tagOf(object), 'chart', 'the layout chart is emitted as a chart object')
		assertEqual(object.chart.type, 'bar', 'with the type lifted back out of the addChart options')
		assertEqual(object.chart.data[0].values.join(','), '3,5', 'and its cached series values')

		const { output, report } = await runStandalone(bytes)
		assertEqual(report.undeclared.length, 0, 'no undeclared round-trip difference')
		const rebuilt = layoutsOf(output).find((layout) => layout.name === 'Dashboard')
		assert(
			decorationOf(rebuilt.shapes).some((shape) => shape.chart),
			'the output layout holds a chart'
		)
	})

	test('a table on a layout is dropped, and says so', async () => {
		// The one shape kind with no `SlideMasterObject` variant. It cannot be authored onto a
		// layout through the write API — that is the gap — and no fixture carries one, so a real
		// write-API `p:graphicFrame` is moved from the slide into the layout part, which is what
		// PowerPoint writes when you paste a table onto a layout. A table needs no relationship of
		// its own, so relocating the element is the whole of the move.
		const bytes = await authored((pptx) => {
			pptx.defineSlideMaster({ title: 'Tabular' })
			const slide = pptx.addSlide({ masterTitle: 'Tabular' })
			slide.addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 4, h: 1, objectName: 'Grid' })
		})

		const zip = await JSZip.loadAsync(bytes)
		const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
		const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(slideXml)
		assert(frame, 'the authored slide holds the table as a graphic frame')
		const layoutPath = 'ppt/slideLayouts/slideLayout2.xml'
		const layoutXml = await zip.file(layoutPath).async('string')
		assert(layoutXml.includes('Tabular'), 'slideLayout2 is the layout defineSlideMaster created')
		zip.file(layoutPath, layoutXml.replace('</p:spTree>', `${frame[0]}</p:spTree>`))
		zip.file('ppt/slides/slide1.xml', slideXml.replace(frame[0], ''))

		const relocated = await zip.generateAsync({ type: 'nodebuffer' })
		const presentation = await Presentation.load(relocated)
		const layout = layoutsOf(presentation).find((entry) => entry.name === 'Tabular')
		assert(
			decorationOf(layout.shapes).some((shape) => shape.table),
			'the layout really holds the table now'
		)

		const ir = readModelToIr(presentation)
		const master = ir.chrome.masters.find((entry) => entry.props.title === 'Tabular')
		assertEqual(master.props.objects, undefined, 'the table produced no object')
		assert(
			ir.fidelity.some((note) => note.construct === `${LAYOUT_NOTE_PREFIX}decoration` && note.shapeName === 'Grid'),
			'and the drop is declared against the shape by name'
		)
	})
})
