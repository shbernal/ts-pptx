import {
	inspectPptx,
	boxAnchor,
	listPptxParts,
	loadPptxPackage,
	overlapArea,
	readPptxBinaryPart,
	readPptxTextPart,
} from '../../dist/inspect.js'
import JSZip from 'jszip'
import { defineRegressionSuite, build, assert, assertEqual } from '../helpers.js'
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const PNG_1X1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), '..', 'read', 'fixtures')

/** Inches, generous next to the ~0.4in group-scale error this guards, tight next to PowerPoint's own EMU rounding (2 EMU ≈ 2.2e-6in). */
const INCH_TOLERANCE = 1e-5

function assertWithin(actual, expected, tolerance, msg) {
	assert(
		Math.abs(actual - expected) <= tolerance,
		`${msg}: expected ${expected} ± ${tolerance}, got ${actual} (off by ${Math.abs(actual - expected)})`
	)
}

/**
 * A minimal package holding `spTreeXml` — enough for inspect, which reads only
 * `presentation.xml` and the slide parts. Built with jszip so the fflate reader
 * under test is not also the writer (see helpers.js).
 */
async function packageWithSpTree(spTreeXml) {
	const zip = new JSZip()
	zip.file(
		'ppt/presentation.xml',
		`<?xml version="1.0"?><p:presentation xmlns:p="${P_NS}"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`
	)
	zip.file(
		'ppt/slides/slide1.xml',
		`<?xml version="1.0"?><p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:cSld><p:spTree>${spTreeXml}</p:spTree></p:cSld></p:sld>`
	)
	return zip.generateAsync({ type: 'uint8array' })
}

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** `<p:sp>` with a name and an explicit transform, for hand-authored shape trees. */
function spXml(id, name, { x = 0, y = 0, cx = 100, cy = 100 } = {}) {
	return `<p:sp>
		<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
		<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr>
	</p:sp>`
}

/** Collect `console.warn` output while `fn` runs. */
async function captureWarnings(fn) {
	const warnings = []
	const original = console.warn
	console.warn = (...args) => warnings.push(args.join(' '))
	try {
		return { result: await fn(), warnings }
	} finally {
		console.warn = original
	}
}

defineRegressionSuite('PPTX inspection primitives', [
	{
		name: 'inspectPptx extracts slide size, named objects, geometry, and text style',
		fn: async () => {
			const { buf } = await build((p) => {
				p.layout = 'LAYOUT_WIDE'
				const slide = p.addSlide()
				slide.addText('Inspect me', {
					x: 1,
					y: 1.25,
					w: 2.5,
					h: 0.5,
					objectName: 'inspect:text',
					fontSize: 18,
					color: '336699',
				})
				slide.addShape(p.ShapeType.rect, {
					x: 4,
					y: 1,
					w: 1.5,
					h: 1,
					objectName: 'inspect:shape',
					fill: { color: 'FF0000' },
				})
				slide.addImage({
					data: `image/png;base64,${PNG_1X1}`,
					x: 6,
					y: 1,
					w: 1,
					h: 1,
					objectName: 'inspect:image',
				})
			})

			const inspection = await inspectPptx(buf)
			assertEqual(inspection.slideSize.widthIn, 13.333, 'slide width')
			assertEqual(inspection.slideSize.heightIn, 7.5, 'slide height')
			assertEqual(inspection.slides.length, 1, 'slide count')
			assertEqual(inspection.slides[0].wordCount, 2, 'word count')

			const elements = new Map(inspection.slides[0].elements.map((element) => [element.name, element]))
			const text = elements.get('inspect:text')
			const shape = elements.get('inspect:shape')
			const image = elements.get('inspect:image')

			assert(text, 'expected named text element')
			assert(shape, 'expected named shape element')
			assert(image, 'expected named image element')
			assertEqual(text.kind, 'text', 'text kind')
			assertEqual(text.text, 'Inspect me', 'text content')
			assertEqual(text.fontSizes[0], 18, 'font size')
			assertEqual(text.colors[0], '336699', 'font color')
			assertEqual(shape.kind, 'shape', 'shape kind')
			assertEqual(shape.fill, 'FF0000', 'shape fill')
			assertEqual(shape.shapeType, 'rect', 'shape type')
			assertEqual(image.kind, 'image', 'image kind')
			assert(Math.abs(text.box.x - 1) < 0.001, 'expected x position in inches')
		},
	},
	{
		name: 'inspect exposes per-text-frame autofit mode and body insets',
		fn: async () => {
			const { buf } = await build((p) => {
				const slide = p.addSlide()
				// Fixed-height box (no autofit) with default body insets — a genuine overflow candidate.
				slide.addText('Fixed', { x: 1, y: 1, w: 2, h: 0.5, objectName: 'fit:none' })
				// Shrink-to-fit (normAutofit) — text downscales rather than overflowing.
				slide.addText('Shrink', { x: 1, y: 2, w: 2, h: 0.5, fit: 'shrink', objectName: 'fit:shrink' })
				// Resize-shape-to-fit (spAutoFit) — authored height is an output, cannot overflow.
				slide.addText('Resize', { x: 1, y: 3, w: 2, h: 0.5, fit: 'resize', objectName: 'fit:resize' })
				// Custom zero insets via the `margin` option.
				slide.addText('Tight', { x: 1, y: 4, w: 2, h: 0.5, margin: 0, objectName: 'fit:margin0' })
				slide.addImage({ data: `image/png;base64,${PNG_1X1}`, x: 6, y: 1, w: 1, h: 1, objectName: 'fit:image' })
			})

			const inspection = await inspectPptx(buf)
			const elements = new Map(inspection.slides[0].elements.map((element) => [element.name, element]))

			const fixed = elements.get('fit:none')
			assertEqual(fixed.autofit, 'none', 'no-autofit box reports none')
			assert(Math.abs(fixed.bodyInsets.left - 0.1) < 1e-6, 'default left inset is 0.1in')
			assert(Math.abs(fixed.bodyInsets.right - 0.1) < 1e-6, 'default right inset is 0.1in')
			assert(Math.abs(fixed.bodyInsets.top - 0.05) < 1e-6, 'default top inset is 0.05in')
			assert(Math.abs(fixed.bodyInsets.bottom - 0.05) < 1e-6, 'default bottom inset is 0.05in')

			assertEqual(elements.get('fit:shrink').autofit, 'normAutofit', 'shrink box reports normAutofit')
			assertEqual(elements.get('fit:resize').autofit, 'spAutoFit', 'resize box reports spAutoFit')

			const tight = elements.get('fit:margin0')
			assertEqual(tight.autofit, 'none', 'margin-only box still reports none')
			assertEqual(tight.bodyInsets.left, 0, 'zero margin is preserved, not defaulted')
			assertEqual(tight.bodyInsets.bottom, 0, 'zero margin is preserved, not defaulted')

			const image = elements.get('fit:image')
			assertEqual(image.autofit, null, 'image without a text frame has no autofit')
			assertEqual(image.bodyInsets, null, 'image without a text frame has no body insets')
		},
	},
	{
		name: 'inspect exposes per-run font props, paragraph structure, and baked normAutofit fontScale',
		fn: async () => {
			const { buf } = await build((p) => {
				const slide = p.addSlide()
				// Two paragraphs (breakLine ends the first), each a run with distinct props.
				slide.addText(
					[
						{ text: 'Bold', options: { bold: true, fontFace: 'Arial', breakLine: true } },
						{ text: 'Ital', options: { italic: true, charSpacing: 2 } },
					],
					{ x: 1, y: 1, w: 3, h: 1, fontSize: 18, objectName: 'runs' }
				)
				// Object-form shrink bakes an explicit fontScale we can read back.
				slide.addText('Scaled', {
					x: 1,
					y: 3,
					w: 2,
					h: 0.5,
					fit: { type: 'shrink', fontScale: 62.5 },
					objectName: 'scaled',
				})
				// Bare shrink → <a:normAutofit/> with no fontScale.
				slide.addText('Bare', { x: 1, y: 4, w: 2, h: 0.5, fit: 'shrink', objectName: 'bare' })
			})

			const inspection = await inspectPptx(buf)
			const elements = new Map(inspection.slides[0].elements.map((element) => [element.name, element]))

			const runs = elements.get('runs')
			assert(runs, 'expected named multi-run element')
			assertEqual(runs.paragraphs.length, 2, 'two source paragraphs preserved')
			const first = runs.paragraphs[0].runs[0]
			const second = runs.paragraphs[1].runs[0]
			assertEqual(first.text, 'Bold', 'first paragraph run text')
			assertEqual(first.bold, true, 'bold flag read from a:rPr@b')
			assertEqual(first.italic, false, 'first run is not italic')
			assertEqual(first.fontFace, 'Arial', 'fontFace read from a:latin@typeface')
			assertEqual(second.italic, true, 'italic flag read from a:rPr@i')
			assertEqual(second.charSpacingPt, 2, 'charSpacing read from a:rPr@spc (hundredths→pt)')
			// Flat textRuns still carries the same enriched props.
			assertEqual(runs.textRuns[0].fontFace, 'Arial', 'flat textRuns also carry fontFace')

			assertEqual(elements.get('scaled').autofit, 'normAutofit', 'object shrink reports normAutofit')
			assertEqual(elements.get('scaled').autofitFontScale, 62.5, 'baked fontScale read back as a percent')
			assertEqual(elements.get('bare').autofit, 'normAutofit', 'bare shrink still reports normAutofit')
			assertEqual(elements.get('bare').autofitFontScale, null, 'bare normAutofit bakes no scale → null')
		},
	},
	{
		// D1: inspect used to report each shape's raw a:xfrm, which for a group child is
		// authored in the group's child space (a:chOff/a:chExt) — not placeable on the
		// slide. That was silently right only for packages this library authored (its
		// writer hardcodes an identity child space) and wrong for any deck PowerPoint has
		// touched: resizing a group alone makes chExt non-identity.
		//
		// Ground truth is PowerPoint's own: slide 1 of the fixture holds the groups and
		// slide 2 the same shapes after a PowerPoint Ungroup, so each grouped child has a
		// twin whose raw geometry IS its absolute geometry. No oracle of ours is involved.
		name: 'inspect composes enclosing group transforms into slide-absolute boxes',
		fn: async () => {
			const inspection = await inspectPptx(join(FIXTURES, 'group-transform.pptx'))
			const [grouped, ungrouped] = inspection.slides
			assertEqual(
				ungrouped.elements.filter((el) => el.kind === 'group').length,
				0,
				'slide 2 is PowerPoint-ungrouped ground truth'
			)

			const twins = new Map(ungrouped.elements.map((el) => [el.name, el]))
			const children = grouped.elements.filter((el) => el.name.includes(' child '))
			assert(children.length >= 21, `expected the fixture's grouped children, got ${children.length}`)

			let scaled = 0
			for (const child of children) {
				const expected = twins.get(child.name.replace(/^(.+?) child /, '$1-ungrouped child '))
				assert(expected, `expected a PowerPoint-ungrouped twin for "${child.name}"`)
				assertWithin(child.box.x, expected.box.x, INCH_TOLERANCE, `${child.name} absolute x`)
				assertWithin(child.box.y, expected.box.y, INCH_TOLERANCE, `${child.name} absolute y`)
				assertWithin(child.box.w, expected.box.w, INCH_TOLERANCE, `${child.name} absolute w`)
				assertWithin(child.box.h, expected.box.h, INCH_TOLERANCE, `${child.name} absolute h`)
				assertWithin(child.rotation, expected.rotation, 1e-6, `${child.name} effective rotation`)
				assertEqual(child.flipH, expected.flipH, `${child.name} effective flipH`)
				assertEqual(child.flipV, expected.flipV, `${child.name} effective flipV`)
				if (child.name.startsWith('scale-rot child ')) scaled++
			}
			// The scale cases are the ones the old raw-a:xfrm read got wrong by ~32% on
			// width; without them this test would pass against the identity-child-space bug.
			assert(scaled > 0, 'expected the scale+rotation group children to be covered')
		},
	},
	{
		// D2: 'p:grpSp' was absent from the harvested key list, so a group reached the
		// output only as a side effect of the generic walker recursing into every object
		// value — its identity and fill were dropped, and the flat element list gave no
		// indication which elements were grouped.
		name: 'inspect reports the group itself and links it to its children',
		fn: async () => {
			const inspection = await inspectPptx(join(FIXTURES, 'group-transform.pptx'))
			const [grouped] = inspection.slides

			const groups = grouped.elements.filter((el) => el.kind === 'group')
			assert(groups.length > 0, 'expected the fixture groups to be reported as elements')

			const byZ = new Map(grouped.elements.map((el) => [el.zIndex, el]))
			const nested = groups.find((group) => group.parentZIndex !== null)
			assert(nested, 'expected the fixture nested group to name its enclosing group')
			assertEqual(byZ.get(nested.parentZIndex).kind, 'group', 'a parentZIndex resolves to a group element')

			for (const group of groups) {
				assert(group.name, 'a group reports its cNvPr name')
				assert(group.childZIndices.length > 0, `group "${group.name}" reports its children`)
				for (const zIndex of group.childZIndices) {
					const child = byZ.get(zIndex)
					assert(child, `child z=${zIndex} of "${group.name}" resolves to an element`)
					assertEqual(child.parentZIndex, group.zIndex, `child "${child.name}" points back at its group`)
				}
			}

			// Top-level shapes must not claim a parent, and only groups have children.
			for (const el of grouped.elements) {
				if (el.kind !== 'group') assertEqual(el.childZIndices.length, 0, `"${el.name}" is not a group`)
			}
			assert(
				grouped.elements.some((el) => el.parentZIndex === null),
				'expected at least one slide-level element'
			)
		},
	},
	{
		// D3: zIndex came from the harvest order ("every p:sp of a node, then every p:pic,
		// then every p:cxnSp, then recurse"), not document order — so an image authored
		// between two text boxes sorted after both. Wrong for any mixed-type slide, with
		// or without groups.
		name: 'inspect zIndex follows document order across element kinds',
		fn: async () => {
			const { buf } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('First', { x: 1, y: 1, w: 1, h: 0.4, objectName: 'z:first' })
				slide.addImage({ data: `image/png;base64,${PNG_1X1}`, x: 2, y: 1, w: 1, h: 1, objectName: 'z:middle' })
				slide.addText('Last', { x: 3, y: 1, w: 1, h: 0.4, objectName: 'z:last' })
			})

			const [slide] = (await inspectPptx(buf)).slides
			const order = slide.elements.sort((a, b) => a.zIndex - b.zIndex).map((el) => el.name)
			assertEqual(order.join(' < '), 'z:first < z:middle < z:last', 'z order matches authored document order')
		},
	},
	{
		// A group whose a:chExt is zero has no child-space ratio, so its children have no
		// resolvable slide position. Reporting the raw child-space box would be a
		// confidently wrong number; per AGENTS.md the element is dropped with a warning
		// instead. (The read API's absoluteFrame returns null on the same input.)
		name: 'inspect omits children of a degenerate group and says so',
		fn: async () => {
			const buf = await packageWithSpTree(`
				${spXml(2, 'healthy sibling', { x: 10, y: 20, cx: 30, cy: 40 })}
				<p:grpSp>
					<p:nvGrpSpPr><p:cNvPr id="3" name="degenerate group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
					<p:grpSpPr><a:xfrm>
						<a:off x="0" y="0"/><a:ext cx="6858000" cy="0"/>
						<a:chOff x="0" y="0"/><a:chExt cx="6858000" cy="0"/>
					</a:xfrm></p:grpSpPr>
					${spXml(4, 'unresolvable child')}
				</p:grpSp>`)

			const { result, warnings } = await captureWarnings(() => inspectPptx(buf))
			const names = result.slides[0].elements.map((el) => el.name)

			assert(!names.includes('unresolvable child'), 'a child with no resolvable position is not reported')
			assert(
				warnings.some((line) => line.includes('unresolvable child') && line.includes('degenerate')),
				`expected a warning naming the dropped child; got ${JSON.stringify(warnings)}`
			)
			// The group itself sits at slide level, so its own box still resolves — only
			// what it maps is unresolvable. And one bad group must not poison the slide.
			assert(names.includes('degenerate group'), 'the degenerate group itself is still reported')
			assert(names.includes('healthy sibling'), 'an unrelated sibling is unaffected')
		},
	},
	{
		name: 'low-level package and geometry helpers are available',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('Parts', { x: 1, y: 1, w: 1, h: 0.4 })
			})

			const pptxPackage = await loadPptxPackage(buf)
			const parts = listPptxParts(pptxPackage)
			assert(parts.includes('ppt/presentation.xml'), 'expected presentation part')
			assert(parts.includes('ppt/slides/slide1.xml'), 'expected slide part')
			assertEqual(boxAnchor({ x: 1, y: 2, w: 3, h: 4 }, 'right', 'x'), 4, 'right anchor')
			assertEqual(boxAnchor({ x: 1, y: 2, w: 3, h: 4 }, 'middle', 'y'), 4, 'middle anchor')
			assertEqual(overlapArea({ x: 0, y: 0, w: 2, h: 2 }, { x: 1, y: 1, w: 2, h: 2 }), 1, 'overlap area')
		},
	},
	{
		name: 'loadPptxPackage accepts a filesystem path (string), and names the path on error',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('On disk', { x: 1, y: 1, w: 1, h: 0.4 })
			})

			const dir = mkdtempSync(join(tmpdir(), 'pptx-inspect-'))
			const filePath = join(dir, 'deck.pptx')
			try {
				writeFileSync(filePath, buf)

				// A string input is a filesystem path, read from disk — NOT latin1
				// binary content (the old JSZip footgun that turned a path into garbage
				// bytes → "Not a valid ZIP archive").
				const fromPath = await loadPptxPackage(filePath)
				assert(listPptxParts(fromPath).includes('ppt/slides/slide1.xml'), 'path input loads the slide part')

				let missingError = null
				try {
					await loadPptxPackage(join(dir, 'does-not-exist.pptx'))
				} catch (err) {
					missingError = err
				}
				assert(missingError, 'a missing path throws')
				assert(
					missingError.message.includes('does-not-exist.pptx'),
					'the error names the offending path rather than an opaque zip error'
				)
				assert(
					!missingError.message.includes('Not a valid ZIP archive'),
					'a missing path is not misreported as a corrupt archive'
				)
			} finally {
				rmSync(dir, { recursive: true, force: true })
			}
		},
	},
	{
		name: 'readPptxBinaryPart returns embedded media bytes; text and binary agree',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addImage({ data: `image/png;base64,${PNG_1X1}`, x: 1, y: 1, w: 1, h: 1 })
			})

			const pptxPackage = await loadPptxPackage(buf)
			const pngPath = listPptxParts(pptxPackage).find((part) => part.startsWith('ppt/media/') && part.endsWith('.png'))
			assert(pngPath, 'expected an embedded png media part')

			const bytes = await readPptxBinaryPart(pptxPackage, pngPath)
			assert(bytes instanceof Uint8Array, 'binary part is a Uint8Array')
			assertEqual(Buffer.from(bytes).toString('base64'), PNG_1X1, 'png bytes round-trip')
			// PNG magic number (\x89 P N G) survives undecoded — UTF-8 decoding would corrupt it.
			assertEqual(Buffer.from(bytes.subarray(1, 4)).toString('latin1'), 'PNG', 'png signature intact')

			const xml = await readPptxTextPart(pptxPackage, 'ppt/slides/slide1.xml')
			const xmlBytes = await readPptxBinaryPart(pptxPackage, 'ppt/slides/slide1.xml')
			assertEqual(new TextDecoder('utf-8').decode(xmlBytes), xml, 'text and binary reads agree')
			assertEqual(await readPptxBinaryPart(pptxPackage, 'ppt/does-not-exist.bin'), null, 'missing part is null')
		},
	},
])
