import {
	inspectPptx,
	boxAnchor,
	listPptxParts,
	loadPptxPackage,
	overlapArea,
	readPptxBinaryPart,
	readPptxTextPart,
} from '../../dist/inspect.js'
import { defineRegressionSuite, build, assert, assertEqual } from '../helpers.js'

const PNG_1X1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

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
				slide.addShape(p.shapes.RECTANGLE, {
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
				// Custom zero insets via the `inset` option (inches).
				slide.addText('Tight', { x: 1, y: 4, w: 2, h: 0.5, inset: 0, objectName: 'fit:inset0' })
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

			const tight = elements.get('fit:inset0')
			assertEqual(tight.autofit, 'none', 'inset-only box still reports none')
			assertEqual(tight.bodyInsets.left, 0, 'zero inset is preserved, not defaulted')
			assertEqual(tight.bodyInsets.bottom, 0, 'zero inset is preserved, not defaulted')

			const image = elements.get('fit:image')
			assertEqual(image.autofit, null, 'image without a text frame has no autofit')
			assertEqual(image.bodyInsets, null, 'image without a text frame has no body insets')
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
