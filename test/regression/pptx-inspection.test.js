import { inspectPptx, boxAnchor, listPptxParts, loadPptxPackage, overlapArea } from '../../dist/inspect.js'
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
])
