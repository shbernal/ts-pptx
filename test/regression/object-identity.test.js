import { vi } from 'vitest'
import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertNonVisualDrawingProperty,
	xmlAttributes,
	xmlOpeningTags,
} from '../helpers.js'

const PNG_1X1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

defineRegressionSuite('Object identity', 'legacy bug-21', [
	{
		name: 'explicit objectName values are emitted as cNvPr names for slide objects',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Named text', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'identity:text' })
				slide.addShape(p.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 1, h: 0.4, objectName: 'identity:shape' })
				slide.addImage({
					data: `image/png;base64,${PNG_1X1}`,
					x: 1.7,
					y: 0.9,
					w: 0.4,
					h: 0.4,
					objectName: 'identity:image',
					altText: 'Identity image',
				})
				slide.addChart(p.charts.BAR, [{ name: 'Series 1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 2.4,
					y: 0.4,
					w: 2,
					h: 1.2,
					objectName: 'identity:chart',
					altText: 'Identity chart',
				})
				slide.addTable([[{ text: 'A1' }, { text: 'B1' }]], {
					x: 4.8,
					y: 0.4,
					w: 2,
					h: 0.6,
					objectName: 'identity:table',
				})
				slide.addMedia({
					type: 'video',
					data: 'video/mp4;base64,AAAA',
					x: 7.2,
					y: 0.4,
					w: 1,
					h: 0.8,
					objectName: 'identity:media',
				})
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			for (const name of [
				'identity:text',
				'identity:shape',
				'identity:image',
				'identity:chart',
				'identity:table',
				'identity:media',
			]) {
				assertNonVisualDrawingProperty(xml, { name }, name)
			}
			assertNonVisualDrawingProperty(xml, { name: 'identity:image', descr: 'Identity image' }, 'image altText')
			assertNonVisualDrawingProperty(xml, { name: 'identity:chart', descr: 'Identity chart' }, 'chart altText')
		},
	},
	{
		name: 'altText is emitted as cNvPr descr for text, shapes, tables, and media',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Named text', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'alt:text', altText: 'Text alt' })
				slide.addShape(p.shapes.RECTANGLE, {
					x: 0.4,
					y: 0.9,
					w: 1,
					h: 0.4,
					objectName: 'alt:shape',
					altText: 'Shape alt',
				})
				slide.addTable([[{ text: 'A1' }]], {
					x: 4.8,
					y: 0.4,
					w: 2,
					h: 0.6,
					objectName: 'alt:table',
					altText: 'Table alt',
				})
				slide.addMedia({
					type: 'video',
					data: 'video/mp4;base64,AAAA',
					x: 7.2,
					y: 0.4,
					w: 1,
					h: 0.8,
					objectName: 'alt:media',
					altText: 'Media alt',
				})
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNonVisualDrawingProperty(xml, { name: 'alt:text', descr: 'Text alt' }, 'text altText')
			assertNonVisualDrawingProperty(xml, { name: 'alt:shape', descr: 'Shape alt' }, 'shape altText')
			assertNonVisualDrawingProperty(xml, { name: 'alt:table', descr: 'Table alt' }, 'table altText')
			assertNonVisualDrawingProperty(xml, { name: 'alt:media', descr: 'Media alt' }, 'media altText')
		},
	},
	{
		name: 'default cNvPr names are emitted when objectName is omitted',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Text', { x: 0.4, y: 0.3, w: 2, h: 0.4 })
				slide.addShape(p.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 1, h: 0.4 })
				slide.addImage({ data: `image/png;base64,${PNG_1X1}`, x: 1.7, y: 0.9, w: 0.4, h: 0.4 })
				slide.addTable([[{ text: 'A1' }]], { x: 4.8, y: 0.4, w: 2, h: 0.6 })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const names = xmlOpeningTags(xml, 'p:cNvPr').map((tag) => xmlAttributes(tag).name)
			for (const prefix of ['Text', 'Shape', 'Image', 'Table']) {
				assert(
					names.some((name) => name && name.startsWith(prefix + ' ')),
					`expected a default ${prefix} name; got: ${names.join(', ')}`
				)
			}
		},
	},
	{
		name: 'invalid object names warn without throwing',
		fn: async () => {
			const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})
			try {
				await build((p) => {
					const slide = p.addSlide()
					slide.addText('Blank name', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: '   ' })
					slide.addShape(p.shapes.RECTANGLE, { x: 0.4, y: 0.9, w: 1, h: 0.4, objectName: 'ctrlname' })
				})
				const messages = warnSpy.mock.calls.map((call) => String(call[0]))
				assert(
					messages.some((m) => m.includes('empty or whitespace-only')),
					`expected empty-name warning; got: ${messages.join(' | ')}`
				)
				assert(
					messages.some((m) => m.includes('control characters')),
					`expected control-char warning; got: ${messages.join(' | ')}`
				)
			} finally {
				warnSpy.mockRestore()
			}
		},
	},
	{
		name: 'duplicate object names on a single slide warn',
		fn: async () => {
			const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})
			try {
				await build((p) => {
					const slide = p.addSlide()
					slide.addText('First', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'dupe:name' })
					slide.addText('Second', { x: 0.4, y: 0.9, w: 2, h: 0.4, objectName: 'dupe:name' })
				})
				const messages = warnSpy.mock.calls.map((call) => String(call[0]))
				assert(
					messages.some((m) => m.includes('duplicate objectName') && m.includes('dupe:name')),
					`expected duplicate-name warning; got: ${messages.join(' | ')}`
				)
			} finally {
				warnSpy.mockRestore()
			}
		},
	},
	{
		name: 'slide master placeholder objectName is emitted on inherited placeholder shapes',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'OBJECT_IDENTITY_MASTER',
					objects: [
						{
							placeholder: {
								options: {
									name: 'title',
									type: 'title',
									x: 0.5,
									y: 0.5,
									w: 5,
									h: 0.7,
									objectName: 'identity:placeholder:title',
								},
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterName: 'OBJECT_IDENTITY_MASTER' })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNonVisualDrawingProperty(xml, { name: 'identity:placeholder:title' }, 'placeholder objectName')
			const placeholder = xmlOpeningTags(xml, 'p:ph').find((tag) => {
				const attrs = xmlAttributes(tag)
				return attrs.idx === '100' && attrs.type === 'title'
			})
			assert(placeholder, 'expected title placeholder metadata; got: ' + xml)
		},
	},
])
