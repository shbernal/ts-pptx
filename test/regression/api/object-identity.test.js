import { ChartType, ShapeType } from '../../../dist/node.js'
import { vi } from 'vitest'
import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertNotIncludes,
	assertNonVisualDrawingProperty,
	xmlAttributes,
	xmlOpeningTags,
	listEntries,
} from '../../helpers.js'

const PNG_1X1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

defineRegressionSuite('Object identity [legacy bug-21]', [
	{
		name: 'explicit objectName values are emitted as cNvPr names for slide objects',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Named text', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'identity:text' })
				slide.addShape(ShapeType.rect, { x: 0.4, y: 0.9, w: 1, h: 0.4, objectName: 'identity:shape' })
				slide.addImage({
					data: `image/png;base64,${PNG_1X1}`,
					x: 1.7,
					y: 0.9,
					w: 0.4,
					h: 0.4,
					objectName: 'identity:image',
					altText: 'Identity image',
				})
				slide.addChart([{ name: 'Series 1', labels: ['A', 'B'], values: [1, 2] }], {
					type: ChartType.bar,
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
		// fork-table-cnvpr-id-collision: a table's cNvPr id must come from the same
		// per-slide `idx + 2` space as every other shape. The legacy table formula
		// (`tableCounter * slideNumber + 1`) could equal another shape's `idx + 2` on the
		// same slide — e.g. on slide 2 a table (→ 1*2+1 = 3) followed by a text box
		// (idx 1 → 1+2 = 3) — yielding a duplicate cNvPr id. PowerPoint then reports the
		// deck as corrupt/unreadable (0x80070570); LibreOffice silently tolerates it.
		name: 'table cNvPr id does not collide with sibling shapes on the same slide',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('slide one', { x: 0.5, y: 0.5, w: 3, h: 0.5 })
				const slide = p.addSlide() // slide 2 — triggers the legacy table-id collision
				slide.addTable([[{ text: 'A1' }, { text: 'B1' }]], { x: 0.5, y: 0.5, w: 4, h: 0.6 })
				slide.addText('sibling text', { x: 0.5, y: 1.5, w: 3, h: 0.5 })
				slide.addShape(ShapeType.rect, { x: 0.5, y: 2.5, w: 2, h: 0.5 })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			const ids = xmlOpeningTags(xml, 'p:cNvPr').map((tag) => xmlAttributes(tag).id)
			assert(ids.length >= 4, `expected the spTree id plus table/text/shape ids; got: ${ids.join(', ')}`)
			assert(
				new Set(ids).size === ids.length,
				`expected unique cNvPr ids on the slide; got duplicates in: ${ids.join(', ')}`
			)
		},
	},
	{
		// fork-slide-number-placeholder-hardcoded-id: the slide-number placeholder formerly emitted a
		// hardcoded `<p:cNvPr id="25">`. Object ids are otherwise allocated `idx + 2` from
		// `_slideObjects`, so a slide with 24 top-level objects gives its 24th object (idx 23) id 25 too
		// — a duplicate cNvPr id PowerPoint repairs (0x80070570). The id now comes from the same
		// monotonic counter, so it cannot alias any shape regardless of slide population.
		name: 'slide-number placeholder id does not collide with a populous slide (formerly hardcoded 25)',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.slideNumber = { x: 0.5, y: '90%' }
				for (let i = 0; i < 24; i++) {
					// idx 0..23 -> ids 2..25; object 23 lands on the old hardcoded slide-number id 25
					slide.addShape(ShapeType.rect, { x: 0.2, y: 0.2 + i * 0.1, w: 1, h: 0.08 })
				}
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const ids = xmlOpeningTags(xml, 'p:cNvPr').map((tag) => Number(xmlAttributes(tag).id))
			assert(new Set(ids).size === ids.length, `expected unique cNvPr ids; got duplicates in: ${ids.join(', ')}`)

			const snTag = xmlOpeningTags(xml, 'p:cNvPr').find(
				(tag) => xmlAttributes(tag).name === 'Slide Number Placeholder 0'
			)
			assert(snTag, 'expected a slide-number placeholder cNvPr; got: ' + xml)
			const snId = Number(xmlAttributes(snTag).id)
			assert(
				snId === Math.max(...ids),
				`expected the slide-number id to be the highest (next free) id; got ${snId} among ${ids.join(', ')}`
			)
		},
	},
	{
		// Group children allocate ids PAST `_slideObjects.length`, so the slide-number id must be taken
		// after the whole walk (top-level objects + group children), not from the top-level count alone.
		name: 'slide-number placeholder id is allocated past group-child ids',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.slideNumber = { x: 0.5, y: '90%' }
				slide.addShape(ShapeType.rect, { x: 0.2, y: 0.2, w: 1, h: 0.4 }) // idx 0 -> id 2
				// group is idx 1 -> id 3; its children seed past length (2) -> ids 4, 5, 6
				slide.addGroup([
					{ rect: { x: 1, y: 1, w: 1, h: 1 } },
					{ rect: { x: 2, y: 1, w: 1, h: 1 } },
					{ rect: { x: 3, y: 1, w: 1, h: 1 } },
				])
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const ids = xmlOpeningTags(xml, 'p:cNvPr').map((tag) => Number(xmlAttributes(tag).id))
			assert(
				new Set(ids).size === ids.length,
				`expected unique cNvPr ids across the nested tree; got: ${ids.join(', ')}`
			)

			const snTag = xmlOpeningTags(xml, 'p:cNvPr').find(
				(tag) => xmlAttributes(tag).name === 'Slide Number Placeholder 0'
			)
			assert(snTag, 'expected a slide-number placeholder cNvPr; got: ' + xml)
			assert(
				Number(xmlAttributes(snTag).id) === Math.max(...ids),
				`expected the slide-number id past every group-child id; got ${xmlAttributes(snTag).id} among ${ids.join(', ')}`
			)
		},
	},
	{
		name: 'altText is emitted as cNvPr descr for text, shapes, tables, and media',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Named text', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'alt:text', altText: 'Text alt' })
				slide.addShape(ShapeType.rect, {
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
		// fork-addtext-objectname-double-escape: Slide.addText(string, opts) wraps the bare string as
		// `[{ text, options }]`, reusing `options` as both the shape-level opts and the lone run's
		// opts. Escaping objectName inside the per-run pass (as well as the shape-level pass) encoded
		// an already-escaped `&` a second time: 'Q&A' -> name="Q&amp;amp;A" instead of "Q&amp;A".
		name: 'addText(string, {objectName}) escapes objectName exactly once',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Q&A', { x: 0.4, y: 0.3, w: 2, h: 0.4, objectName: 'Q&A' })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNonVisualDrawingProperty(xml, { name: 'Q&amp;A' }, 'single-escaped objectName')
			assertNotIncludes(xml, '&amp;amp;', 'objectName must not be double-escaped')
		},
	},
	{
		name: 'default cNvPr names are emitted when objectName is omitted',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Text', { x: 0.4, y: 0.3, w: 2, h: 0.4 })
				slide.addShape(ShapeType.rect, { x: 0.4, y: 0.9, w: 1, h: 0.4 })
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
					slide.addShape(ShapeType.rect, { x: 0.4, y: 0.9, w: 1, h: 0.4, objectName: 'ctrlname' })
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
				p.addSlide({ masterTitle: 'OBJECT_IDENTITY_MASTER' })
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
	{
		// fork-placeholder-objectname-collision: a master/layout placeholder without an explicit
		// objectName defaults to its declared name (then type, then idx) rather than the plain
		// text-box counter, which counts only `_type === text` objects and so would tag every
		// placeholder with a duplicate `Text 1`. Two named placeholders must get distinct
		// Selection Pane identities and must not fire the duplicate-objectName warning.
		name: 'master placeholders default to distinct names (no duplicate objectName)',
		fn: async () => {
			const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {})
			try {
				const { zip } = await build((p) => {
					p.defineSlideMaster({
						title: 'PH_IDENTITY_MASTER',
						objects: [
							{ placeholder: { options: { name: 'title-ph', type: 'title', x: 0.5, y: 0.3, w: 9, h: 1 }, text: '' } },
							{
								placeholder: {
									options: { name: 'body-ph', type: 'body', idx: 1, x: 0.5, y: 1.5, w: 9, h: 4 },
									text: '',
								},
							},
						],
					})
					p.addSlide({ masterTitle: 'PH_IDENTITY_MASTER' })
				})

				const messages = warnSpy.mock.calls.map((call) => String(call[0]))
				assert(
					!messages.some((m) => m.includes('duplicate objectName')),
					`expected no duplicate-objectName warning; got: ${messages.join(' | ')}`
				)

				// Placeholders are emitted on the master's layout part.
				const layoutNames = listEntries(zip).filter((n) => /ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))
				const layoutXmls = await Promise.all(layoutNames.map((n) => readEntry(zip, n)))
				const layoutXml = layoutXmls.find((xml) => xml.includes('name="title-ph"'))
				assert(layoutXml, `expected a layout carrying the placeholder default names; got: ${layoutNames.join(', ')}`)
				const phNames = xmlOpeningTags(layoutXml, 'p:cNvPr')
					.map((tag) => xmlAttributes(tag).name)
					.filter((n) => n === 'title-ph' || n === 'body-ph')
				assert(
					phNames.includes('title-ph') && phNames.includes('body-ph'),
					`expected both placeholder default names; got: ${phNames.join(', ')}`
				)
				assert(
					new Set(phNames).size === phNames.length,
					`expected distinct placeholder names; got: ${phNames.join(', ')}`
				)
			} finally {
				warnSpy.mockRestore()
			}
		},
	},
	{
		// The index base used to differ per definer: six kinds counted from 0 and four from 1,
		// and `addChart` did neither — it derived `Chart 0` by counting the chart objects already
		// on the slide. They are all 1-based now, which is the base PowerPoint itself uses, so a
		// definer added later cannot pick the other convention by accident.
		name: 'every kind numbers its default objectName from 1',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('t', { x: 0.4, y: 0.3, w: 2, h: 0.4 })
				slide.addShape(ShapeType.rect, { x: 0.4, y: 0.9, w: 1, h: 0.4 })
				slide.addImage({ data: `image/png;base64,${PNG_1X1}`, x: 1.7, y: 0.9, w: 0.4, h: 0.4 })
				slide.addConnector({ type: 'straight', x1: 3, y1: 1, x2: 4, y2: 1 })
				slide.addTable([['a']], { x: 0.4, y: 1.5, w: 2 })
				slide.addChart([{ name: 's', labels: ['a'], values: [1] }], { type: ChartType.bar, x: 3, y: 2, w: 3, h: 2 })
				slide.addGroup([{ rect: { x: 6, y: 1, w: 1, h: 1 } }])
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const names = xmlOpeningTags(xml, 'p:cNvPr').map((tag) => xmlAttributes(tag).name)
			// `addShape` and `addText` share the `text` bucket, so the shape is `Shape 2`; the
			// group's own child took `Shape 3` before the group itself was named.
			for (const expected of ['Text 1', 'Shape 2', 'Image 1', 'Connector 1', 'Table 1', 'Chart 1', 'Group 1']) {
				assert(names.includes(expected), `expected a default name ${expected}; got: ${names.join(', ')}`)
			}
			assert(!names.some((n) => n.endsWith(' 0')), `expected no 0-suffixed default name; got: ${names.join(', ')}`)
		},
	},
])
