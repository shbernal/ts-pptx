import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertContentTypeDefault,
	assertNoContentTypeDefault,
	assertContentTypeOverride,
	contentTypeDefaultExtensions,
} from '../helpers.js'

// 1x1 PNG (red pixel)
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='
// 1x1 JPEG
const JPG_DATA =
	'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AL+AB//Z'

const CHART_FREE_MEDIA_DEFAULTS = ['jpeg', 'jpg', 'svg', 'gif', 'm4v', 'mp4', 'vml', 'xlsx']
const EMPTY_DECK_MEDIA_DEFAULTS = ['png', ...CHART_FREE_MEDIA_DEFAULTS]

function assertNoDefaults(xml, extensions) {
	for (const extension of extensions) assertNoContentTypeDefault(xml, extension)
}

defineRegressionSuite('Content type defaults', 'legacy bug-16', [
	{
		name: 'PNG-only deck emits png Default but not jpeg/jpg/svg/gif/m4v/mp4/vml/xlsx',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'xml')
			assertContentTypeDefault(xml, 'rels')
			assertContentTypeDefault(xml, 'png')
			assertNoDefaults(xml, CHART_FREE_MEDIA_DEFAULTS)
		},
	},
	{
		name: 'empty deck emits only xml + rels Defaults (no media defaults)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide()
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'xml')
			assertContentTypeDefault(xml, 'rels')
			assertNoDefaults(xml, EMPTY_DECK_MEDIA_DEFAULTS)
		},
	},
	{
		name: 'PNG + JPEG deck emits both png and jpeg/jpg Defaults; gif/svg absent',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
				s.addImage({ data: JPG_DATA, x: 3, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			const defaults = contentTypeDefaultExtensions(xml)
			assert(defaults.includes('png'), 'expected png Default; got: ' + defaults.join(', '))
			// jpeg images are recorded with extn "jpg" and type "image/jpg" in this codebase
			assert(
				defaults.includes('jpg') || defaults.includes('jpeg'),
				'expected jpg or jpeg Default for JPEG image; got: ' + defaults.join(', ')
			)
			assertNoDefaults(xml, ['gif', 'svg', 'm4v', 'mp4', 'vml', 'xlsx'])
		},
	},
	{
		name: 'chart deck emits xlsx Default',
		fn: async () => {
			const { pres, zip } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [{ name: 'series1', labels: ['a', 'b'], values: [1, 2] }], { x: 1, y: 1, w: 4, h: 3 })
			})
			void pres
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'xlsx')
		},
	},
	{
		name: 'regression - structural Override entries still emitted',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			for (const partName of [
				'/ppt/presentation.xml',
				'/ppt/slideMasters/slideMaster1.xml',
				'/ppt/theme/theme1.xml',
				'/ppt/theme/theme2.xml',
				'/ppt/slides/slide1.xml',
				'/ppt/slideLayouts/slideLayout1.xml',
				'/ppt/notesMasters/notesMaster1.xml',
				'/ppt/notesSlides/notesSlide1.xml',
			]) {
				assertContentTypeOverride(xml, partName)
			}
		},
	},
])
