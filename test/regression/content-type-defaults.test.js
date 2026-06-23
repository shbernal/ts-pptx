import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertContentTypeDefault,
	assertNoContentTypeDefault,
	assertContentTypeOverride,
	contentTypeDefaultExtensions,
	contentTypeForExtension,
} from '../helpers.js'

// 1x1 PNG (red pixel)
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='
// 1x1 JPEG
const JPG_DATA =
	'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AL+AB//Z'
// Minimal EMF/WMF payloads (bytes are immaterial to content-type generation; w/h are
// supplied so no intrinsic-size measurement is attempted). The `image/emf` / `image/wmf`
// mime in the data URI is what the extension sniff reads back as the `emf` / `wmf` extn.
const EMF_DATA = 'data:image/emf;base64,AQAAAA=='
const WMF_DATA = 'data:image/wmf;base64,1tZ0AA=='

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
			// JPEG images keep extn "jpeg" from the data: mime sniff; the content type is image/jpeg
			assert(defaults.includes('jpeg'), 'expected jpeg Default for JPEG image; got: ' + defaults.join(', '))
			assertEqual(contentTypeForExtension(xml, 'jpeg'), 'image/jpeg', 'jpeg Default ContentType')
			assertNoDefaults(xml, ['gif', 'svg', 'm4v', 'mp4', 'vml', 'xlsx'])
		},
	},
	{
		name: 'EMF image emits emf Default with OOXML-correct image/x-emf content type',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: EMF_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'emf')
			assertEqual(contentTypeForExtension(xml, 'emf'), 'image/x-emf', 'emf Default ContentType')
		},
	},
	{
		name: 'WMF image emits wmf Default with OOXML-correct image/x-wmf content type',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: WMF_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'wmf')
			assertEqual(contentTypeForExtension(xml, 'wmf'), 'image/x-wmf', 'wmf Default ContentType')
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
