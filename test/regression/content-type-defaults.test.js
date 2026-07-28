import { ChartType } from '../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
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

// Minimal SVG, supplied as bytes so nothing has to be read from disk.
const SVG_DATA =
	'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxIiBoZWlnaHQ9IjEiLz4='

// The A/V content type is resolved from the media *extension*, so the payload is never
// decoded to decide it — four bytes are enough (same shortcut as media-loop.test.js).
const AV_DATA = 'base64,AAAA'

// What PowerPoint itself authors for each embedded media extension. `mp4`, `mpg` and `mpeg`
// are the only entries that depend on whether the item is audio or video, so they appear in
// both tables with different expected content types — that difference is the point.
// The last row of each table is an extension the mapping does not list, which falls through
// to `<mtype>/<extn>`.
const VIDEO_CONTENT_TYPES = [
	['mp4', 'video/mp4'],
	['m4v', 'video/mp4'],
	['mov', 'video/quicktime'],
	['avi', 'video/avi'],
	['wmv', 'video/x-ms-wmv'],
	['mpg', 'video/mpeg'],
	['mpeg', 'video/mpeg'],
	['ogv', 'video/ogg'],
	['webm', 'video/webm'],
	['3gp', 'video/3gp'],
]

const AUDIO_CONTENT_TYPES = [
	['mp4', 'audio/mp4'],
	['mpg', 'audio/mpeg'],
	['mpeg', 'audio/mpeg'],
	['mp3', 'audio/mpeg'],
	['m4a', 'audio/mp4'],
	['wav', 'audio/x-wav'],
	['wma', 'audio/x-ms-wma'],
	['aac', 'audio/aac'],
	['oga', 'audio/ogg'],
	['ogg', 'audio/ogg'],
	['flac', 'audio/flac'],
	['opus', 'audio/opus'],
]

/** One deck carrying one media item per row of `table`; returns its `[Content_Types].xml`. */
async function buildMediaDeck(type, table) {
	const { zip } = await build((p) => {
		const s = p.addSlide()
		table.forEach(([extn], idx) => {
			// `extn` is a documented MediaProps option and wins over the data-URI mime sniff,
			// so each row is one media rel with exactly the extension under test.
			s.addMedia({ type, extn, data: `${type}/${extn};${AV_DATA}`, x: 0.5, y: 0.25 * idx + 0.25, w: 1, h: 0.2 })
		})
	})
	return readEntry(zip, '[Content_Types].xml')
}

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
		// The `image/jpg` mime is not the IANA spelling but is widespread in the wild; the deck
		// must still declare `image/jpeg`, which is what PowerPoint authors for a `.jpg` part.
		name: 'an image/jpg data mime keeps the jpg extension but declares image/jpeg',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: JPG_DATA.replace('data:image/jpeg;', 'data:image/jpg;'), x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'jpg')
			assertEqual(contentTypeForExtension(xml, 'jpg'), 'image/jpeg', 'jpg Default ContentType')
		},
	},
	{
		// A background image's extension comes from its `path`, so a `.svg` background is the one
		// route that reaches the svg mapping (`addImage` registers its SVG rel with the content
		// type spelled out, and SVG image *fills* are rejected outright).
		name: 'SVG background emits an svg Default with image/svg+xml',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'svg-bkgd', background: { path: 'assets/logo.svg', data: SVG_DATA } })
				p.addSlide({ masterTitle: 'svg-bkgd' })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'svg')
			assertEqual(contentTypeForExtension(xml, 'svg'), 'image/svg+xml', 'svg Default ContentType')
		},
	},
	{
		// A background used to take its extension from `path` alone. With no path, `addBackground()`
		// substituted the `preencoded.png` placeholder, so bytes of any other format were embedded in
		// a `.png` part the package declared as `image/png` — the exact Default/payload mismatch
		// PowerPoint offers to "repair". Reachable from a plain `background: { data }` call.
		name: 'a data-only SVG background declares image/svg+xml, not image/png',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'data-only-svg', background: { data: SVG_DATA } })
				p.addSlide({ masterTitle: 'data-only-svg' })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertContentTypeDefault(xml, 'svg')
			assertEqual(contentTypeForExtension(xml, 'svg'), 'image/svg+xml', 'svg Default ContentType')
			assertNoContentTypeDefault(xml, 'png')
			const media = listEntries(zip).filter((name) => name.includes('media/'))
			assertEqual(media.length, 1, 'expected exactly one media part')
			assert(media[0].endsWith('.svg'), `expected an .svg media part; got ${media[0]}`)
		},
	},
	{
		// The mime carries the same weight for every other format: a data-only GIF background must
		// not be announced as PNG either.
		name: 'a data-only GIF background declares image/gif',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'data-only-gif',
					background: { data: 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==' },
				})
				p.addSlide({ masterTitle: 'data-only-gif' })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertEqual(contentTypeForExtension(xml, 'gif'), 'image/gif', 'gif Default ContentType')
			assertNoContentTypeDefault(xml, 'png')
		},
	},
	{
		// Precedence when the two sources disagree: the bytes decide, matching `addImage()`. A
		// caller who names the file `.png` but hands over SVG has still handed over SVG.
		name: 'a background data: mime wins over a disagreeing path extension',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'mime-wins', background: { path: 'assets/logo.png', data: SVG_DATA } })
				p.addSlide({ masterTitle: 'mime-wins' })
			})
			const xml = await readEntry(zip, '[Content_Types].xml')
			assertEqual(contentTypeForExtension(xml, 'svg'), 'image/svg+xml', 'svg Default ContentType')
			assertNoContentTypeDefault(xml, 'png')
		},
	},
	{
		// A Default whose ContentType disagrees with what PowerPoint expects for that extension is
		// exactly the mismatch that produces the "repair" prompt, so the whole table is asserted
		// rather than a representative sample.
		name: 'video extensions declare the content types PowerPoint authors',
		fn: async () => {
			const xml = await buildMediaDeck('video', VIDEO_CONTENT_TYPES)
			for (const [extn, contentType] of VIDEO_CONTENT_TYPES) {
				assertEqual(contentTypeForExtension(xml, extn), contentType, `${extn} Default ContentType (video)`)
			}
		},
	},
	{
		name: 'audio extensions declare the content types PowerPoint authors',
		fn: async () => {
			const xml = await buildMediaDeck('audio', AUDIO_CONTENT_TYPES)
			for (const [extn, contentType] of AUDIO_CONTENT_TYPES) {
				assertEqual(contentTypeForExtension(xml, extn), contentType, `${extn} Default ContentType (audio)`)
			}
		},
	},
	{
		name: 'chart deck emits xlsx Default',
		fn: async () => {
			const { pres, zip } = await build((p) => {
				const s = p.addSlide()
				s.addChart([{ name: 'series1', labels: ['a', 'b'], values: [1, 2] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 4,
					h: 3,
				})
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
