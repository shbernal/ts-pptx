/**
 * A media item emits TWO relationships against one Target: the ECMA one first (`audio`, `video`)
 * and the Microsoft 2007 `media` one second. Which of the two a rel gets is decided by whether
 * that Target has been emitted already, and the probe answering that used to substring-scan the
 * emitted XML of EVERY rel on the slide -- hyperlinks included.
 *
 * That is not a hypothetical collision. An online video and a hyperlink to the same URL is an
 * ordinary thing to author, the hyperlink is emitted first, and its Target satisfied the media
 * probe: the video pair came out as two MS-media rels with no ECMA `video` rel at all, while
 * `<a:videoFile r:link>` went on pointing at the first of them. The probe is scoped to the media
 * loop now, which is the only collection the pairing lives in.
 */
import { assert, assertEqual, build, defineRegressionSuite, readEntry } from '../../helpers.js'

const VIDEO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

/** `{ rId: Type }` for every relationship on slide 1. */
async function slideRels(zip) {
	const xml = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
	return Object.fromEntries([...xml.matchAll(/Id="(rId\d+)" Type="([^"]+)"/g)].map((m) => [m[1], m[2]]))
}

const LINK = 'https://www.youtube.com/embed/Dph6ynRVyUc'

defineRegressionSuite('Media relationship pairing', [
	{
		name: 'an online video pairs one ECMA video rel with one MS media rel',
		fn: async () => {
			// The baseline: without this, the collision case below cannot mean anything.
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'online', link: LINK, x: 1, y: 3, w: 4, h: 3 })
			})
			const types = Object.values(await slideRels(zip))
			assertEqual(types.filter((t) => t === VIDEO_REL).length, 1, 'exactly one ECMA video rel')
			assertEqual(types.filter((t) => t === MS_MEDIA_REL).length, 1, 'exactly one MS media rel')
		},
	},
	{
		name: 'a hyperlink to the same URL does not steal the video rel',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('watch', { x: 1, y: 1, w: 3, h: 1, hyperlink: { url: LINK } })
				slide.addMedia({ type: 'online', link: LINK, x: 1, y: 3, w: 4, h: 3 })
			})
			const rels = await slideRels(zip)
			const types = Object.values(rels)
			assertEqual(types.filter((t) => t === VIDEO_REL).length, 1, `exactly one ECMA video rel; got ${types.join(' ')}`)
			assertEqual(types.filter((t) => t === MS_MEDIA_REL).length, 1, `exactly one MS media rel; got ${types.join(' ')}`)

			// And the slide's own reference resolves to the right one of the pair.
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const videoFile = /<a:videoFile[^>]*r:link="(rId\d+)"/.exec(xml)
			assert(videoFile, `expected an <a:videoFile r:link>; got: ${xml}`)
			assertEqual(rels[videoFile[1]], VIDEO_REL, '<a:videoFile> must point at the ECMA video rel')
		},
	},
	{
		name: 'two media items sharing a Target still pair with each other',
		fn: async () => {
			// The behaviour the probe exists for, and the half that must NOT change: within the
			// media collection, the second rel against a Target is still the MS one.
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addMedia({ type: 'audio', data: 'audio/mpeg;base64,AAAA', x: 1, y: 1, w: 2, h: 1 })
			})
			const types = Object.values(await slideRels(zip))
			assertEqual(types.filter((t) => t === AUDIO_REL).length, 1, `one ECMA audio rel; got ${types.join(' ')}`)
			assertEqual(types.filter((t) => t === MS_MEDIA_REL).length, 1, `one MS media rel; got ${types.join(' ')}`)
		},
	},
])
