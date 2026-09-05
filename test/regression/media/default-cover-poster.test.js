/**
 * `addMedia` no longer names the default play-button poster. The artwork is 54 kB of PNG that a
 * deck carrying no media never needs, so the rel goes out flagged and empty and the async media
 * pass fills the bytes in.
 *
 * That seam is invisible to every other gate. A flagged rel nobody resolves keeps an empty
 * `data`; the package assembler skips a rel whose payload it cannot decode; and the slide still
 * emits a `<p:pic>` whose `r:embed` points at a poster part that was never written. PowerPoint
 * calls a dangling part reference a repair, and nothing upstream of it would have failed. So the
 * bytes are asserted at the part, not the rel.
 */
import JSZip from 'jszip'
import { TsPptx, assert, assertEqual, build, defineRegressionSuite } from '../../helpers.js'

/** Smaller than the play-button artwork (55,784 bytes), far larger than any 1x1 test PNG. */
const POSTER_FLOOR = 40_000

const LINK = 'https://www.youtube.com/embed/Dph6ynRVyUc'

/** Every `ppt/media/*.png` part of a built deck, as `{ name, bytes }`, largest first. */
async function pngParts(zip) {
	const names = Object.keys(zip.files).filter((name) => /^ppt\/media\/.*\.png$/.test(name))
	const parts = await Promise.all(
		names.map(async (name) => ({ name, bytes: await zip.file(name).async('uint8array') }))
	)
	return parts.sort((a, b) => b.bytes.length - a.bytes.length)
}

defineRegressionSuite('Default video poster', [
	{
		name: 'an embedded video with no cover still emits the poster part',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 4, h: 3 })
			})
			const parts = await pngParts(zip)
			assertEqual(parts.length, 1, `exactly one poster part; got ${parts.map((p) => p.name).join(' ')}`)
			assert(
				parts[0].bytes.length > POSTER_FLOOR,
				`poster part is ${parts[0].bytes.length} bytes; the deferred artwork never arrived`
			)
		},
	},
	{
		name: 'an online video with no cover still emits the poster part',
		fn: async () => {
			// The online branch pushes its cover rel separately, so it can be missed on its own.
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'online', link: LINK, x: 1, y: 1, w: 4, h: 3 })
			})
			const parts = await pngParts(zip)
			assertEqual(parts.length, 1, `exactly one poster part; got ${parts.map((p) => p.name).join(' ')}`)
			assert(parts[0].bytes.length > POSTER_FLOOR, `poster part is ${parts[0].bytes.length} bytes`)
		},
	},
	{
		name: 'a caller-supplied cover is used instead, and the default is not embedded',
		fn: async () => {
			const cover =
				'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'online', link: LINK, cover, x: 1, y: 1, w: 4, h: 3 })
			})
			const parts = await pngParts(zip)
			assertEqual(parts.length, 1, `exactly one poster part; got ${parts.map((p) => p.name).join(' ')}`)
			assert(
				parts[0].bytes.length < POSTER_FLOOR,
				`the caller's cover was replaced by the default (${parts[0].bytes.length} bytes)`
			)
		},
	},
	{
		name: 'writing the same deck twice resolves the poster both times',
		fn: async () => {
			// The media pass skips a rel that already has data, so a second write must find the
			// bytes the first one left rather than re-resolving into an empty rel.
			const pres = new TsPptx()
			pres.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 4, h: 3 })

			const first = await pngParts(await JSZip.loadAsync(await pres.toBytes()))
			const second = await pngParts(await JSZip.loadAsync(await pres.toBytes()))
			assertEqual(second.length, first.length, 'the same poster parts on both writes')
			assertEqual(second[0].bytes.length, first[0].bytes.length, 'the poster part is the same size on both writes')
			assert(second[0].bytes.length > POSTER_FLOOR, 'the second write emitted an empty poster')
		},
	},
])
