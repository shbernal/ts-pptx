import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// upstream-issue-1434: `loop`/`loopCount` make embedded audio/video repeat.
// PowerPoint stores looping as `repeatCount` on the media node's <p:cTn> inside a
// slide-level <p:timing> tree, targeting the picture by its <p:cNvPr> id (spid).
defineRegressionSuite('Media looping', [
	{
		name: 'loop:true emits a p:timing tree with repeatCount="indefinite"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('<p:timing>'), 'expected a <p:timing> tree')
			assert(xml.includes('nodeType="tmRoot"'), 'expected the tmRoot time node')
			assert(xml.includes('<p:video>'), 'expected a <p:video> media node')
			assert(xml.includes('repeatCount="indefinite"'), 'expected repeatCount="indefinite" for loop:true')
		},
	},
	{
		name: 'loopCount maps to a finite repeatCount (N*1000)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loopCount: 3 })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('repeatCount="3000"'), 'expected loopCount:3 -> repeatCount="3000"')
		},
	},
	{
		name: 'media node targets the picture by spid (cNvPr id = mediaRid + 2)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// First media on slide 1: video rId=2, so the picture's cNvPr id = mediaRid + 2 = 4
			const cNvPrMatch = xml.match(/<p:cNvPr id="(\d+)"[^>]*><a:hlinkClick[^>]*action="ppaction:\/\/media"/)
			assert(cNvPrMatch, 'expected the media picture cNvPr with media hlink action')
			const spid = cNvPrMatch[1]
			assert(
				xml.includes(`<p:spTgt spid="${spid}"/>`),
				`expected timing target spid="${spid}" to match the media cNvPr id`
			)
		},
	},
	{
		name: 'no timing tree without loop options',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2 })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!xml.includes('<p:timing>'), 'expected no <p:timing> tree when media does not loop')
		},
	},
	{
		name: 'audio uses a:audioFile and a p:audio timing node',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'audio', data: 'audio/mp3;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('<a:audioFile r:link='), 'expected <a:audioFile> for audio media (not <a:videoFile>)')
			assert(!xml.includes('<a:videoFile'), 'expected no <a:videoFile> for audio media')
			assert(xml.includes('<p:audio>'), 'expected a <p:audio> timing node for looping audio')
			assert(!xml.includes('<p:video>'), 'expected no <p:video> timing node for audio')
		},
	},
	{
		name: 'video still uses a:videoFile and a p:video timing node',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('<a:videoFile r:link='), 'expected <a:videoFile> for video media')
			assert(!xml.includes('<a:audioFile'), 'expected no <a:audioFile> for video media')
			assert(xml.includes('<p:video>'), 'expected a <p:video> timing node for looping video')
		},
	},
	{
		name: 'multiple looping media share one p:timing tree',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
				s.addMedia({ type: 'video', data: 'video/mp4;base64,BBBB', x: 5, y: 1, w: 3, h: 2, loopCount: 2 })
			})

			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.match(/<p:timing>/g).length === 1, 'expected exactly one <p:timing> tree')
			assert((xml.match(/<p:video>/g) || []).length === 2, 'expected two <p:video> media nodes')
		},
	},
])
