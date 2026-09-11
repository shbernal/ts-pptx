import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertEqual,
	contentTypeForExtension,
	selfClosingTags,
	xmlAttributes,
} from '../../helpers.js'

defineRegressionSuite('Slide backgrounds [legacy bug-12]', [
	{
		name: 'solid-color slide.background <p:bgPr> contains <a:effectLst/>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = { color: '0088CC' }
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<p:bg><p:bgPr><a:solidFill>[\s\S]*?<\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/.test(xml),
				'expected <p:bgPr> to contain <a:solidFill>...</a:solidFill><a:effectLst/>; got: ' + xml
			)
		},
	},
	{
		name: 'image-background still emits <a:effectLst/> (regression guard)',
		fn: async () => {
			// 1x1 transparent PNG
			const b64png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = { data: 'image/png;base64,' + b64png }
				s.addText('hi', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<p:bg><p:bgPr>[\s\S]*<a:effectLst\/><\/p:bgPr><\/p:bg>/.test(xml),
				'expected image bgPr to keep <a:effectLst/>; got: ' + xml
			)
		},
	},
	{
		// A background given only as a `path` has no `data:` mime to sniff, so the extension comes
		// off the file name — and a `.jpg` is deliberately renamed to `jpeg` on the way in. Both
		// spellings are the same JPEG bytes, but PowerPoint complains at startup about a package
		// that declares one and stores the other, so the part, the Default content type and the rel
		// Target all have to land on `jpeg` together. Nothing else exercises the path-only branch:
		// every other background fixture supplies `data`, which wins over `path`.
		name: 'a jpg background supplied by path is stored, typed and targeted as jpeg',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = { path: 'demos/common/images/cc_logo.jpg' }
			})
			const rels = selfClosingTags(await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels'), 'Relationship')
				.map((tag) => xmlAttributes(tag))
				.filter((attrs) => attrs.Type.endsWith('/image'))
			assertEqual(rels.length, 1, 'expected one background image rel')
			assert(rels[0].Target.endsWith('.jpeg'), `expected a .jpeg rel Target; got ${rels[0].Target}`)

			const part = rels[0].Target.replace(/^\.\./, 'ppt')
			assert(listEntries(zip).includes(part), `rel Target ${rels[0].Target} has no part at ${part}`)
			assertEqual(
				contentTypeForExtension(await readEntry(zip, '[Content_Types].xml'), 'jpeg'),
				'image/jpeg',
				'the jpeg Default content type'
			)

			// The bytes came off disk, not from an inlined placeholder.
			const bytes = await zip.file(part).async('uint8array')
			assert(bytes.length > 1000, `expected the jpg to be read from disk; got ${bytes.length} bytes`)
			assertEqual(bytes[0], 0xff, 'first JPEG SOI byte')
			assertEqual(bytes[1], 0xd8, 'second JPEG SOI byte')
		},
	},
	// An image background took its rel id from the media list alone, as `_relsMedia.length + 1`, and
	// named its part after the slide or layout title. So a hyperlink already holding `rId1` got a
	// second `rId1`, and two titles that sanitize alike ("A B", "A-B", or a master titled
	// "Slide 1" beside slide 1) wrote one media part that both backgrounds pointed at.
	{
		name: 'an image background beside a hyperlink takes a fresh rel id',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('link', { x: 1, y: 1, w: 3, h: 1, hyperlink: { url: 'https://example.invalid/' } })
				s.background = { data: PNG_A }
			})
			await assertBackgroundResolves(zip, 'ppt/slides/slide1.xml', PNG_A)
		},
	},
	{
		name: 'a layout image background beside a hyperlinked object takes a fresh rel id',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'LINKED',
					background: { data: PNG_A },
					objects: [
						{
							text: {
								text: 'link',
								options: { x: 1, y: 1, w: 3, h: 1, hyperlink: { url: 'https://example.invalid/' } },
							},
						},
					],
				})
				p.addSlide({ masterTitle: 'LINKED' })
			})
			await assertBackgroundResolves(zip, await layoutPartFor(zip, 'LINKED'), PNG_A)
		},
	},
	{
		name: 'layouts whose titles sanitize alike keep their own background images',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'A B', background: { data: PNG_A } })
				p.defineSlideMaster({ title: 'A-B', background: { data: PNG_B } })
				p.addSlide({ masterTitle: 'A B' })
				p.addSlide({ masterTitle: 'A-B' })
			})
			await assertBackgroundResolves(zip, await layoutPartFor(zip, 'A B'), PNG_A)
			await assertBackgroundResolves(zip, await layoutPartFor(zip, 'A-B'), PNG_B)
		},
	},
	{
		name: 'a layout titled like a slide keeps its own background image',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'Slide 1', background: { data: PNG_A } })
				const s = p.addSlide({ masterTitle: 'Slide 1' })
				s.background = { data: PNG_B }
			})
			await assertBackgroundResolves(zip, await layoutPartFor(zip, 'Slide 1'), PNG_A)
			await assertBackgroundResolves(zip, 'ppt/slides/slide1.xml', PNG_B)
		},
	},
	{
		name: 'reassigning an image background replaces its rel rather than adding one',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = { data: PNG_A }
				s.background = { data: PNG_B }
			})
			const images = relsOf(await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')).filter((rel) =>
				rel.Type.endsWith('/image')
			)
			assertEqual(images.length, 1, 'one image rel for the one background')
			await assertBackgroundResolves(zip, 'ppt/slides/slide1.xml', PNG_B)
		},
	},
	{
		name: 'a colour background assigned after an image background paints the colour',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = { data: PNG_A }
				s.background = { color: 'FF0000' }
			})
			const bg = /<p:bg>[\s\S]*?<\/p:bg>/.exec(await readEntry(zip, 'ppt/slides/slide1.xml'))?.[0] ?? ''
			assert(!bg.includes('a:blipFill'), `the image background is gone; got ${bg}`)
			assert(/<a:srgbClr val="FF0000"/.test(bg), `the colour background is painted; got ${bg}`)
		},
	},
])

/** Two different 1x1 PNGs, so a part that holds the wrong one is caught by its bytes. */
const PNG_A =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const PNG_B =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

/** The attributes of every `Relationship` in a `.rels` part. */
function relsOf(xml) {
	return selfClosingTags(xml, 'Relationship').map((tag) => xmlAttributes(tag))
}

/** The `ppt/slideLayouts/slideLayoutN.xml` whose `p:cSld@name` is `title`. */
async function layoutPartFor(zip, title) {
	for (const name of listEntries(zip).filter((entry) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(entry))) {
		if ((await readEntry(zip, name)).includes(`<p:cSld name="${title}"`)) return name
	}
	throw new Error(`no layout named ${title}`)
}

/**
 * The part's rel ids are unique, its background `r:embed` names an image rel, and that rel's part
 * holds `expected`'s bytes.
 */
async function assertBackgroundResolves(zip, partName, expected) {
	const relsName = partName.replace(/([^/]+)$/, '_rels/$1.rels')
	const rels = relsOf(await readEntry(zip, relsName))
	const ids = rels.map((rel) => rel.Id)
	assertEqual(new Set(ids).size, ids.length, `${relsName} rel ids are unique: ${ids.join(',')}`)

	const embed = /<p:bg>[\s\S]*?<a:blip r:embed="([^"]+)"/.exec(await readEntry(zip, partName))?.[1]
	assert(embed, `${partName} has an image background`)
	const rel = rels.find((candidate) => candidate.Id === embed)
	assert(rel?.Type.endsWith('/image'), `${partName} background ${embed} names an image rel`)

	const mediaPath = `ppt/${rel.Target.replace(/^\.\.\//, '')}`
	assert(listEntries(zip).includes(mediaPath), `${rel.Target} is in the package`)
	const bytes = await zip.file(mediaPath).async('base64')
	assertEqual(bytes, expected.split('base64,')[1], `${partName} background shows its own image`)
}
