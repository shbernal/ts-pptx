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

defineRegressionSuite('Slide backgrounds', 'legacy bug-12', [
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
		//
		// This leaves two branches in `gen/define/background.ts` deliberately red — the `_relsMedia
		// || []` and `_name || ''` fallbacks. Both callers (`slide.background =` and
		// `defineSlideMaster({ background })`) construct their target with `_relsMedia: []` and a
		// non-empty `_name` (`Slide N`, or the title `defineSlideMaster` refuses to go without), so
		// neither fallback is reachable from the public surface.
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
])
