import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// A master with a single title placeholder; a slide created against it should
// inherit the placeholder (rendered as a <p:sp> with <p:ph type="title" .../>).
function defineMaster(p, title) {
	p.defineSlideMaster({
		title,
		objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 }, text: '' } }],
	})
}

defineRegressionSuite('addSlide masterTitle', [
	{
		name: 'canonical `masterTitle` applies the named master',
		fn: async () => {
			const { zip } = await build((p) => {
				defineMaster(p, 'MT_MASTER')
				p.addSlide({ masterTitle: 'MT_MASTER' }).addText('Title', { placeholder: 'title' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="title"/.test(xml), 'expected title placeholder from masterTitle; got: ' + xml)
		},
	},
	{
		// fork-slidemaster-title-unescaped: a title containing XML metacharacters used to reach
		// `<p:cSld name="...">` verbatim, producing a non-well-formed slideLayout part. It must be
		// escaped in the emitted XML while still matching `masterTitle` by its raw (unescaped) value.
		name: 'a title with XML metacharacters is escaped in the slideLayout and still matches masterTitle raw',
		fn: async () => {
			const { zip } = await build((p) => {
				defineMaster(p, 'R&D <Team> "Q1"')
				p.addSlide({ masterTitle: 'R&D <Team> "Q1"' }).addText('Title', { placeholder: 'title' })
			})
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="title"/.test(slideXml), 'expected masterTitle lookup to still match; got: ' + slideXml)

			const layoutXml = await readEntry(zip, 'ppt/slideLayouts/slideLayout2.xml')
			assert(
				layoutXml.includes('<p:cSld name="R&amp;D &lt;Team&gt; &quot;Q1&quot;">'),
				'expected escaped title in slideLayout <p:cSld name="...">; got: ' + layoutXml
			)
		},
	},
])
