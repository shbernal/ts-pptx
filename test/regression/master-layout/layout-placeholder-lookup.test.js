import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// A slide object that names a layout placeholder takes options from it at definition time and its
// frame and `<p:ph>` from it at write time. The two lookups disagreed about what counts as a layout
// placeholder: definition matched only placeholder objects, while the write-time lookup matched any
// layout object carrying a `placeholder` option, a master's plain text box included. So a slide
// object could take its frame and `<p:ph>` from an object its options never came from. With no
// placeholder of that name on the layout, the object still records the type it named and writes a
// bare `<p:ph type>` of its own, so the frame is what tells the two lookups apart.

const EMU_PER_INCH = 914400

defineRegressionSuite('Layout placeholder lookup', [
	{
		name: 'a layout text box carrying a placeholder option is not treated as that placeholder',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'DECOY',
					objects: [{ text: { text: 'not a placeholder', options: { placeholder: 'body', x: 5, y: 4, w: 2, h: 1 } } }],
				})
				p.addSlide({ masterTitle: 'DECOY' }).addText('hello', { placeholder: 'body' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shape = xml.slice(xml.indexOf('<p:sp>'))
			const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(shape)
			assert(off, `expected an <a:off> on the slide's text box; got: ${shape.slice(0, 600)}`)
			assert(
				Number(off[1]) !== 5 * EMU_PER_INCH || Number(off[2]) !== 4 * EMU_PER_INCH,
				`the slide's text box took the layout text box's frame; got: ${off[0]}`
			)
		},
	},
])
