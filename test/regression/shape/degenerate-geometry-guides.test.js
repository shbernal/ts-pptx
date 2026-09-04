import { defineRegressionSuite, build, readEntry, assert, captureDiagnostics } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// `rectRadius` is stated as a length and emitted as a fraction of the shape's SHORTER side, so
// that side is a divisor. A shape whose shorter side is zero put the division on zero and wrote
// `fmla="val Infinity"`. A guide formula is a plain string in the schema, so nothing downstream
// refuses that -- this is the only place it can be caught.
defineRegressionSuite('Degenerate preset-geometry guides', [
	{
		name: 'a text box with an explicit zero height omits the guide rather than dividing by zero',
		fn: async () => {
			// `h: 0` rather than an omitted height: both definers now default one (`addShape` to 1in,
			// `addText` to 0.3in), so a *stated* zero is what still reaches the emitter with
			// `cy === 0`. It is kept rather than rescued on purpose -- silence and an explicit zero
			// are different statements -- which is what leaves this the reachable spelling.
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('hi', {
						x: 1,
						y: 1,
						w: 2,
						h: 0,
						shape: 'roundRect',
						rectRadius: 0.1,
						line: { color: 'FF0000' },
					})
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('Infinity'), `no Infinity may reach a guide formula; got: ${xml}`)
			assert(codes.includes('shape/degenerate-extent'), `expected the warning; got ${JSON.stringify(codes)}`)
		},
	},
	{
		name: 'a non-finite rectRadius is reported and dropped',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addShape('roundRect', { x: 1, y: 1, w: 2, h: 1, rectRadius: Number.NaN })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('NaN'), `no NaN may reach a guide formula; got: ${xml}`)
			assert(!xml.includes('<a:gd '), `the guide is omitted; got: ${xml}`)
			assert(
				codes.includes('geometry/invalid-shape-adjust'),
				`expected the invalid-adjust warning; got ${JSON.stringify(codes)}`
			)
		},
	},
	{
		name: 'a non-finite arcThicknessRatio is reported and dropped, leaving the angle guides',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addShape('blockArc', {
						x: 1,
						y: 1,
						w: 2,
						h: 2,
						angleRange: [270, 0],
						arcThicknessRatio: Number.NaN,
					})
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('NaN'), `no NaN may reach a guide formula; got: ${xml}`)
			assert(xml.includes('<a:gd name="adj1"'), `the two angle guides still emit; got: ${xml}`)
			assert(!xml.includes('<a:gd name="adj3"'), `adj3 is omitted; got: ${xml}`)
			assert(
				codes.includes('geometry/invalid-shape-adjust'),
				`expected the invalid-adjust warning; got ${JSON.stringify(codes)}`
			)
		},
	},
])
