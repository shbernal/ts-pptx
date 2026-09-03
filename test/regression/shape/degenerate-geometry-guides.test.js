import { defineRegressionSuite, build, readEntry, assert, captureDiagnostics } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// `rectRadius` is stated as a length and emitted as a fraction of the shape's SHORTER side, so
// that side is a divisor. A text object with no stated height reaches the emitter with `cy === 0`;
// `objects/text.ts` rescues a height for a LINE-LESS one, so any line at all puts the division
// back on zero and wrote `fmla="val Infinity"`. A guide formula is a plain string in the schema,
// so nothing downstream refuses that -- this is the only place it can be caught.
defineRegressionSuite('Degenerate preset-geometry guides', [
	{
		name: 'a text box with a line and no height also omits the guide rather than dividing by zero',
		fn: async () => {
			// `addShapeDefinition` stamps `line: { type: 'none' }` on every shape it defines, so the
			// rescue is unreachable from `addShape` as well -- but `addShape` defaults a height, and
			// `addText` does not, which is what leaves this the reachable spelling.
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('hi', {
						x: 1,
						y: 1,
						w: 2,
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
