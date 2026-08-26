import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// Regression for upstream #24: the preset-geometry gates were truthiness checks, so
// legitimate zero values were dropped. rectRadius: 0 (a sharp-corner rectangle) lost
// its `adj` guide entirely and PowerPoint fell back to the preset's default rounding;
// arcThicknessRatio: 0 lost its adj3 while angleRange: [0, 0] still emitted adj1/adj2
// (an array is truthy even when its contents are zero). Zero is a deliberate value,
// not an unset option.
defineRegressionSuite('Zero-valued geometry guides [upstream-24]', [
	{
		name: 'roundRect rectRadius: 0 emits adj=0 (sharp corners)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('roundRect', { x: 1, y: 1, w: 2, h: 1, rectRadius: 0 })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(xml.includes('<a:gd name="adj" fmla="val 0"/>'), `expected adj=0; got: ${xml.slice(0, 900)}`)
		},
	},
	{
		name: 'blockArc angleRange: [0,0] with arcThicknessRatio: 0 emits all three guides',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('blockArc', {
					x: 4,
					y: 1,
					w: 2,
					h: 1,
					angleRange: [0, 0],
					arcThicknessRatio: 0,
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(xml.includes('<a:gd name="adj1" fmla="val 0"/>'), 'expected adj1=0')
			assert(xml.includes('<a:gd name="adj2" fmla="val 0"/>'), 'expected adj2=0')
			assert(xml.includes('<a:gd name="adj3" fmla="val 0"/>'), 'expected adj3=0')
		},
	},
	{
		name: 'nonzero values are byte-identical to the previous behavior',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('roundRect', { x: 1, y: 1, w: 2, h: 1, rectRadius: 0.5 })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(xml.includes('<a:gd name="adj" fmla="val 50000"/>'), `expected adj=50000; got: ${xml.slice(0, 900)}`)
		},
	},
])
