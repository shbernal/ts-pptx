import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// Regression for upstream #25: TextPropsOptions did not declare angleRange,
// arcThicknessRatio, points or shapeAdjust — but genXmlPresetGeom reads all four on a
// styled text frame and honors them. The runtime worked while TypeScript rejected the
// identical object literal (TS2353), so correct code failed to compile.
//
// The JSDoc annotation below routes this literal through the published declarations
// under `typecheck:test` (checkJs keeps object-literal shape checking), so a key that
// disappears from TextPropsOptions fails CI even though plain-JS tests would pass.

/** @type {import('../../../dist/node.js').TextPropsOptions} */
const geometryTextOptions = {
	shape: 'blockArc',
	angleRange: [350, 10],
	arcThicknessRatio: 0.25,
}

defineRegressionSuite('Text-frame geometry declarations [upstream-25]', [
	{
		name: 'a blockArc text frame emits its preset-geometry adjustment guides',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('arc', { x: 1, y: 1, w: 3, h: 2, ...geometryTextOptions })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(xml.includes('<a:gd name="adj1" fmla="val 21000000"/>'), 'expected adj1=21000000')
			assert(xml.includes('<a:gd name="adj2" fmla="val 600000"/>'), 'expected adj2=600000')
			assert(xml.includes('<a:gd name="adj3" fmla="val 12500"/>'), 'expected adj3=12500')
		},
	},
])
