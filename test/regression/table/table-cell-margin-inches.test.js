import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// Acceptance: table cell `margin` is INCHES (matching x/y/w/h and the PowerPoint dialog).
// The pre-v3.8.0 magnitude heuristic — a component `>= 1` read as POINTS — is gone; every
// value is now inches (`inch2Emu`, 914400 EMU/in). A `>= 1` value is still honored as inches
// (it just warns once as a likely legacy points value).

defineRegressionSuite('Table cell margin units', [
	{
		name: 'fractional cell margin is inches (0.5in => 457200 EMU on all four sides)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([[{ text: 'a', options: { margin: 0.5 } }]], { x: 1, y: 1, w: 4, colW: [4] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				xml.includes('marL="457200" marR="457200" marT="457200" marB="457200"'),
				'expected 0.5in cell margin to emit 457200 EMU on all sides; got: ' + xml
			)
		},
	},
	{
		name: 'a `>= 1` cell margin is honored as inches, not reinterpreted as points',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// margin is [T, R, B, L]; 2in top/bottom, 1in right/left.
				s.addTable([[{ text: 'b', options: { margin: [2, 1, 2, 1] } }]], { x: 1, y: 1, w: 4, colW: [4] })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// Old points path would have emitted ptsToEmuLenient(2)=25400 / ptsToEmuLenient(1)=12700.
			assert(
				xml.includes('marL="914400" marR="914400" marT="1828800" marB="1828800"'),
				'expected [2,1,2,1]in cell margin to emit inch-based EMU (not the legacy points values); got: ' + xml
			)
		},
	},
])
