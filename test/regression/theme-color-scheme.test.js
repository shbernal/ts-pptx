import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: ThemeProps.colorScheme must let callers override the theme1.xml <a:clrScheme>
// slots; unset slots keep the Office defaults, dk1/lt1 overrides switch from <a:sysClr> to
// <a:srgbClr>, and invalid hex warns + keeps the default (upstream gitbrent/PptxGenJS#1243).

defineRegressionSuite('Theme color scheme overrides (upstream #1243)', [
	{
		name: 'overridden slots emit srgbClr; unset slots keep Office defaults',
		fn: async () => {
			const { zip } = await build((p) => {
				p.theme = { colorScheme: { accent1: 'C00000', dk2: '1F3864', hlink: '#0070C0' } }
				p.addSlide().addText('x', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			assert(xml.includes('<a:accent1><a:srgbClr val="C00000"/></a:accent1>'), 'accent1 override missing')
			assert(xml.includes('<a:dk2><a:srgbClr val="1F3864"/></a:dk2>'), 'dk2 override missing')
			// Leading '#' is stripped.
			assert(xml.includes('<a:hlink><a:srgbClr val="0070C0"/></a:hlink>'), 'hlink override (with #) missing')
			// Unset slots retain Office defaults.
			assert(xml.includes('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'), 'accent2 default changed')
			assert(xml.includes('<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'), 'folHlink default changed')
			// dk1/lt1 untouched → still sysClr.
			assert(xml.includes('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'), 'dk1 default changed')
			// Exactly one clrScheme, all 12 slots present.
			assert((xml.match(/<a:clrScheme /g) || []).length === 1, 'expected exactly one clrScheme')
		},
	},
	{
		name: 'dk1/lt1 overrides switch from sysClr to srgbClr',
		fn: async () => {
			const { zip } = await build((p) => {
				p.theme = { colorScheme: { dk1: '101010', lt1: 'FAFAFA' } }
				p.addSlide()
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			assert(xml.includes('<a:dk1><a:srgbClr val="101010"/></a:dk1>'), 'dk1 should be srgbClr when overridden')
			assert(xml.includes('<a:lt1><a:srgbClr val="FAFAFA"/></a:lt1>'), 'lt1 should be srgbClr when overridden')
			assert(!xml.includes('<a:sysClr'), 'sysClr must be gone once dk1/lt1 are overridden')
		},
	},
	{
		name: 'invalid hex keeps the Office default (no degenerate color)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.theme = { colorScheme: { accent1: 'nope', accent2: 'FFF' } }
				p.addSlide()
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			assert(
				xml.includes('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'),
				'invalid accent1 must fall back to default'
			)
			assert(
				xml.includes('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'),
				'3-digit accent2 must fall back to default'
			)
		},
	},
	{
		name: 'no theme → unchanged default Office scheme',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide()
			})
			const xml = await readEntry(zip, 'ppt/theme/theme1.xml')
			assert(xml.includes('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'), 'default dk1 missing')
			assert(xml.includes('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'), 'default accent1 missing')
		},
	},
])
