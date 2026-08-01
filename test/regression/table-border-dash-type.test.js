import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'

// `BorderProps.dashType` -> `a:prstDash/@val` on a table cell border and on a custom
// table-style region.
//
// `BorderProps.type` is a three-way switch, so before this every dashed border -- dotted,
// long-dash, dash-dot -- came out as the single `sysDash` preset. `dashType` names the
// `ST_PresetLineDashVal` preset directly and wins over `type` when both are set; the one
// exception is `type: 'none'`, which suppresses the border before any dash is chosen.
//
// A value outside the enum would make the slide part schema-invalid, which PowerPoint
// reports as a corrupt file rather than a mis-set option, so it is checked before emission.

/** Every `ST_PresetLineDashVal` value, which is exactly what `dashType` accepts. */
const ALL_DASHES = [
	'solid',
	'dot',
	'dash',
	'lgDash',
	'dashDot',
	'lgDashDot',
	'lgDashDotDot',
	'sysDash',
	'sysDot',
	'sysDashDot',
	'sysDashDotDot',
]

/** The `a:prstDash/@val` of every border in the part, in document order. */
function dashValues(xml) {
	return [...xml.matchAll(/<a:prstDash val="([^"]*)"\/>/g)].map((m) => m[1])
}

const AT = { x: 1, y: 1, w: 9 }

defineRegressionSuite('Table border dashType', [
	{
		name: 'every ST_PresetLineDashVal value reaches a:prstDash unchanged',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					// One single-sided cell per dash: `border` as a bare object broadcasts to all
					// four sides, so each cell contributes four identical prstDash values.
					const row = ALL_DASHES.map((dash) => ({ text: dash, options: { border: { type: 'solid', dashType: dash } } }))
					p.addSlide().addTable([row], AT)
				})
			)

			assert(!codes.includes('border/invalid-dash-type'), 'valid values must not warn; got: ' + JSON.stringify(codes))
			assert(
				!codes.includes('border/unknown-key'),
				'`dashType` is a known BorderProps key; got: ' + JSON.stringify(codes)
			)

			const values = dashValues(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(values.length, ALL_DASHES.length * 4, 'four sides per cell')
			for (const [idx, dash] of ALL_DASHES.entries()) {
				const forCell = values.slice(idx * 4, idx * 4 + 4)
				assertEqual(forCell.join(','), [dash, dash, dash, dash].join(','), `cell ${idx} keeps ${dash}`)
			}
		},
	},
	{
		name: 'dashType wins over type, and type alone still collapses onto sysDash',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								// `type:'solid'` says nothing about the dash pattern; `dashType` does.
								{ text: 'A', options: { border: { type: 'solid', dashType: 'sysDot' } } },
								// No `dashType`: the coarse switch is all there is, so it is sysDash.
								{ text: 'B', options: { border: { type: 'dash' } } },
							],
						],
						AT
					)
				})
			)

			const values = dashValues(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(values.slice(0, 4).join(','), 'sysDot,sysDot,sysDot,sysDot', 'dashType overrides type')
			assertEqual(values.slice(4, 8).join(','), 'sysDash,sysDash,sysDash,sysDash', 'type alone still means sysDash')
		},
	},
	{
		name: "type:'none' beats dashType — a suppressed edge draws nothing at all",
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'A', options: { border: { type: 'none', dashType: 'lgDash' } } }]], AT)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assertEqual(dashValues(xml).length, 0, 'a suppressed border emits no prstDash')
			assert(xml.includes('<a:noFill/>'), 'it emits a:noFill instead')
		},
	},
	{
		name: 'an unrecognized dashType is reported and falls back to what type implies',
		fn: async () => {
			const { result, codes, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[[{ text: 'A', options: { border: { type: 'dash', dashType: /** @type {any} */ ('dotted') } } }]],
						AT
					)
				})
			)

			assert(
				codes.includes('border/invalid-dash-type'),
				'expected the border/invalid-dash-type code; got: ' + JSON.stringify(codes)
			)
			const diagnostic = diagnostics.find((d) => d.code === 'border/invalid-dash-type')
			assertEqual(diagnostic.detail.received, 'dotted', 'the diagnostic names the offending value')

			const values = dashValues(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(values.join(','), 'sysDash,sysDash,sysDash,sysDash', "falls back to what type:'dash' implies")
		},
	},
])
