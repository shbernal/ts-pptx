import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertIncludes, assertNotIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// Scatter data labels (`showLabel` + `dataLabelFormatScatter`) drive a large,
// otherwise-uncovered block of gen-charts. The three formats emit structurally
// different `<c:dLbls>`:
//   - 'custom'   → per-point `<c:dLbl>` rich text only (no XVALUE/YVALUE fields)
//   - 'customXY' → per-point `<c:dLbl>` rich text PLUS `type="XVALUE"`/`"YVALUE"` fields,
//                  appended only for non-blank labels (blank/space labels are skipped)
//   - 'XY'       → a single chart-level `<c:dLbls>` with showVal/showCatName toggled on
// These are locked here as regression contracts so the scatter-label emission
// cannot silently change shape.

const XY_SERIES = [
	{ name: 'X-Axis', values: [1, 2, 3] },
	{ name: 'Y1', values: [4, 5, 6], labels: [['Alpha', 'Beta', 'Gamma']] },
]

defineRegressionSuite('Chart scatter data labels', [
	{
		name: "dataLabelFormatScatter 'custom': per-point rich <c:dLbl> without XVALUE/YVALUE fields",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(XY_SERIES, {
					type: ChartType.scatter,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showLabel: true,
					dataLabelFormatScatter: 'custom',
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:dLbl>', 'custom emits per-point dLbl')
			// One <c:dLbl> per label in the (single) label group.
			assert((xml.match(/<c:dLbl>/g) || []).length === 3, 'one dLbl per label')
			assertIncludes(xml, '<a:t>Alpha</a:t>', 'custom label text is rendered as rich run')
			assertNotIncludes(xml, 'type="XVALUE"', "custom must NOT append the XVALUE field (that's customXY)")
			assertNotIncludes(xml, 'type="YVALUE"', 'custom must NOT append the YVALUE field')
		},
	},
	{
		name: "dataLabelFormatScatter 'customXY': appends XVALUE/YVALUE fields, skipping blank labels",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X-Axis', values: [1, 2, 3] },
						// Middle label is blank (a single space) → its XY suffix must be skipped.
						{ name: 'Y1', values: [4, 5, 6], labels: [['Alpha', ' ', 'Gamma']] },
					],
					{
						type: ChartType.scatter,
						x: 1,
						y: 1,
						w: 6,
						h: 4,
						showLabel: true,
						dataLabelFormatScatter: 'customXY',
						dataLabelPosition: 't',
					}
				)
			})
			const xml = await chartXml(zip)
			// Two non-blank labels → two XVALUE and two YVALUE fields (blank label emits no XY suffix).
			assert(
				(xml.match(/type="XVALUE"/g) || []).length === 2,
				'two XVALUE fields (blank label skipped); got: ' + (xml.match(/type="XVALUE"/g) || []).length
			)
			assert((xml.match(/type="YVALUE"/g) || []).length === 2, 'two YVALUE fields (blank label skipped)')
			assertIncludes(xml, '<c:dLblPos val="t"/>', 'dataLabelPosition flows into dLblPos')
			// XVALUE/YVALUE field cache text is the series name bracket-wrapped; pre-existing
			// asymmetry (XVALUE has no closing bracket) is intentionally preserved, not "fixed".
			assertIncludes(xml, '<a:t>[Y1</a:t>', 'XVALUE field cache text is [seriesName (no closing bracket)')
			assertIncludes(xml, '<a:t>[Y1]</a:t>', 'YVALUE field cache text is [seriesName]')
		},
	},
	{
		name: "dataLabelFormatScatter 'XY': single chart-level dLbls with showVal/showCatName on",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X-Axis', values: [1, 2, 3] },
						{ name: 'Y1', values: [4, 5, 6] },
					],
					{
						type: ChartType.scatter,
						x: 1,
						y: 1,
						w: 6,
						h: 4,
						showLabel: true,
						dataLabelFormatScatter: 'XY',
					}
				)
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:dLbls>', 'XY emits a chart-level dLbls block')
			assertNotIncludes(xml, '<c:dLbl>', 'XY does not emit per-point dLbl entries')
			assertIncludes(xml, '<c:showVal val="1"/>', 'XY turns showVal on')
			assertIncludes(xml, '<c:showCatName val="1"/>', 'XY turns showCatName on')
		},
	},
])
