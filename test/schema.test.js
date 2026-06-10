// Schema-validation fixtures. Each case builds a representative `.pptx`
// and asserts the OpenXmlValidator (via OOXMLValidatorCLI) reports no
// errors.
//
// Fixtures are intentionally small and orthogonal — they exercise one
// API surface each — so when an error appears we can localise it.
//
// Run with: pnpm run test:schema

import { build, assert } from './helpers.js'
import { validateBuf } from './validator.js'

async function expectNoSchemaErrors(buf, label) {
	const errors = await validateBuf(buf)
	if (errors.length === 0) return
	const summary = errors
		.slice(0, 5)
		.map((e) => `  - [${e.ErrorType}] ${e.Description} (path: ${(e.Path && e.Path.PartUri) || '?'})`)
		.join('\n')
	const more = errors.length > 5 ? `\n  ...(${errors.length - 5} more)` : ''
	assert(false, `${label}: ${errors.length} schema error(s):\n${summary}${more}`)
}

export default [
	{
		name: 'empty deck (one slide, no content)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide()
			})
			await expectNoSchemaErrors(buf, 'empty-deck')
		},
	},
	{
		name: 'single text box',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'single-text')
		},
	},
	{
		name: 'company metadata with XML entities',
		fn: async () => {
			const { buf } = await build((p) => {
				p.company = 'A & B <C>'
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'company-metadata-xml-entities')
		},
	},
	{
		// Asserts the body-property serialization stays schema-valid. Note: this
		// proves the XML is well-formed, not that PowerPoint/LibreOffice renders
		// a particular layout (see UPSTREAMING_CANDIDATES.md "Text Box Behavior").
		name: 'text box with margins',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1, margin: [10, 5, 10, 5] })
			})
			await expectNoSchemaErrors(buf, 'text-margins')
		},
	},
	{
		name: 'text box with vertical alignment',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('top', { x: 1, y: 1, w: 4, h: 1, valign: 'top' })
				s.addText('middle', { x: 1, y: 2, w: 4, h: 1, valign: 'middle' })
				s.addText('bottom', { x: 1, y: 3, w: 4, h: 1, valign: 'bottom' })
			})
			await expectNoSchemaErrors(buf, 'text-valign')
		},
	},
	{
		name: 'text box with mixed bold/color runs',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText(
					[
						{ text: 'bold red ', options: { bold: true, color: 'FF0000' } },
						{ text: 'plain ', options: {} },
						{ text: 'blue', options: { color: '0000FF' } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			await expectNoSchemaErrors(buf, 'text-mixed-runs')
		},
	},
	{
		// Serialization-contract fixture: breakLine: false on a CRLF-containing run must
		// produce valid OOXML (upstream-issue-1138). The rendering result is layout-dependent
		// and not asserted here — see UPSTREAMING_CANDIDATES.md "Text Box Behavior".
		name: 'rich text with breakLine: false on CRLF-containing run',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText(
					[
						{ text: 'first\nsecond', options: { breakLine: false } },
						{ text: ' tail', options: {} },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			await expectNoSchemaErrors(buf, 'rich-text-breakline-false')
		},
	},
	{
		name: 'single rectangle shape',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
			})
			await expectNoSchemaErrors(buf, 'single-shape')
		},
	},
	{
		name: 'shape with shadow',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: { color: '00B0B9' },
					shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-with-shadow')
		},
	},
	{
		name: 'shape with inner shadow',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: { color: '00B0B9' },
					shadow: { type: 'inner', blur: 6, offset: 2, color: '000000', opacity: 0.15 },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-with-inner-shadow')
		},
	},
	{
		name: 'shape with native linear gradient fill',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: {
						type: 'gradient',
						gradient: {
							kind: 'linear',
							angle: 90,
							scaled: true,
							stops: [
								{ position: 0, color: '451DC7' },
								{ position: 100, color: '0B003D', transparency: 10 },
							],
						},
					},
				})
			})
			await expectNoSchemaErrors(buf, 'shape-native-linear-gradient')
		},
	},
	{
		name: 'shape with pattern fill',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: {
						type: 'pattern',
						pattern: { preset: 'diagCross', fgColor: '003366', bgColor: 'FFFFFF' },
					},
				})
			})
			await expectNoSchemaErrors(buf, 'shape-pattern-fill')
		},
	},
	{
		name: 'solid-color slide background',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.background = { color: '0088CC' }
				s.addText('hi', { x: 1, y: 1 })
			})
			await expectNoSchemaErrors(buf, 'solid-bg')
		},
	},
	{
		name: 'native linear gradient slide background',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.background = {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 90,
						scaled: true,
						stops: [
							{ position: 0, color: '451DC7' },
							{ position: 100, color: '0B003D' },
						],
					},
				}
				s.addText('hi', { x: 1, y: 1 })
			})
			await expectNoSchemaErrors(buf, 'native-linear-gradient-bg')
		},
	},
	{
		name: 'bullet text',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('item', { x: 1, y: 1, w: 4, h: 0.5, bullet: true })
			})
			await expectNoSchemaErrors(buf, 'bullet-text')
		},
	},
	{
		name: 'simple table',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'A1' }, { text: 'B1' }],
						[{ text: 'A2' }, { text: 'B2' }],
					],
					{ x: 1, y: 1, w: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'simple-table')
		},
	},
	{
		name: 'table with hasHeader',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
					],
					{ x: 1, y: 1, w: 4, hasHeader: true }
				)
			})
			await expectNoSchemaErrors(buf, 'table-has-header')
		},
	},
	{
		name: 'table with built-in style and all style flags',
		fn: async () => {
			const { TABLE_STYLE } = await import('../dist/index.js')
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
						[{ text: 'A2' }, { text: 'B2' }],
						[{ text: 'Total' }, { text: '42' }],
					],
					{
						x: 1,
						y: 1,
						w: 4,
						tableStyle: TABLE_STYLE.MEDIUM_STYLE_2_ACCENT_1,
						hasHeader: true,
						hasFooter: true,
						hasBandedRows: true,
						hasBandedColumns: false,
						hasFirstColumn: false,
						hasLastColumn: false,
					}
				)
			})
			await expectNoSchemaErrors(buf, 'table-built-in-style-all-flags')
		},
	},
	{
		name: 'custom table style exercising every region',
		fn: async () => {
			const { buf } = await build((p) => {
				const style = p.defineTableStyle({
					name: 'Brand & <Banded>',
					wholeTbl: { border: { type: 'solid', color: 'D9D9D9', pt: 0.5 } },
					firstRow: { fill: '1A2B3C', color: 'FFFFFF', bold: true },
					lastRow: { fill: 'CCCCCC', bold: true, italic: true },
					firstCol: { color: '1A2B3C', bold: true },
					lastCol: { color: '1A2B3C' },
					band1H: { fill: 'EAF1F8' },
					band2H: { fill: 'FFFFFF' },
					band1V: { fill: 'F4F7FB' },
					band2V: { fill: 'FFFFFF' },
				})
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }, { text: 'Col C' }],
						[{ text: 'A1' }, { text: 'B1' }, { text: 'C1' }],
						[{ text: 'A2' }, { text: 'B2' }, { text: 'C2' }],
						[{ text: 'Total' }, { text: '42' }, { text: '99' }],
					],
					{
						x: 1,
						y: 1,
						w: 6,
						tableStyle: style,
						hasHeader: true,
						hasFooter: true,
						hasBandedRows: true,
						hasBandedColumns: true,
						hasFirstColumn: true,
						hasLastColumn: true,
					}
				)
			})
			await expectNoSchemaErrors(buf, 'custom-table-style-all-regions')
		},
	},
	{
		name: 'custom table style with TRBL border array',
		fn: async () => {
			const { buf } = await build((p) => {
				const style = p.defineTableStyle({
					name: 'Outline Only',
					firstRow: {
						fill: '004400',
						color: 'FFFFFF',
						border: [
							{ type: 'solid', color: '000000', pt: 2 },
							{ type: 'none' },
							{ type: 'dash', color: '888888', pt: 1 },
							{ type: 'none' },
						],
					},
				})
				p.addSlide().addTable(
					[
						[{ text: 'H1' }, { text: 'H2' }],
						[{ text: 'a' }, { text: 'b' }],
					],
					{ x: 1, y: 1, w: 4, tableStyle: style, hasHeader: true }
				)
			})
			await expectNoSchemaErrors(buf, 'custom-table-style-trbl-border')
		},
	},
	{
		name: 'embedded PNG image',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({ data: 'image/png;base64,' + b64, x: 1, y: 1, w: 1, h: 1 })
			})
			await expectNoSchemaErrors(buf, 'embedded-png')
		},
	},
	{
		name: 'image clipped to a freeform custGeom path',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [{ x: 1, y: 0 }, { x: 2, y: 2 }, { x: 0, y: 2 }, { close: true }],
				})
			})
			await expectNoSchemaErrors(buf, 'image-custgeom')
		},
	},
	{
		name: 'image with duotone recolor',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					duotone: { shadow: '250F6B', highlight: 'FFFFFF' },
				})
			})
			await expectNoSchemaErrors(buf, 'image-duotone')
		},
	},
	{
		name: 'image cover/contain sizing emits schema-valid srcRect (incl. negative contain inset)',
		fn: async () => {
			// 1x1 PNG (natural square): cover crops, contain pads with a negative srcRect inset —
			// both must stay schema-valid (CT_RelativeRect permits negative ST_Percentage).
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					sizing: { type: 'cover', w: 4, h: 3 },
				})
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 5,
					y: 0.5,
					w: 4,
					h: 3,
					sizing: { type: 'contain', w: 4, h: 3 },
				})
			})
			await expectNoSchemaErrors(buf, 'image-cover-contain')
		},
	},
	{
		name: 'text caps: all-caps and small-caps run properties',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'ALL CAPS ', options: { caps: 'all' } },
						{ text: 'Small Caps ', options: { caps: 'small' } },
						{ text: 'Normal', options: { caps: 'none' } },
					],
					{ x: 1, y: 1, w: 6, h: 0.5 }
				)
			})
			await expectNoSchemaErrors(buf, 'text-caps')
		},
	},
	{
		name: 'scatter chart with valAxisCrossesAt zero',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [0, 1, 2] },
						{ name: 'Y-Value 1', values: [1, 4, 9], labels: ['A', 'B', 'C'] },
					],
					{ x: 1, y: 1, w: 6, h: 3, valAxisCrossesAt: 0 }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-val-axis-crosses-at-zero')
		},
	},
	{
		name: 'scatter chart with independent axis format codes',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [0, 1, 2] },
						{ name: 'Y-Value 1', values: [1, 4, 9], labels: ['A', 'B', 'C'] },
					],
					{ x: 1, y: 1, w: 6, h: 3, catAxisLabelFormatCode: '0.0', valAxisLabelFormatCode: '#,##0' }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-independent-axis-format-codes')
		},
	},
	{
		name: 'bar chart with valAxisCrossBetween midCat',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					valAxisCrossBetween: 'midCat',
				})
			})
			await expectNoSchemaErrors(buf, 'bar-chart-cross-between-midcat')
		},
	},
	{
		name: 'bullet color (buClr separate from text color)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('item', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					bullet: { color: 'FF0000', characterCode: '2022' },
					color: '000000',
				})
			})
			await expectNoSchemaErrors(buf, 'bullet-color')
		},
	},
	{
		name: 'line chart with transparent marker fill',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					chartColors: ['transparent'],
				})
			})
			await expectNoSchemaErrors(buf, 'line-chart-transparent-marker')
		},
	},
	{
		name: 'line chart with null values defaults to gap',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C', 'D'], values: [1, null, 3, 4] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
				})
			})
			await expectNoSchemaErrors(buf, 'line-chart-null-values-gap')
		},
	},
	{
		name: 'line chart with per-series lineDashValues',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
						{ name: 'S2', labels: ['A', 'B', 'C'], values: [4, 3, 2] },
						{ name: 'S3', labels: ['A', 'B', 'C'], values: [2, 4, 1] },
					],
					{ x: 1, y: 1, w: 6, h: 3, lineDashValues: ['solid', 'dash', 'lgDashDot'] }
				)
			})
			await expectNoSchemaErrors(buf, 'line-chart-per-series-dash')
		},
	},
	{
		name: 'image hyperlink with query-string ampersand produces valid XML',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					hyperlink: { url: 'https://example.com/page?a=1&b=2&c=3' },
				})
			})
			await expectNoSchemaErrors(buf, 'image-hyperlink-query-string')
		},
	},
	{
		name: 'combo chart with per-subchart legend suppression',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					[
						{
							type: p.charts.BAR,
							data: [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }],
							options: {},
						},
						{
							type: p.charts.LINE,
							data: [{ name: 'Target', labels: ['Q1', 'Q2', 'Q3'], values: [15, 15, 15] }],
							options: { showLegend: false },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3, showLegend: true }
				)
			})
			await expectNoSchemaErrors(buf, 'combo-chart-subchart-legend-suppress')
		},
	},
	{
		// #1355: a scatter subchart in a combo chart needs its category (X) axis
		// emitted as a <c:valAx>, not a <c:catAx>. Emitting a catAx made
		// PowerPoint flag the file for repair. Scatter rides the secondary axes.
		name: 'combo chart with bar and scatter on secondary axes',
		fn: async () => {
			const labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					[
						{
							type: p.charts.BAR,
							data: [{ name: 'Bottom', labels, values: [17, 26, 53, 10, 4] }],
							options: { barDir: 'bar', barGrouping: 'clustered' },
						},
						{
							type: p.charts.SCATTER,
							data: [
								{ name: 'X-Axis', labels, values: [1, 2, 3, 4, 5] },
								{ name: 'Y', labels, values: [25, 35, 55, 10, 5] },
							],
							options: { secondaryValAxis: true, secondaryCatAxis: true },
						},
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showLegend: false,
						valAxes: [{ valAxisTitle: 'Primary' }, { valAxisTitle: 'Secondary' }],
						catAxes: [{ catAxisTitle: 'Primary Cat' }, { catAxisHidden: true }],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'combo-bar-scatter-secondary-axes')
		},
	},
	{
		name: 'chart with per-series color and data-label overrides',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'Alpha', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] },
						{ name: 'Beta', labels: ['Q1', 'Q2', 'Q3'], values: [15, 25, 5] },
						{ name: 'Gamma', labels: ['Q1', 'Q2', 'Q3'], values: [5, 10, 20] },
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showValue: true,
						dataLabelColor: '000000',
						dataLabelFontSize: 10,
						seriesOptions: [
							{ color: 'FF0000', dataLabelColor: 'FFFFFF', dataLabelFontBold: true },
							{ color: '00AA00', dataLabelFontSize: 14, dataLabelFontItalic: true },
							{ lineSize: 0 },
						],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'chart-series-options')
		},
	},
	{
		name: 'line chart with per-series lineSize overrides',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{ name: 'Thick', labels: ['Jan', 'Feb', 'Mar'], values: [1, 2, 3] },
						{ name: 'Normal', labels: ['Jan', 'Feb', 'Mar'], values: [3, 2, 1] },
						{ name: 'Hidden', labels: ['Jan', 'Feb', 'Mar'], values: [2, 2, 2] },
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						lineSize: 2,
						seriesOptions: [{ lineSize: 6 }, {}, { lineSize: 0 }],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'chart-series-linesize-overrides')
		},
	},
	{
		name: 'firstSlideNum sets presentation starting slide number',
		fn: async () => {
			const { buf } = await build((p) => {
				p.firstSlideNum = 5
				const slide = p.addSlide()
				slide.addText('', { x: 0, y: 0, w: 1, h: 1, slideNumber: { x: 0.5, y: 0.5 } })
			})
			await expectNoSchemaErrors(buf, 'first-slide-num')
		},
	},
	{
		name: 'round2SameRect and round2DiagRect with rectRadius emit adj1/adj2',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.ROUND_2_SAME_RECTANGLE, {
					x: 0.5,
					y: 0.5,
					w: 3,
					h: 2,
					rectRadius: 0.1,
					fill: { color: '4472C4' },
				})
				s.addShape(p.shapes.ROUND_2_DIAG_RECTANGLE, {
					x: 4,
					y: 0.5,
					w: 3,
					h: 2,
					rectRadius: 0.15,
					fill: { color: 'ED7D31' },
				})
			})
			await expectNoSchemaErrors(buf, 'round2-rect-adj1-adj2')
		},
	},
	{
		name: 'bar chart with multi-level category labels (multiLvlStrRef)',
		fn: async () => {
			const LABELS = [
				['Gear', 'Berg', 'Motr', 'Swch', 'Plug', 'Cord', 'Pump', 'Leak', 'Seal'],
				['Mech', '', '', 'Elec', '', '', 'Hydr', '', ''],
			]
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'West', labels: LABELS, values: [11, 8, 3, 0, 11, 3, 0, 0, 0] },
						{ name: 'Ctrl', labels: LABELS, values: [0, 11, 6, 19, 12, 5, 0, 0, 0] },
						{ name: 'East', labels: LABELS, values: [0, 3, 2, 0, 0, 0, 4, 3, 1] },
					],
					{ x: 1, y: 1, w: 6, h: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-multilevel-categories')
		},
	},
	{
		name: 'custom document properties (string, integer, float, boolean, date)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.setCustomProperty('Author', 'Jane Smith')
				p.setCustomProperty('Version', 3)
				p.setCustomProperty('Score', 1.5)
				p.setCustomProperty('Published', true)
				p.setCustomProperty('CreatedAt', new Date('2026-01-01T00:00:00Z'))
				p.addSlide()
			})
			await expectNoSchemaErrors(buf, 'custom-document-properties')
		},
	},
	{
		name: 'slide master roundRect object and roundRect placeholder',
		fn: async () => {
			const { buf } = await build((p) => {
				p.defineSlideMaster({
					title: 'ROUNDRECT_MASTER',
					objects: [
						{
							roundRect: {
								x: 0.5,
								y: 0.5,
								w: 2,
								h: 1,
								rectRadius: 0.1,
								fill: { color: 'E8F0FE' },
							},
						},
						{
							placeholder: {
								options: {
									name: 'title',
									type: 'title',
									x: 0.5,
									y: 2,
									w: 9,
									h: 1.5,
									shape: 'roundRect',
									rectRadius: 0.15,
								},
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterName: 'ROUNDRECT_MASTER' })
			})
			await expectNoSchemaErrors(buf, 'slide-master-roundrect')
		},
	},
	{
		name: 'bar chart with per-point customLabels',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [10, 20, 30], customLabels: ['Low', '', 'High'] },
						{ name: 'Series 2', labels: ['A', 'B', 'C'], values: [15, 5, 25], customLabels: ['', 'Min', ''] },
					],
					{ x: 1, y: 1, w: 6, h: 3, showValue: true }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-custom-labels')
		},
	},
	{
		name: 'pie chart with per-point customLabels',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.PIE,
					[
						{
							name: 'Status',
							labels: ['Red', 'Amber', 'Green'],
							values: [10, 30, 60],
							customLabels: ['At Risk', 'Watch', 'On Track'],
						},
					],
					{ x: 1, y: 1, w: 4, h: 3, showValue: true }
				)
			})
			await expectNoSchemaErrors(buf, 'pie-chart-custom-labels')
		},
	},
	{
		name: 'bar chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{
							name: 'Status',
							labels: ['A', 'B', 'C', 'D'],
							values: [10, 20, 38, 2],
							pointStyles: [
								{ border: { pt: 2, color: 'FF0000' } },
								{},
								{ fill: '00B050', border: { type: 'dash', color: '404040' } },
								{ border: { type: 'none' } },
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-point-styles')
		},
	},
	{
		name: 'pie chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.PIE,
					[
						{
							name: 'Status',
							labels: ['Red', 'Amber', 'Green'],
							values: [10, 30, 60],
							pointStyles: [{ border: { pt: 3, color: 'C00000' } }, {}, { fill: '70AD47' }],
						},
					],
					{ x: 1, y: 1, w: 4, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'pie-chart-point-styles')
		},
	},
	{
		name: 'line chart with per-point pointStyles (border)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C', 'D'],
							values: [4, 8, 6, 10],
							pointStyles: [
								{},
								{ border: { pt: 2, color: 'FF0000' } },
								{},
								{ border: { type: 'dash', color: '0070C0' } },
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'line-chart-point-styles')
		},
	},
	{
		name: 'area chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.AREA,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C'],
							values: [5, 9, 7],
							pointStyles: [{ fill: 'FFC000' }, {}, { border: { pt: 1, color: '404040' } }],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'area-chart-point-styles')
		},
	},
	{
		name: 'scatter chart with per-point pointStyles (border)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{
							name: 'Y-Values',
							values: [3, 6, 2, 8],
							pointStyles: [{ border: { pt: 2, color: 'FF0000' } }, {}, { fill: '00B050' }, {}],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-chart-point-styles')
		},
	},
]
