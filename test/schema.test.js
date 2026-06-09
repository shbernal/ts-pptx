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
]
