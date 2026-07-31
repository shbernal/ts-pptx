import { ChartType } from '../../dist/node.js'
import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'

// A key that is not part of `BorderProps` used to be discarded in total silence, so a border
// authored with the wrong name for its thickness rendered at the 1pt default and nothing said so.
//
// TypeScript already rejects the stray key when the border is written inline at the call site. It
// deliberately does not when the border is built as a variable first -- a variable may legitimately
// be a supertype -- and that is exactly the shape a shared grid style encourages, so the mistake
// reaches generation typed as valid. A `.pptx` offers no second signal: nothing throws and the deck
// opens fine, only a little heavier than asked for.
//
// The check lives in `resolveBorderWidth`, the one function every emitted border resolves its width
// through, so the coverage claim under test is that it fires well beyond tables. The reach is not
// universal: a caller that rebuilds a border from a fixed key list before generation strips the
// unknown key first, and the third case documents the one place that still happens.
//
// Assertions go through the emitted part, not the options object: what a caller can observe is the
// attribute that reaches the XML. `w` is in hundredths of a point -- 0.5pt = 6350, 1pt = 12700.

/** Pull every `w="…"` on an `<a:ln>` in document order. */
function lineWidths(xml) {
	return [...xml.matchAll(/<a:ln[^>]*\sw="(\d+)"/g)].map((m) => m[1])
}

const ROWS = [[{ text: 'A' }, { text: 'B' }]]
const AT = { x: 1, y: 1, w: 9 }

defineRegressionSuite('Border unknown key', [
	{
		name: 'a table border carrying an unknown key is reported and the value is dropped',
		fn: async () => {
			// `pt` is not a `BorderProps` key -- the thickness field is `width`. Assigning to a
			// variable first is what slips it past the excess-property check.
			const border = { type: 'solid', color: 'FF0000', pt: 0.5 }
			const { result, codes, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(ROWS, { ...AT, border: /** @type {any} */ (border) })
				})
			)

			assert(
				codes.includes('border/unknown-key'),
				'expected the border/unknown-key code; got: ' + JSON.stringify(codes)
			)

			// The code is the contract, but the offending key is the whole point of the report, so
			// pin that it travels -- in `detail`, which is structured and safe to assert on.
			const diagnostic = diagnostics.find((d) => d.code === 'border/unknown-key')
			assertEqual(diagnostic.detail.received, 'pt', 'the diagnostic names the offending key')

			// And confirm the premise: the authored 0.5 really was discarded for the 1pt default.
			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const widths = lineWidths(xml)
			assert(widths.length > 0, 'expected at least one <a:ln> on the table')
			assert(
				widths.every((w) => w === '12700'),
				'the unknown key must be ignored, leaving the 1pt default; got: ' + JSON.stringify(widths)
			)
		},
	},
	{
		name: 'the same border spelled `width` is silent and reaches the XML',
		fn: async () => {
			const border = { type: 'solid', color: 'FF0000', width: 0.5 }
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(ROWS, { ...AT, border })
				})
			)

			assert(!codes.includes('border/unknown-key'), 'a well-formed border must not warn; got: ' + JSON.stringify(codes))

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const widths = lineWidths(xml)
			assert(widths.length > 0, 'expected at least one <a:ln> on the table')
			assert(
				widths.every((w) => w === '6350'),
				'0.5pt must reach the XML as w="6350"; got: ' + JSON.stringify(widths)
			)
		},
	},
	{
		name: 'a chart border is vetted by the same seam',
		fn: async () => {
			// A distinct key from the table case on purpose: the report is emitted through `warnOnce`,
			// which dedupes on code AND message, so reusing `pt` here would be suppressed by the first
			// case and the assertion would pass without the chart path ever being exercised.
			//
			// `plotArea.border` and not `chartArea.border`: the two normalize differently, and only
			// one of them can reach this seam. `plotArea.border` is defaulted in place, so whatever
			// the caller wrote survives to `resolveBorderWidth`. `chartArea.border` is REBUILT from a
			// fixed key list (`gen/define/chart.ts`), which discards an unknown key -- and `cap` with
			// it -- before generation sees it, so no border check downstream of that rebuild can fire
			// for it. That is the same fixed-key-list trap `gen/define/table.ts` already replaced with
			// a spread; the chart side has not had the same treatment.
			const plotArea = { border: { color: 'FF0000', thickness: 3 } }
			const { codes, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart([{ name: 'S1', labels: ['A'], values: [1] }], {
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						type: ChartType.bar,
						plotArea: /** @type {any} */ (plotArea),
					})
				})
			)

			assert(
				codes.includes('border/unknown-key'),
				'expected the border/unknown-key code from the chart path; got: ' + JSON.stringify(codes)
			)
			const diagnostic = diagnostics.find((d) => d.code === 'border/unknown-key')
			assertEqual(diagnostic.detail.received, 'thickness', 'the diagnostic names the offending key')
		},
	},
])
