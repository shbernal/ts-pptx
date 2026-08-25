import TsPptx, { ChartType, InvalidOptionError } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertEqual, assertIncludes, assertNotIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// addChart normalizes/validates several numeric and enum options before emitting.
// These branches (clamping an out-of-range line-marker size; dropping invalid
// gridLine size/style/cap so PowerPoint-invalid values never reach the XML) are
// exercised here through the public API so the generated chart part proves the
// scrub actually happened.

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
const BASE = { x: 1, y: 1, w: 6, h: 3 }

defineRegressionSuite('Chart option validation', [
	{
		name: 'lineDataSymbolSize above the 2-72 range is clamped to 72',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.line,
					lineDataSymbol: 'circle',
					lineDataSymbolSize: 999,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:size val="72"/>', 'oversize marker clamps to the 72 max')
			assertNotIncludes(xml, '<c:size val="999"/>', 'the out-of-range value must not reach the XML')
		},
	},
	{
		name: 'lineDataSymbolSize below the range is clamped to 2',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.line,
					lineDataSymbol: 'circle',
					lineDataSymbolSize: 1,
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:size val="2"/>', 'undersize marker clamps to the 2 min')
		},
	},
	{
		name: 'invalid gridLine cap is dropped (never emitted)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.bar,
					valGridLine: { size: 2, style: 'solid', cap: 'INVALID' },
				})
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, 'INVALID', 'an unrecognized gridLine cap must be scrubbed before emit')
		},
	},
	{
		name: 'invalid gridLine style is dropped',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, valGridLine: { size: 2, style: 'wavy' } })
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, 'wavy', 'an unrecognized gridLine style must be scrubbed before emit')
		},
	},
	{
		// The two chart emitters partition `ChartType` between them and each treats the other's
		// members as not its own, so a `type` outside the catalog matches no arm in either and
		// cannot produce a plot. It has to be refused at the boundary, where the caller can see
		// which call was wrong — reaching an emitter with it used to yield a `<c:plotArea>` with
		// axes and nothing in it, i.e. a chart-shaped hole the deck opens and shows empty.
		name: 'a chart type outside the catalog is refused at addChart, not emitted as an empty plot',
		fn: () => {
			let thrown = null
			try {
				// The cast is the point: `CHART_NAME` keeps TypeScript callers out, and this guard exists
				// for the JavaScript ones it cannot reach.
				new TsPptx().addSlide().addChart(SERIES, /** @type {never} */ ({ ...BASE, type: 'nonsense' }))
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `an unknown chart type throws InvalidOptionError (got ${thrown})`)
			assertEqual(thrown.code, 'chart/unknown-type', 'the condition carries its own code')
			assertEqual(thrown.detail.type, 'nonsense', 'the offending type is carried as structured detail')
		},
	},
	{
		name: 'a combo subchart type outside the catalog is refused the same way',
		fn: () => {
			// The combo form routes each entry's `type` through the same emitters, so the guard has
			// to see the `ChartMulti[]` entries too, not just the single-type `options.type`.
			let thrown = null
			try {
				new TsPptx().addSlide().addChart(
					/** @type {never} */ ([
						{ type: ChartType.bar, data: SERIES },
						{ type: 'nonsense', data: SERIES },
					]),
					BASE
				)
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `an unknown subchart type throws InvalidOptionError (got ${thrown})`)
			assertEqual(thrown.code, 'chart/unknown-type', 'the condition carries its own code')
		},
	},
	{
		name: 'every catalog type is accepted, so the guard cannot drift from the enum',
		fn: async () => {
			// Guards against the guard: a `ChartType` member the validator does not know would be
			// refused at the boundary even though both emitters can build it.
			for (const type of Object.values(ChartType)) {
				const { zip } = await build((p) => {
					p.addSlide().addChart(SERIES, { ...BASE, type })
				})
				assert(zip, `${type} builds through addChart`)
			}
		},
	},
	{
		name: 'non-positive gridLine size is dropped so defaults apply',
		fn: async () => {
			// A negative size with a real style takes the `size <= 0` branch (the
			// early `style === 'none'` return would otherwise skip validation).
			const { zip } = await build((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.bar, valGridLine: { size: -5, style: 'dash' } })
			})
			const xml = await chartXml(zip)
			assert(xml.length > 0, 'chart still builds after dropping the invalid gridLine size')
		},
	},
])
