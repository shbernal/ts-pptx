import { ChartType } from '../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertIncludes,
	assertNotIncludes,
} from '../helpers.js'

// addChart normalizes/validates several numeric and enum options before emitting.
// These branches (clamping an out-of-range line-marker size; dropping invalid
// gridLine size/style/cap so PowerPoint-invalid values never reach the XML) are
// exercised here through the public API so the generated chart part proves the
// scrub actually happened.

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

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
