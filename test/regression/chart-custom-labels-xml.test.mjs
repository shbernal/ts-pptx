import { ChartType } from '../../dist/node.js'
import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries, assert } from '../helpers.js'

// Characterization test for per-point `customLabels` (makeCustomDLblXml) on a category-axis chart
// (bar/line/area/radar) — ZERO coverage anywhere: no demo chart sets `customLabels`, so this
// `<c:dLbl>` shape carries no baseline parts, and no existing regression test exercises it. Pins
// the byte-level detail the el()/voidEl() migration must preserve: the custom label's `<a:t>` text
// is escaped (previously a manual encodeXmlEntities call, now centralized via el()'s auto-escaping).

async function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry')
	return readEntry(zip, path)
}

describe('chart customLabels (makeCustomDLblXml)', () => {
	test('per-point customLabels text is escaped and rendered as a c:dLbl rich run', async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2], customLabels: ['Q&A', null] }], {
				type: ChartType.bar,
				x: 1,
				y: 1,
				w: 6,
				h: 4,
			})
		})
		const xml = await chartXml(zip)
		expect(xml).toContain('<c:dLbl><c:idx val="0"/>')
		expect(xml).toContain('<a:t>Q&amp;A</a:t>')
		// The second point's customLabel is null → no dLbl for idx 1.
		expect(xml).not.toContain('<c:dLbl><c:idx val="1"/>')
	})
})
