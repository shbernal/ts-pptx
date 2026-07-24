import { ChartType } from '../../dist/node.js'
import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries, assert } from '../helpers.js'

// Characterization test for the chart `metadata` extLst that the byte-identity harness CANNOT see —
// the demo deck never sets chart `metadata`, so `<c:extLst>`/`<pgm:item>` carry ZERO baseline parts.
// Pins the byte-level detail the el()/voidEl() migration must preserve: key/value escaping on
// `<pgm:item>` (previously manual encodeXmlEntities calls, now centralized via voidEl()'s attrs).

async function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry')
	return readEntry(zip, path)
}

const DATA = [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }]

describe('chart metadata extLst', () => {
	test('metadata keys/values are escaped and wrapped in the pgm:metadata extension', async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart(DATA, {
				type: ChartType.bar,
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				metadata: { 'q&a': 'a<b>c' },
			})
		})
		const xml = await chartXml(zip)
		expect(xml).toContain(
			'<c:extLst><c:ext uri="{094A432E-1F6C-499B-95B8-B57DC9536949}"><pgm:metadata xmlns:pgm="http://ts-pptx.com/schema/chart/metadata"><pgm:item key="q&amp;a" value="a&lt;b&gt;c"/></pgm:metadata></c:ext></c:extLst>'
		)
	})

	test('no metadata option emits no extLst', async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart(DATA, { type: ChartType.bar, x: 1, y: 1, w: 6, h: 4 })
		})
		const xml = await chartXml(zip)
		expect(xml).not.toContain('pgm:metadata')
		expect(xml).not.toContain('<c:extLst>')
	})
})
