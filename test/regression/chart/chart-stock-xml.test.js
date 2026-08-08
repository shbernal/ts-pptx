import { describe, expect, test } from 'vitest'
import { build, readEntry, listEntries } from '../../helpers.js'

// Stock (high-low-close) is a CLASSIC (c:) chart, not a chartEx layout. Confirmed against the four
// stock charts PowerPoint authors (`Shapes.AddChart2(-1, {88|89|90|91}, …)`) and reads back as
// ChartType 88/89/90/91, the pieces specific to a stock chart are:
//   - a `<c:stockChart>` holding fixed-order price series drawn with INVISIBLE lines, plus a
//     `<c:hiLowLines>` element connecting each category's high/low;
//   - the open-close styles (ohlc / vohlc) additionally emit `<c:upDownBars>`; the three-value
//     styles (hlc / vhlc) instead mark the final "close" series with a dot marker;
//   - the volume styles (vhlc / vohlc) lead with a `<c:barChart>` Volume series on the PRIMARY axis
//     pair and put the price `<c:stockChart>` on a SECONDARY axis pair (4 axes total; the second
//     category axis is hidden with `<c:delete val="1"/>`).
// The generic classic-chart wiring (content type, chart rel, embedded xlsx) is covered elsewhere;
// here we pin the stock-specific structure.

const LABELS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
const S = (name, values) => ({ name, labels: LABELS, values })
const HIGH = S('High', [55, 57, 57, 58, 58])
const LOW = S('Low', [11, 12, 13, 11, 35])
const CLOSE = S('Close', [32, 35, 34, 35, 43])
const OPEN = S('Open', [20, 33, 30, 33, 37])
const VOL = S('Volume', [1200, 1500, 900, 1700, 1400])

const STYLE_DATA = {
	hlc: [HIGH, LOW, CLOSE],
	ohlc: [OPEN, HIGH, LOW, CLOSE],
	vhlc: [VOL, HIGH, LOW, CLOSE],
	vohlc: [VOL, OPEN, HIGH, LOW, CLOSE],
}

async function buildStock(stockStyle, extra = {}) {
	return build((p) => {
		p.addSlide().addChart(STYLE_DATA[stockStyle], { type: 'stock', stockStyle, x: 1, y: 1, w: 8, h: 4.5, ...extra })
	})
}

describe('stock (classic) chart', () => {
	test('emits a classic chart part (no chartEx, no AlternateContent)', async () => {
		const { zip } = await buildStock('hlc')
		const entries = listEntries(zip)
		expect(entries).toContain('ppt/charts/chart1.xml')
		expect(entries.some((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))).toBe(false)
		const contentTypes = await readEntry(zip, '[Content_Types].xml')
		expect(contentTypes).toContain(
			'<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
		)
		const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(slide).not.toContain('<mc:AlternateContent')
		expect(slide).toContain('<c:chart')
	})

	test('hlc: one stockChart with High/Low/Close, hiLowLines, no upDownBars, close marked with a dot', async () => {
		const { zip } = await buildStock('hlc')
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect(xml).not.toContain('<c:barChart>')
		expect((xml.match(/<c:stockChart>/g) || []).length).toBe(1)
		expect((xml.match(/<c:ser>/g) || []).length).toBe(3)
		expect(xml).toContain('<c:v>High</c:v>')
		expect(xml).toContain('<c:v>Low</c:v>')
		expect(xml).toContain('<c:v>Close</c:v>')
		expect(xml).toContain('<c:hiLowLines>')
		expect(xml).not.toContain('<c:upDownBars>')
		// The close series (last of a 3-value style) gets a dot marker; the others are symbol="none".
		expect(xml).toContain('<c:symbol val="dot"/>')
		expect((xml.match(/<c:symbol val="none"\/>/g) || []).length).toBe(2)
		// Two axis definitions only (category + value), both primary — so 4 <c:axId> total
		// (2 references inside <c:stockChart> + 1 in each of the 2 axis defs).
		expect((xml.match(/<c:axId /g) || []).length).toBe(4)
	})

	test('ohlc: four price series with upDownBars and no dot markers', async () => {
		const { zip } = await buildStock('ohlc')
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect(xml).not.toContain('<c:barChart>')
		expect((xml.match(/<c:ser>/g) || []).length).toBe(4)
		expect(xml).toContain('<c:v>Open</c:v>')
		expect(xml).toContain('<c:hiLowLines>')
		expect(xml).toContain('<c:upDownBars>')
		expect(xml).not.toContain('<c:symbol val="dot"/>')
	})

	test('vhlc: Volume barChart on primary axes + price stockChart on secondary axes (4 axes, hidden 2nd cat)', async () => {
		const { zip } = await buildStock('vhlc')
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		// A leading barChart (Volume) then the stockChart (High/Low/Close).
		expect((xml.match(/<c:barChart>/g) || []).length).toBe(1)
		expect((xml.match(/<c:stockChart>/g) || []).length).toBe(1)
		expect(xml).toContain('<c:v>Volume</c:v>')
		expect(xml.indexOf('<c:barChart>')).toBeLessThan(xml.indexOf('<c:stockChart>'))
		// barChart references the PRIMARY axis pair; stockChart references the SECONDARY pair.
		expect(xml).toMatch(
			/<c:barChart>[\s\S]*<c:axId val="2094734554"\/><c:axId val="2094734552"\/>[\s\S]*<\/c:barChart>/
		)
		expect(xml).toMatch(
			/<c:stockChart>[\s\S]*<c:axId val="2094734555"\/><c:axId val="2094734553"\/>[\s\S]*<\/c:stockChart>/
		)
		// Four axis definitions: catAx, valAx, valAx (secondary, right), catAx (secondary, hidden).
		expect((xml.match(/<c:catAx>/g) || []).length).toBe(2)
		expect((xml.match(/<c:valAx>/g) || []).length).toBe(2)
		// The secondary category axis is hidden.
		expect(xml).toContain('<c:delete val="1"/>')
		expect(xml).toContain('<c:hiLowLines>')
		expect(xml).not.toContain('<c:upDownBars>')
	})

	test('vohlc: Volume bar + four price series + upDownBars (5 series total)', async () => {
		const { zip } = await buildStock('vohlc')
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect((xml.match(/<c:barChart>/g) || []).length).toBe(1)
		expect((xml.match(/<c:stockChart>/g) || []).length).toBe(1)
		expect((xml.match(/<c:ser>/g) || []).length).toBe(5)
		expect(xml).toContain('<c:v>Volume</c:v>')
		expect(xml).toContain('<c:v>Open</c:v>')
		expect(xml).toContain('<c:upDownBars>')
	})

	test('stockStyle defaults to hlc when omitted or invalid', async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart([HIGH, LOW, CLOSE], { type: 'stock', x: 1, y: 1, w: 8, h: 4.5 })
		})
		const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
		expect((xml.match(/<c:stockChart>/g) || []).length).toBe(1)
		expect(xml).not.toContain('<c:barChart>')
		expect(xml).not.toContain('<c:upDownBars>')
	})
})
