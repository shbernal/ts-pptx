// Probe; produces no committed fixture. Builds the sensitivity trio probe-scatter-x-cache-format.ps1
// renders: which number format paints a scatter's X axis labels when the X cache and the X axis
// disagree. The axis is written source-linked, `<c:numFmt formatCode="0.00" sourceLinked="1"/>`.
//   as-written      the scatter as the writer emits it with valLabelFormatCode '0%' and
//                   catAxisLabelFormatCode '0.00': the X cache says 0%, the axis 0.00
//   cache-patched   the same bytes with only the X cache's formatCode set to 0.00
//   axis-unlinked   the same bytes with only the X axis's sourceLinked set to 0
//
// Measured 2026-09-14 before the fix: as-written paints 0% 20% 40% …, cache-patched and
// axis-unlinked paint 0.00 0.20 0.40 …. A source-linked axis paints in its cache's format, and
// cache-patched and axis-unlinked differing from as-written in opposite halves of the pair is what
// shows the render reads both. So the X cache has to carry the X axis format.
import fs from 'node:fs'
import JSZip from 'jszip'
import TsPptx, { ChartType } from '../../../../dist/node.js'

const outDir = new URL('../../../../.tmp/scatter-x-cache-format/', import.meta.url)
fs.mkdirSync(outDir, { recursive: true })

const pres = new TsPptx()
pres.addSlide().addChart(
	[
		{ name: 'X', values: [0.25, 0.5, 0.75, 1] },
		{ name: 'Y', values: [1, 2, 3, 4] },
	],
	{
		type: ChartType.scatter,
		x: 0.5,
		y: 0.5,
		w: 9,
		h: 4.5,
		valLabelFormatCode: '0%',
		catAxisLabelFormatCode: '0.00',
		showLegend: false,
	}
)
const base = Buffer.from(await pres.toBytes())
const chartPart = Object.keys((await JSZip.loadAsync(base)).files).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))

async function variant(name, edit) {
	const zip = await JSZip.loadAsync(base)
	const before = await zip.file(chartPart).async('string')
	const after = edit(before)
	if (edit && after === before) throw new Error(`${name}: the patch changed nothing`)
	zip.file(chartPart, after)
	fs.writeFileSync(new URL(`${name}.pptx`, outDir), await zip.generateAsync({ type: 'nodebuffer' }))
	console.log('wrote', name, 'X cache', /<c:xVal>[\s\S]*?<c:formatCode>([^<]*)</.exec(after)?.[1])
}
await variant('as-written', (xml) => xml)
await variant('cache-patched', (xml) =>
	xml.replace(/(<c:xVal>[\s\S]*?<c:formatCode>)[^<]*(<\/c:formatCode>)/, '$10.00$2')
)
await variant('axis-unlinked', (xml) =>
	xml.replace('<c:numFmt formatCode="0.00" sourceLinked="1"/>', '<c:numFmt formatCode="0.00" sourceLinked="0"/>')
)
