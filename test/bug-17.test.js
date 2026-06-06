import JSZip from 'jszip'
import PptxGenJS from '../src/bld/pptxgen.js'
import { build, listEntries, assert } from './helpers.js'

// 1x1 PNG (red pixel) for image-only deck case
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

function chartsOrEmbeddingsEntries(zip) {
	return listEntries(zip).filter(
		(p) =>
			p.startsWith('ppt/charts/') || p === 'ppt/charts' || p.startsWith('ppt/embeddings/') || p === 'ppt/embeddings'
	)
}

export default [
	{
		name: 'empty deck (text-only) does not create ppt/charts or ppt/embeddings dirs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello world', { x: 1, y: 1, w: 4, h: 1 })
			})
			const stray = chartsOrEmbeddingsEntries(zip)
			assert(
				stray.length === 0,
				'expected no charts/embeddings entries for chart-free deck; got: ' + JSON.stringify(stray)
			)
		},
	},
	{
		name: 'image-only deck does not create ppt/charts or ppt/embeddings dirs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const stray = chartsOrEmbeddingsEntries(zip)
			assert(
				stray.length === 0,
				'expected no charts/embeddings entries for image-only deck; got: ' + JSON.stringify(stray)
			)
		},
	},
	{
		name: 'chart-present deck still creates chart and embedding parts (regression)',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const data = [
				{
					name: 'Series 1',
					labels: ['Cat A', 'Cat B', 'Cat C'],
					values: [10, 20, 30],
				},
			]
			slide.addChart(pres.ChartType.bar, data, { x: 1, y: 1, w: 6, h: 3 })
			const buf = await pres.stream()
			const zip = await JSZip.loadAsync(buf)
			const entries = listEntries(zip)
			const chartEntries = entries.filter((p) => p.startsWith('ppt/charts/'))
			const embedEntries = entries.filter((p) => p.startsWith('ppt/embeddings/'))
			assert(
				chartEntries.length > 0,
				'expected ppt/charts/ entries when chart present; got: ' + JSON.stringify(entries)
			)
			assert(
				embedEntries.length > 0,
				'expected ppt/embeddings/ entries when chart present; got: ' + JSON.stringify(entries)
			)
		},
	},
]
