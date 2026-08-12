/*
 * Streaming a generated deck straight to an HTTP response.
 *
 * The showcase decks in `demos/showcases` end in `writeFile()`. This one never touches the
 * disk: `pptx.toBytes()` hands back the package bytes, which go out as the response body.
 * That is the shape a server generating a deck per request actually needs — no temp file,
 * nothing to clean up afterwards.
 *
 * Express is here only to have a server; it is not a ts-pptx dependency.
 *
 * USAGE: pnpm --dir demos/node run demo-stream   → http://localhost:3000/
 */
import TsPptx, { SchemeColor } from '@shbernal/ts-pptx'
import express from 'express'

const app = express()
const exportName = 'TsPptx_Node_Demo_Stream.pptx'

const pptx = new TsPptx()
const slide = pptx.addSlide()
slide.addText(
	[
		{ text: 'TsPptx', options: { fontSize: 48, color: SchemeColor.accent1, breakLine: true } },
		{ text: 'Node Stream Demo', options: { fontSize: 24, color: SchemeColor.accent6, breakLine: true } },
		{ text: 'Generated per request, never written to disk.', options: { fontSize: 16, color: SchemeColor.accent3 } },
	],
	{ x: 1, y: 1, w: '80%', h: 3, align: 'center', fill: SchemeColor.background2 }
)

try {
	const data = await pptx.toBytes()

	app.get('/', (_req, res) => {
		res.writeHead(200, {
			'Content-disposition': `attachment;filename=${exportName}`,
			'Content-Length': data.byteLength,
		})
		res.end(data)
	})

	app.listen(3000, () => {
		console.log(`ts-pptx ${pptx.version} — stream demo listening on http://localhost:3000/`)
		console.log('Visit it to download the generated deck. Ctrl-C to quit.')
	})
} catch (err) {
	console.error(`stream demo failed: ${err}`)
	process.exitCode = 1
}
