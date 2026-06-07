/*
 * NAME: demo_stream.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20210410
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: install dependencies with pnpm, npm, or yarn
 *
 * USAGE: `node demo_stream.js`
 */

// ============================================================================
import { Buffer } from "node:buffer";
import pptxgen from "@shbernal/pptxgenjs";
import express from "express"; // @note Only required for streaming test (not a req for PptxGenJS)
const app = express(); // @note Only required for streaming test (not a req for PptxGenJS)
//let exportName = `PptxGenJS_Node_Demo_Stream_${new Date().toISOString()}.pptx`;
const exportName = `PptxGenJS_Node_Demo_Stream.pptx`;

// EXAMPLE: Export presentation to stream
const pptx = new pptxgen();
const slide = pptx.addSlide();
slide.addText(
	[
		{ text: "PptxGenJS", options: { fontSize: 48, color: pptx.colors.ACCENT1, breakLine: true } },
		{ text: "Node Stream Demo", options: { fontSize: 24, color: pptx.colors.ACCENT6, breakLine: true } },
		{ text: "(pretty cool huh?)", options: { fontSize: 24, color: pptx.colors.ACCENT3 } },
	],
	{ x: 1, y: 1, w: "80%", h: 3, align: "center", fill: pptx.colors.BACKGROUND2 }
);

// Export presentation: Save to stream (instead of `write` or `writeFile`)
try {
	const data = await pptx.stream();
	const body = typeof data === "string" ? Buffer.from(data, "binary") : Buffer.from(data);

	app.get("/", (_req, res) => {
		res.writeHead(200, { "Content-disposition": `attachment;filename=${exportName}`, "Content-Length": body.length });
		res.end(body);
	});

	app.listen(3000, () => {
		console.log(`\n\n--------------------==~==~==~==[ STARTING STREAM DEMO... ]==~==~==~==--------------------\n`);
		console.log(`* pptxgenjs ver: ${pptx.version}`);
		console.log(`* save location: ${process.cwd()}`);
		console.log(`\n`);
		console.log("PptxGenJS Node Stream Demo app listening on port 3000!");
		console.log("Visit: http://localhost:3000/");
		console.log(`\n`);
		console.log("(press Ctrl-C to quit demo)");
	});
} catch (err) {
	console.log("ERROR: " + err);
	console.log(`\n--------------------==~==~==~==[ ... STREAM DEMO COMPLETE ]==~==~==~==--------------------\n\n`);
}
