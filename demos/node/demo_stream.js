/*
 * NAME: demo_stream.js
 * AUTH: Brent Ely
 * DATE: 20210410
 * DESC: ts-pptx feature demos for Node.js
 * REQS: install dependencies with pnpm, npm, or yarn
 *
 * USAGE: `node demo_stream.js`
 */

// ============================================================================
import { Buffer } from "node:buffer";
import TsPptx from "@shbernal/ts-pptx";
import express from "express"; // @note Only required for streaming test (not a req for TsPptx)
const app = express(); // @note Only required for streaming test (not a req for TsPptx)
//let exportName = `TsPptx_Node_Demo_Stream_${new Date().toISOString()}.pptx`;
const exportName = `TsPptx_Node_Demo_Stream.pptx`;

// EXAMPLE: Export presentation to stream
const pptx = new TsPptx();
const slide = pptx.addSlide();
slide.addText(
	[
		{ text: "TsPptx", options: { fontSize: 48, color: pptx.SchemeColor.accent1, breakLine: true } },
		{ text: "Node Stream Demo", options: { fontSize: 24, color: pptx.SchemeColor.accent6, breakLine: true } },
		{ text: "(pretty cool huh?)", options: { fontSize: 24, color: pptx.SchemeColor.accent3 } },
	],
	{ x: 1, y: 1, w: "80%", h: 3, align: "center", fill: pptx.SchemeColor.background2 },
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
		console.log(`* ts-pptx ver: ${pptx.version}`);
		console.log(`* save location: ${process.cwd()}`);
		console.log(`\n`);
		console.log("TsPptx Node Stream Demo app listening on port 3000!");
		console.log("Visit: http://localhost:3000/");
		console.log(`\n`);
		console.log("(press Ctrl-C to quit demo)");
	});
} catch (err) {
	console.log("ERROR: " + err);
	console.log(`\n--------------------==~==~==~==[ ... STREAM DEMO COMPLETE ]==~==~==~==--------------------\n\n`);
}
