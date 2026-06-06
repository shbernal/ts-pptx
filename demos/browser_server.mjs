/**
 * NAME: browser_server.mjs
 * DESC: Local web server for ./demos/browser/index.html
 * DESC: since we moved to modules (`jsm`) with v3.6.0, merely opening the local file in browser gives CORS errors
 * REQS: express
 * USAGE: `node browser_server.mjs`
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20210602
 */

// Use `createRequire` as `require` wont work by default in modules
import { createRequire } from "module";
const require = createRequire(import.meta.url);

import express from "express";
const app = express();
const port = 8000;
const DEMO_URL = `http://localhost:${port}/browser/index.html`;

app.use("/browser", express.static("./browser"));
app.use("/common", express.static("./common"));
app.use("/modules", express.static("./modules"));

app.listen(port, () => {
	console.log(`\n----------------------==~==~==~==[ SERVER RUNNING ]==~==~==~==----------------------\n`);
	console.log(`The pptxgenjs browser demo is now live at: ${DEMO_URL}\n`);
	console.log(`(Press Ctrl-C to stop)`);
	console.log(`\n----------------------==~==~==~==[ SERVER RUNNING ]==~==~==~==----------------------\n`);
});

// Skip the auto-open when running headless or in CI. `PPTXGEN_NO_OPEN=1` is an
// explicit opt-out for any other caller that doesn't want a browser tab.
if (!process.env.CI && !process.env.PPTXGEN_NO_OPEN) {
	let start = process.platform == "darwin" ? "open" : process.platform == "win32" ? "start" : "xdg-open";
	require("child_process").exec(start + " " + DEMO_URL);
}
