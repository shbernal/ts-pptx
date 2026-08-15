#!/usr/bin/env node
/**
 * Static file server for the browser lane's adapter harness.
 *
 * The site's demos page is the fixture for the *bundled* browser path (Vite resolves the
 * `browser` export condition, Rollup tree-shakes it). It cannot be the fixture for
 * `src/runtime/browser.ts`'s three loader functions, because the deck it builds draws
 * every asset and so never calls one — see docs/testing.md "Browser Lane".
 *
 * This serves the repo as-is so `test/browser/harness/index.html` can load
 * `dist/browser.js` over plain `<script type="module">`, with no bundler in the path at
 * all. Two things fall out of that:
 *
 *   - What runs in the browser is the file that ships, not a re-bundling of it. If
 *     `dist/browser.js` ever reaches a `node:*` import the page fails to load, loudly.
 *   - The harness needs real URLs for media and fonts, which is exactly what the three
 *     loaders take. A 404 under an allowed prefix is a real 404, so the failure branches
 *     are reachable without stubbing `fetch`.
 *
 * Paths are served from an allowlist of prefixes rather than from the repo root
 * wholesale: a test server has no business handing out `.git/` even on loopback.
 *
 *   node scripts/browser-harness-server.mjs [--port 4174] [--host 127.0.0.1]
 *
 * Playwright starts this as its second `webServer` (see playwright.config.ts); it is not
 * meant to be long-lived.
 */

import fs from 'node:fs'
import http from 'node:http'
import path from 'node:path'
import { ROOT, parseCliOrExit } from './script-utils.mjs'

/**
 * Repo-relative prefixes this server will serve, and nothing else.
 *
 * The two `node_modules/` mounts are the bare specifiers reachable from
 * `dist/browser.js`, which the harness page resolves through an import map: `fflate`
 * (imported statically by `dist/zip.js`) and `opentype.js` (imported *dynamically* by
 * the measure/fit chunk, so it only appears once a font is registered). Nothing else is
 * needed — `@xmldom/xmldom` lives in the `shapes` chunk, which only the read and inspect
 * entries pull in.
 */
const MOUNTS = [
	'dist/',
	'node_modules/fflate/esm/',
	'node_modules/opentype.js/dist/',
	'test/browser/harness/',
	// Media and font fixtures, reused rather than duplicated as new binaries.
	'demos/common/images/',
	'test/read/fixtures/fonts/',
]

/**
 * Content types the harness actually depends on. `loadMedia` builds its data URI from
 * the `Blob` type, which is the response `Content-Type` — and `createSvgPngPreview`
 * feeds that data URI to an `Image`, which refuses to decode it under the wrong type.
 * So this table is load-bearing, not decoration.
 */
/** @type {Record<string, string>} */
const CONTENT_TYPES = {
	'.css': 'text/css; charset=utf-8',
	'.html': 'text/html; charset=utf-8',
	'.jpg': 'image/jpeg',
	'.js': 'text/javascript; charset=utf-8',
	'.json': 'application/json; charset=utf-8',
	'.mjs': 'text/javascript; charset=utf-8',
	'.png': 'image/png',
	'.svg': 'image/svg+xml',
	'.ttf': 'font/ttf',
}

const USAGE = `Static server for the browser test harness.

  node scripts/browser-harness-server.mjs --port 4174

Options:
  --port <n>   port to listen on (default 4174)
  --host <ip>  interface to bind (default 127.0.0.1)
  -h, --help   show this message`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: { port: { type: 'string', default: '4174' }, host: { type: 'string', default: '127.0.0.1' } },
})
const PORT = Number(values.port)
if (!Number.isInteger(PORT) || PORT < 0 || PORT > 65535) {
	console.error(`--port must be an integer between 0 and 65535, got ${JSON.stringify(values.port)}`)
	process.exit(2)
}
const HOST = values.host

/**
 * Map a request path to an absolute file, or null if it is not under a mount.
 * @param {string} urlPath
 * @returns {string | null}
 */
function resolveRequest(urlPath) {
	const decoded = decodeURIComponent(urlPath.split('?')[0] ?? '')
	// A directory request serves its index, so the harness page can be reached at the
	// mount path itself and `./harness.mjs` beside it still resolves.
	const rel = (decoded.endsWith('/') ? decoded + 'index.html' : decoded).replace(/^\/+/, '')
	if (!MOUNTS.some((mount) => rel.startsWith(mount))) return null

	const file = path.resolve(ROOT, rel)
	// Defence in depth: the prefix check above is on the *request*, this one is on the
	// resolved path, so `..` cannot walk out of a mount that it entered legitimately.
	if (!file.startsWith(ROOT + path.sep)) return null
	return file
}

const server = http.createServer((req, res) => {
	const file = resolveRequest(req.url ?? '/')
	if (!file || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
		res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' })
		res.end(`404 ${req.url}`)
		return
	}
	res.writeHead(200, {
		'content-type': CONTENT_TYPES[path.extname(file).toLowerCase()] ?? 'application/octet-stream',
		// The harness is rebuilt per run and every assertion depends on reading the
		// current `dist/`; a cached chunk from a previous build would be a silent lie.
		'cache-control': 'no-store',
	})
	fs.createReadStream(file).pipe(res)
})

server.listen(PORT, HOST, () => {
	console.log(`browser harness: http://${HOST}:${PORT}/`)
})
