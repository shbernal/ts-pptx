'use strict'

// End-to-end test harness for the Node demo scripts under `demos/node/`.
//
// Covers the three release-relevant scripts:
//   * `demo`        — `node demo.js`        — single-slide diagnostic deck.
//   * `demo-all`    — `node demo.js All`    — runs `runEveryTest()`, large deck.
//   * `demo-stream` — `node demo_stream.js` — Express server on :3000, GET /
//                                              returns the deck as a stream.
//
// Strategy:
//   1. Setup ensures `demos/node/node_modules/` exists (runs
//      `npm --prefix demos/node install` once if missing). This is the same
//      step a contributor performs by hand the first time they try the
//      Node demos.
//   2. The Node demos resolve `import pptxgen from "pptxgenjs"` to the
//      repo's `dist/pptxgen.cjs.js`. The runner refreshes that file via
//      `npm run build` before this test runs (and restores it on exit) so
//      the demos exercise live source rather than the published v4.0.1
//      release artefact.
//   3. Per-case work happens inside `withTempDir`:
//        - `demo`/`demo-all` spawn `node demo.js [All]` with `cwd` set to
//          the tempdir. `demo.js` calls `pptx.writeFile({ fileName })`
//          which writes to `process.cwd()`, so the .pptx lands in the
//          tempdir. We glob for any `*.pptx` > 0 bytes and validate.
//        - `demo-stream` spawns `node demo_stream.js`, polls
//          `http://localhost:3000/` until it responds (the Express
//          listener takes ~50-200ms to bind), saves the response body to
//          the tempdir, validates, and SIGINTs the child.
//
// Transient prerequisites:
//   * `demos/node/demo.js` includes `slide.addImage({ path: "https://..." })`
//     fetching `cc_logo.jpg` from raw.githubusercontent.com. The test
//     therefore requires network access. Fails loudly when offline.
//   * `npm --prefix demos/node install` requires registry access on first run.

const { spawn } = require('child_process')
const fs = require('fs')
const http = require('http')
const path = require('path')

const { isInstalled } = require('../validator')
const { assert, expectNoSchemaErrors, withTempDir } = require('./_helpers')

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const NODE_DIR = path.join(REPO_ROOT, 'demos', 'node')
const NODE_MODULES = path.join(NODE_DIR, 'node_modules')

const DEMO_TIMEOUT_MS = 60000
const ALL_TIMEOUT_MS = 120000
const STREAM_READY_TIMEOUT_MS = 15000
const STREAM_POLL_INTERVAL_MS = 250

function runChild (cmd, args, opts) {
	return new Promise((resolve, reject) => {
		const child = spawn(cmd, args, Object.assign({ stdio: ['ignore', 'pipe', 'pipe'] }, opts || {}))
		let stdout = ''
		let stderr = ''
		const timer = (opts && opts.timeoutMs)
			? setTimeout(() => {
				try { child.kill('SIGKILL') } catch (_) { /* ignore */ }
				reject(new Error(cmd + ' ' + args.join(' ') + ' timed out after ' + opts.timeoutMs + 'ms; stdout=' + stdout.slice(-500) + ' stderr=' + stderr.slice(-500)))
			}, opts.timeoutMs)
			: null
		child.stdout.on('data', d => { stdout += d.toString('utf8') })
		child.stderr.on('data', d => { stderr += d.toString('utf8') })
		child.on('error', err => {
			if (timer) clearTimeout(timer)
			reject(err)
		})
		child.on('exit', (code, signal) => {
			if (timer) clearTimeout(timer)
			if (code === 0) return resolve({ stdout, stderr })
			reject(new Error(cmd + ' ' + args.join(' ') + ' exited with code=' + code + ' signal=' + signal + '; stderr=' + stderr.slice(-1000)))
		})
	})
}

async function setup () {
	if (!isInstalled()) throw new Error('OOXMLValidatorCLI not installed; run ./tools/ooxml-validator/install.sh')
	if (!fs.existsSync(NODE_DIR)) throw new Error('demos/node missing at ' + NODE_DIR)

	// Idempotent install. `node_modules` is git-ignored under demos/node so a
	// fresh clone never has it; subsequent runs reuse the previous install.
	if (!fs.existsSync(NODE_MODULES)) {
		console.log('  [node] installing demos/node/node_modules (one-time)...')
		const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
		await runChild(npm, ['install', '--no-audit', '--no-fund', '--prefix', NODE_DIR], {
			cwd: REPO_ROOT,
			timeoutMs: 120000,
			env: Object.assign({}, process.env, { CI: '1' })
		})
		// `npm install` rewrites package-lock.json when the committed lock
		// disagrees with package.json (the demos/node lock pins an older
		// pptxgenjs). Restore it so the harness leaves the working tree clean.
		require('child_process').spawnSync('git', ['-P', 'checkout', '--', 'demos/node/package-lock.json'], {
			cwd: REPO_ROOT, encoding: 'utf8'
		})
	}
}

async function runDemoCase (opts) {
	const { args, label, timeoutMs } = opts
	await withTempDir('pptxgen-release-node-', async (tmpDir) => {
		// `demo.js` resolves images via paths relative to its own directory
		// (e.g. `../common/images/cc_logo.jpg`). Spawning from a tempdir would
		// break that resolution and cause `writeFile` to reject silently. We
		// spawn from NODE_DIR so the demo's relative paths work, then capture
		// the new .pptx by snapshotting before/after.
		const before = new Set(fs.readdirSync(NODE_DIR).filter(f => f.toLowerCase().endsWith('.pptx')))
		const demoScript = path.join(NODE_DIR, 'demo.js')
		try {
			await runChild(process.execPath, [demoScript].concat(args || []), {
				cwd: NODE_DIR,
				timeoutMs: timeoutMs || DEMO_TIMEOUT_MS,
				env: Object.assign({}, process.env, { CI: '1' })
			})
			const after = fs.readdirSync(NODE_DIR).filter(f => f.toLowerCase().endsWith('.pptx'))
			const fresh = after
				.filter(f => !before.has(f))
				.map(f => ({ name: f, full: path.join(NODE_DIR, f), size: fs.statSync(path.join(NODE_DIR, f)).size }))
				.filter(x => x.size > 0)
				.sort((a, b) => b.size - a.size)
			assert(
				fresh.length >= 1,
				label + ': expected at least one new .pptx in ' + NODE_DIR + ' after `node demo.js ' + (args || []).join(' ') + '`; before=' + JSON.stringify([...before]) + ' after=' + JSON.stringify(after)
			)
			// Move the largest fresh deck into the tempdir for validation. The
			// move keeps NODE_DIR clean so a subsequent run sees a fresh
			// "before" snapshot and we don't leak artefacts into the working
			// tree.
			const target = fresh[0]
			const tmpFile = path.join(tmpDir, target.name)
			fs.renameSync(target.full, tmpFile)
			// Also remove any other fresh .pptx files (extra writes in
			// demo.js's no-arg path, etc.) so working tree stays clean.
			for (const extra of fresh.slice(1)) {
				try { fs.unlinkSync(extra.full) } catch (_) { /* ignore */ }
			}
			await expectNoSchemaErrors(tmpFile, label)
		} catch (err) {
			// On failure, clean up any fresh artefacts we left in NODE_DIR.
			const after = fs.readdirSync(NODE_DIR).filter(f => f.toLowerCase().endsWith('.pptx'))
			for (const f of after) {
				if (!before.has(f)) {
					try { fs.unlinkSync(path.join(NODE_DIR, f)) } catch (_) { /* ignore */ }
				}
			}
			throw err
		}
	})
}

function fetchOnce (url) {
	return new Promise((resolve, reject) => {
		const req = http.get(url, res => {
			const chunks = []
			res.on('data', c => chunks.push(c))
			res.on('end', () => {
				if (res.statusCode !== 200) return reject(new Error('GET ' + url + ' status=' + res.statusCode))
				resolve(Buffer.concat(chunks))
			})
			res.on('error', reject)
		})
		req.on('error', reject)
		req.setTimeout(30000, () => {
			try { req.destroy(new Error('socket timeout')) } catch (_) { /* ignore */ }
		})
	})
}

async function pollUntilReady (url, deadline) {
	let lastErr
	while (Date.now() < deadline) {
		try {
			return await fetchOnce(url)
		} catch (e) {
			lastErr = e
			await new Promise(r => setTimeout(r, STREAM_POLL_INTERVAL_MS))
		}
	}
	throw new Error('stream server did not respond at ' + url + ' before deadline; lastErr=' + (lastErr && lastErr.message))
}

function killAndWait (child) {
	return new Promise(resolve => {
		if (!child || child.exitCode !== null) return resolve()
		const t = setTimeout(() => {
			try { child.kill('SIGKILL') } catch (_) { /* ignore */ }
		}, 3000)
		child.once('exit', () => {
			clearTimeout(t)
			resolve()
		})
		try { child.kill('SIGINT') } catch (_) { /* ignore */ }
	})
}

async function runStreamCase () {
	await withTempDir('pptxgen-release-node-stream-', async (tmpDir) => {
		const streamScript = path.join(NODE_DIR, 'demo_stream.js')
		const child = spawn(process.execPath, [streamScript], {
			cwd: NODE_DIR,
			stdio: ['ignore', 'pipe', 'pipe'],
			env: Object.assign({}, process.env, { CI: '1' })
		})
		let stderrBuf = ''
		child.stderr.on('data', d => { stderrBuf += d.toString('utf8') })
		// Discard stdout but observe it for early exits.
		child.stdout.on('data', () => { /* drain */ })

		// If the child dies before we GET, fail with the captured stderr.
		const earlyExit = new Promise((_, reject) => {
			child.once('exit', (code, signal) => {
				reject(new Error('demo_stream.js exited unexpectedly before serving: code=' + code + ' signal=' + signal + ' stderr=' + stderrBuf.slice(-1000)))
			})
		})

		try {
			const deadline = Date.now() + STREAM_READY_TIMEOUT_MS
			const body = await Promise.race([
				pollUntilReady('http://127.0.0.1:3000/', deadline),
				earlyExit
			])
			assert(Buffer.isBuffer(body) && body.length > 0, 'stream/demo_stream: empty response body')
			const out = path.join(tmpDir, 'PptxGenJS_Node_Demo_Stream.pptx')
			fs.writeFileSync(out, body)
			await expectNoSchemaErrors(out, 'node/demo-stream')
		} finally {
			await killAndWait(child)
		}
	})
}

const KF_LOWERCASE_GUID = 'getUuid emits lowercase hex; OOXML schema requires uppercase A-F (src/gen-utils.ts getUuid)'
const KF_ZIP_DIR_ENTRIES = 'zip.folder() emits directory entries that OpenXmlValidator rejects with OpenXmlPackageException (src/pptxgen.ts ~line 501)'
const KF_DEMO_ALL_CASCADE = 'demo-all runs runEveryTest() which cascades both KF_LOWERCASE_GUID (sections) and KF_ZIP_DIR_ENTRIES'

module.exports = {
	setup,
	cases: [
		{
			name: 'node/demo: produces a schema-valid .pptx',
			fn: () => runDemoCase({ args: [], label: 'node/demo' }),
			// `demo.js` no-arg path adds an `addImage({ path: "..." })` from
			// raw.githubusercontent.com (network-dependent) and one `addImage`
			// from a local jpg. The deck does NOT call `addSection`, so it
			// avoids both KF_LOWERCASE_GUID and KF_ZIP_DIR_ENTRIES paths.
			// We expect this case to PASS.
		},
		{
			name: 'node/demo-all: produces a schema-valid .pptx',
			fn: () => runDemoCase({ args: ['All'], label: 'node/demo-all', timeoutMs: ALL_TIMEOUT_MS }),
			// `runEveryTest()` ⇒ `execGenSlidesFuncs(['Master','Chart',...])`
			// which exercises the same gen-xml paths the failing browser cases
			// hit. Marked `knownFailure` until S7+S8 land.
			knownFailure: KF_DEMO_ALL_CASCADE
		},
		{
			name: 'node/demo-stream: produces a schema-valid .pptx',
			fn: runStreamCase
			// The stream demo writes a single text-only slide with no sections,
			// no charts, and no images. Should be schema-clean. PASS expected.
		}
	]
}
