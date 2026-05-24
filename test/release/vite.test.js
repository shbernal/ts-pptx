'use strict'

// End-to-end test harness for the Vite demo (`demos/vite-demo/`).
//
// Strategy:
//   1. Setup ensures `demos/vite-demo/node_modules/` exists (runs
//      `npm --prefix demos/vite-demo install` once if missing). Same shape
//      as the Node test's setup. Idempotent — short-circuits on a sentinel
//      that the install populated.
//   2. After install, restore `demos/vite-demo/package-lock.json` from git
//      so `npm install`'s lockfile rewrite doesn't dirty the working tree.
//      (The committed lock pins `pptxgenjs@^3.12.0` while package.json
//      requires `^4.0.1`; npm rewrites the lock to match. Pre-existing
//      drift, not introduced by this run.)
//   3. After install, invoke `gulp reactTestCode reactTestDefs` so the
//      freshly-installed published v4.0.1 in
//      `demos/vite-demo/node_modules/pptxgenjs/` is overridden with live
//      source (`dist/pptxgen.es.js` and `types/index.d.ts`). The runner's
//      `buildFreshDist()` deliberately does NOT pre-invoke these gulp
//      tasks — doing so would create `demos/vite-demo/node_modules/pptxgenjs/`
//      before npm install runs, and npm then skips installing the rest
//      of the workspace's deps (react, vite, @types) because pptxgenjs
//      "is already there". Running gulp here, AFTER install completes,
//      avoids that trap.
//   4. Single case `vite/build` spawns `npm --prefix demos/vite-demo run
//      build`, which runs `tsc -b && vite build`. Asserts:
//        - exit 0
//        - `demos/vite-demo/dist/index.html` exists (Vite entry HTML)
//        - at least one `*.js` asset under `demos/vite-demo/dist/assets/`
//          (Vite emits hashed chunks)
//        - the entry HTML references a `/assets/...js` chunk (soft check
//          via regex; cosmetic Vite changes won't break this)
//
// Working tree:
//   * `demos/vite-demo/node_modules/` is gitignored — no cleanup needed.
//   * `demos/vite-demo/dist/` is gitignored — no cleanup needed.
//   * `demos/vite-demo/package-lock.json` IS tracked — restored after
//     install in setup.

const { spawn, spawnSync } = require('child_process')
const fs = require('fs')
const path = require('path')

const { assert } = require('./_helpers')

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const VITE_DIR = path.join(REPO_ROOT, 'demos', 'vite-demo')
const VITE_NODE_MODULES = path.join(VITE_DIR, 'node_modules')
// Sentinel that proves a successful prior install (pptxgenjs's es build is
// the runtime entrypoint imported by App.tsx, so its presence is the most
// reliable indicator that node_modules is usable).
const VITE_INSTALL_SENTINEL = path.join(VITE_NODE_MODULES, 'pptxgenjs', 'dist', 'pptxgen.es.js')
const VITE_DIST_DIR = path.join(VITE_DIR, 'dist')

const INSTALL_TIMEOUT_MS = 240000 // 4 min — sass-embedded native build can be slow on first install
const BUILD_TIMEOUT_MS = 120000 // 2 min — first-build TS + Vite + SCSS

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
			reject(new Error(cmd + ' ' + args.join(' ') + ' exited with code=' + code + ' signal=' + signal + '; stdout(tail)=' + stdout.slice(-2000) + ' stderr(tail)=' + stderr.slice(-2000)))
		})
	})
}

async function setup () {
	if (!fs.existsSync(VITE_DIR)) throw new Error('demos/vite-demo missing at ' + VITE_DIR)

	// Idempotent install. `node_modules` is git-ignored under demos/vite-demo
	// so a fresh clone never has it; subsequent runs reuse the previous install.
	if (!fs.existsSync(VITE_INSTALL_SENTINEL)) {
		console.log('  [vite] installing demos/vite-demo/node_modules (one-time, may take a couple of minutes for sass-embedded)...')
		const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
		await runChild(npm, ['install', '--no-audit', '--no-fund', '--prefix', VITE_DIR], {
			cwd: REPO_ROOT,
			timeoutMs: INSTALL_TIMEOUT_MS,
			env: Object.assign({}, process.env, { CI: '1' })
		})
		// `npm install` rewrites package-lock.json when the committed lock
		// disagrees with package.json (the demos/vite-demo lock pins an
		// older pptxgenjs and a stale `name`/`version`). Restore it so the
		// harness leaves the working tree clean.
		spawnSync('git', ['-P', 'checkout', '--', 'demos/vite-demo/package-lock.json'], {
			cwd: REPO_ROOT, encoding: 'utf8'
		})
	}

	// Invoke `gulp reactTestCode reactTestDefs` AFTER install so live
	// source overrides the freshly-installed published v4.0.1. The runner's
	// `buildFreshDist()` deliberately does NOT pre-invoke these gulp tasks
	// (doing so would poison npm's install heuristic — see _runner.js
	// `buildFreshDist()` comment). Running them here guarantees live-source
	// override on both first and subsequent runs.
	const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'
	await runChild(npx, ['gulp', 'reactTestCode', 'reactTestDefs'], {
		cwd: REPO_ROOT,
		timeoutMs: 60000,
		env: Object.assign({}, process.env, { CI: '1' })
	})

	// Sanity check: confirm the override actually landed.
	if (!fs.existsSync(VITE_INSTALL_SENTINEL)) {
		throw new Error('expected pptxgenjs es entry at ' + VITE_INSTALL_SENTINEL + ' after install + gulp; not found')
	}
}

async function runBuildCase () {
	const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
	await runChild(npm, ['--prefix', VITE_DIR, 'run', 'build'], {
		cwd: REPO_ROOT,
		timeoutMs: BUILD_TIMEOUT_MS,
		env: Object.assign({}, process.env, { CI: '1' })
	})

	// Vite emits `dist/index.html` and chunks under `dist/assets/`.
	const indexHtml = path.join(VITE_DIST_DIR, 'index.html')
	assert(fs.existsSync(indexHtml), 'expected ' + indexHtml + ' after `vite build`')
	const indexStat = fs.statSync(indexHtml)
	assert(indexStat.size > 0, 'expected non-empty ' + indexHtml + ' after `vite build`')

	const assetsDir = path.join(VITE_DIST_DIR, 'assets')
	assert(fs.existsSync(assetsDir) && fs.statSync(assetsDir).isDirectory(),
		'expected assets directory at ' + assetsDir + ' after `vite build`')
	const jsAssets = fs.readdirSync(assetsDir).filter(f => f.endsWith('.js'))
	assert(jsAssets.length >= 1,
		'expected at least one .js asset under ' + assetsDir + '; got ' + JSON.stringify(fs.readdirSync(assetsDir)))

	// Soft check: the entry HTML should reference an /assets/...js chunk.
	// Regex match keeps cosmetic Vite output changes from breaking the
	// harness while still catching gross misconfiguration (e.g. blank
	// dist/index.html or missing module wiring).
	const html = fs.readFileSync(indexHtml, 'utf8')
	assert(/assets\/[^"']+\.js/.test(html),
		'expected ' + indexHtml + ' to reference a /assets/*.js chunk; head=' + html.slice(0, 500))
}

module.exports = {
	setup,
	cases: [
		{
			name: 'vite/build: tsc + vite build produce expected dist/ outputs',
			fn: runBuildCase
		}
	]
}
