'use strict'

// Runner for release-time end-to-end tests under `test/release/`.
//
// Mirrors the format of `test/run.js` so a future `release-test` script can
// chain it cleanly. Picks up files matching `*.test.js` (excluding files
// starting with `_` such as this runner and the server helper).
//
// Each test file exports either:
//   * an array of `{ name, fn, [knownFailure] }` objects (case-only form), or
//   * an object `{ setup?, teardown?, cases }` where `setup`/`teardown` are
//     awaited once per file (around all its cases). `teardown` is invoked even
//     if a case throws (try/finally semantics) so server child processes,
//     browsers, and any swapped files are always cleaned up.
//
// `fn` is awaited. If a case sets `knownFailure: '<reason>'`, a thrown error
// is reported as `KNOWN-FAIL: <reason>` (excluded from the Failed count) so
// the harness can still pin a deferred-bug case to the matrix without making
// the run red. Cases that pass while marked `knownFailure` are reported with
// an `UNEXPECTED-PASS` warning.
//
// Pre-test bundle build:
//   The browser harness substitutes a JSZip+pptxgenjs IIFE bundle for the
//   CDN URL referenced in `demos/browser/index.html`. The committed
//   `dist/pptxgen.bundle.js` reflects the published v4.0.1 release and may
//   not match HEAD (e.g. after a fix that has not yet been re-bundled).
//   To exercise CURRENT source, we:
//     1. Run `gulp build` to refresh `src/bld/pptxgen.gulp.js` (IIFE) from
//        the latest TypeScript sources.
//     2. Concatenate `libs/*` + `src/bld/pptxgen.gulp.js` into
//        `test/release/_tmp/pptxgen.bundle.js`.
//   This keeps `dist/` and `demos/browser/js/` untouched (they belong to the
//   release pipeline, not the inner-loop build) while guaranteeing the
//   harness tests the live code path. The path of the freshly-built bundle
//   is exposed via `process.env.PPTXGEN_LOCAL_BUNDLE`.

const { spawn } = require('child_process')
const fs = require('fs')
const path = require('path')

const REPO_ROOT = path.resolve(__dirname, '..', '..')
const BLD_IIFE = path.join(REPO_ROOT, 'src', 'bld', 'pptxgen.gulp.js')
const LIBS_DIR = path.join(REPO_ROOT, 'libs')
const TMP_DIR = path.join(__dirname, '_tmp')
const TMP_BUNDLE = path.join(TMP_DIR, 'pptxgen.bundle.js')

function runCmd (cmd, args, opts) {
	return new Promise((resolve, reject) => {
		const child = spawn(cmd, args, Object.assign({ cwd: REPO_ROOT, stdio: 'inherit' }, opts || {}))
		child.on('error', reject)
		child.on('exit', (code, signal) => {
			if (code === 0) return resolve()
			reject(new Error(cmd + ' ' + args.join(' ') + ' exited with code=' + code + ' signal=' + signal))
		})
	})
}

async function buildFreshBundle () {
	console.log('Building fresh IIFE bundle for release tests...')
	const npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'
	// `gulp build` produces `src/bld/pptxgen.gulp.js` (an IIFE registering
	// `window.PptxGenJS`) plus the cjs/es siblings. We only need the IIFE.
	await runCmd(npx, ['gulp', 'build'])
	if (!fs.existsSync(BLD_IIFE)) {
		throw new Error('expected ' + BLD_IIFE + ' after `gulp build`; the gulp task did not produce the IIFE output')
	}
	if (!fs.existsSync(LIBS_DIR)) {
		throw new Error('expected ' + LIBS_DIR + ' to exist; cannot assemble bundle without libs')
	}
	fs.mkdirSync(TMP_DIR, { recursive: true })
	const libFiles = fs.readdirSync(LIBS_DIR).filter(f => f.endsWith('.js')).sort().map(f => path.join(LIBS_DIR, f))
	const parts = libFiles.concat([BLD_IIFE])
	const out = fs.createWriteStream(TMP_BUNDLE)
	out.write('/* PptxGenJS release-test bundle @ ' + new Date().toISOString() + ' */\n')
	for (const f of parts) {
		out.write(fs.readFileSync(f))
		out.write('\n;\n') // safety separator between concatenated scripts
	}
	out.end()
	await new Promise((resolve, reject) => {
		out.on('finish', resolve)
		out.on('error', reject)
	})
	const stat = fs.statSync(TMP_BUNDLE)
	console.log('  ' + path.relative(REPO_ROOT, TMP_BUNDLE) + ': ' + stat.size + ' bytes (parts: ' + parts.map(p => path.basename(p)).join(' + ') + ')')
	process.env.PPTXGEN_LOCAL_BUNDLE = TMP_BUNDLE
}

const failures = []
const successes = []
const knownFails = [] // cases that failed but were marked `knownFailure`
const unexpectedPasses = [] // cases marked `knownFailure` that passed

async function loadAndRun () {
	const dir = __dirname
	const files = fs.readdirSync(dir)
		.filter(f => /\.test\.js$/.test(f) && !f.startsWith('_'))
		.sort()
	for (const f of files) {
		const full = path.join(dir, f)
		const mod = require(full)
		// Support two export shapes:
		//   1. Array of cases (legacy, still supported).
		//   2. Object `{ setup?, teardown?, cases }` (per-file lifecycle).
		let cases
		let setup
		let teardown
		if (Array.isArray(mod)) {
			cases = mod
		} else if (mod && Array.isArray(mod.cases)) {
			cases = mod.cases
			setup = typeof mod.setup === 'function' ? mod.setup : undefined
			teardown = typeof mod.teardown === 'function' ? mod.teardown : undefined
		} else {
			throw new Error('test module ' + f + ' must export an array of cases or `{ setup?, teardown?, cases }`')
		}

		const ctx = {}
		let setupOk = false
		try {
			if (setup) await setup(ctx)
			setupOk = true
			for (const c of cases) {
				try {
					await c.fn(ctx)
					if (c.knownFailure) {
						unexpectedPasses.push({ name: c.name, reason: c.knownFailure })
						console.log('  UNEXPECTED-PASS ' + c.name + ' (was marked knownFailure: ' + c.knownFailure + ')')
					} else {
						successes.push(c.name)
						console.log('  ok ' + c.name)
					}
				} catch (e) {
					if (c.knownFailure) {
						knownFails.push({ name: c.name, reason: c.knownFailure, error: e })
						console.log('  KNOWN-FAIL ' + c.name + ' (' + c.knownFailure + ')')
					} else {
						failures.push({ name: c.name, error: e })
						console.log('  FAIL ' + c.name + ': ' + (e && e.message ? e.message : e))
					}
				}
			}
		} catch (e) {
			// Setup failure: report as a synthetic failure so the runner exits non-zero.
			failures.push({ name: f + ' [setup]', error: e })
			console.log('  FAIL ' + f + ' [setup]: ' + (e && e.message ? e.message : e))
		} finally {
			if (teardown) {
				try {
					await teardown(ctx)
				} catch (e) {
					// Teardown failure should also fail the run — leaking server/browser
					// or a swapped file is a real problem.
					failures.push({ name: f + ' [teardown]', error: e })
					console.log('  FAIL ' + f + ' [teardown]: ' + (e && e.message ? e.message : e))
				}
			}
			void setupOk // (informational only; we run cases iff setup succeeded above)
		}
	}
}

;(async () => {
	console.log('Running PptxGenJS release-time tests')
	try {
		await buildFreshBundle()
	} catch (e) {
		console.error('Pre-test bundle build failed: ' + (e && e.message ? e.message : e))
		process.exit(2)
	}
	await loadAndRun()
	console.log(
		'\nPassed: ' + successes.length +
		'  Failed: ' + failures.length +
		'  Known-Fail: ' + knownFails.length +
		'  Unexpected-Pass: ' + unexpectedPasses.length
	)
	if (knownFails.length > 0) {
		console.log('\nKnown failures (deferred bugs, not counted as failures):')
		knownFails.forEach(k => console.log('  - ' + k.name + ' :: ' + k.reason))
	}
	if (unexpectedPasses.length > 0) {
		console.log('\nUnexpected passes (cases marked knownFailure that succeeded — review and remove the marker):')
		unexpectedPasses.forEach(u => console.log('  - ' + u.name + ' (was: ' + u.reason + ')'))
		// Unexpected passes are a soft warning: not a failure but worth surfacing
		// loudly so the marker is removed promptly.
	}
	if (failures.length > 0) {
		failures.forEach(f => console.log(f.name + ' -- ' + ((f.error && f.error.stack) || (f.error && f.error.message) || f.error)))
		process.exit(1)
	}
})()
