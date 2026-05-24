'use strict'

// Spawns the demo browser server (`demos/browser_server.mjs`) in a child
// process. The server uses `app.use("/browser", express.static("./browser"))`
// so it MUST be started from the `demos/` directory; we set `cwd` accordingly.
//
// Resolves once the "SERVER RUNNING" banner appears on stdout. The auto-`open`
// side-effect at the end of `browser_server.mjs` is fire-and-forget; we ignore
// any browser launch failure since we drive the page programmatically.

const { spawn } = require('child_process')
const path = require('path')

const DEMOS_DIR = path.resolve(__dirname, '..', '..', 'demos')
const SERVER_SCRIPT = 'browser_server.mjs'
const READY_BANNER = 'SERVER RUNNING'
const READY_TIMEOUT_MS = 10000

function startServer () {
	return new Promise((resolve, reject) => {
		const child = spawn(process.execPath, [SERVER_SCRIPT], {
			cwd: DEMOS_DIR,
			stdio: ['ignore', 'pipe', 'pipe'],
			env: Object.assign({}, process.env, { CI: '1' })
		})

		let resolved = false
		let stderrBuf = ''

		const timer = setTimeout(() => {
			if (resolved) return
			resolved = true
			try { child.kill('SIGTERM') } catch (_) { /* ignore */ }
			reject(new Error('demo server did not become ready within ' + READY_TIMEOUT_MS + 'ms; stderr=' + stderrBuf.slice(0, 1000)))
		}, READY_TIMEOUT_MS)

		child.stdout.on('data', (chunk) => {
			const s = chunk.toString('utf8')
			if (!resolved && s.includes(READY_BANNER)) {
				resolved = true
				clearTimeout(timer)
				resolve({
					child,
					kill: () => stopServer(child)
				})
			}
		})

		child.stderr.on('data', (chunk) => {
			stderrBuf += chunk.toString('utf8')
		})

		child.on('exit', (code, signal) => {
			if (!resolved) {
				resolved = true
				clearTimeout(timer)
				reject(new Error('demo server exited before ready (code=' + code + ' signal=' + signal + '); stderr=' + stderrBuf.slice(0, 1000)))
			}
		})

		child.on('error', (err) => {
			if (!resolved) {
				resolved = true
				clearTimeout(timer)
				reject(err)
			}
		})
	})
}

function stopServer (child) {
	return new Promise((resolve) => {
		if (!child || child.exitCode !== null) return resolve()
		const timer = setTimeout(() => {
			try { child.kill('SIGKILL') } catch (_) { /* ignore */ }
		}, 3000)
		child.once('exit', () => {
			clearTimeout(timer)
			resolve()
		})
		try { child.kill('SIGTERM') } catch (_) { /* ignore */ }
	})
}

module.exports = { startServer, DEMOS_DIR }
