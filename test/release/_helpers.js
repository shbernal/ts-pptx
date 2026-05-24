'use strict'

// Shared helpers for release-time tests under `test/release/`.
//
// Provides:
//   * `expectNoSchemaErrors(filePath, label)` — runs the OOXML validator on
//     the given `.pptx` path and throws an Error summarising the first few
//     errors when validation fails. Returns silently on a clean file.
//   * `withTempDir(prefix, fn)` — creates a fresh tempdir under
//     `os.tmpdir()`, awaits `fn(dir)`, and removes the tempdir afterwards
//     even if `fn` throws.
//   * `assert(cond, msg)` — throws an Error tagged `assertion failed:` so
//     misuse surfaces clearly in the runner output.
//
// Kept separate from `_runner.js` so individual test files don't pull in
// the runner's bundle-build orchestration just to assert validity.

const fs = require('fs')
const os = require('os')
const path = require('path')

const { runValidatorOnFile } = require('../validator')

function assert (cond, msg) {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

async function expectNoSchemaErrors (filePath, label) {
	const errors = await runValidatorOnFile(filePath)
	if (errors.length === 0) return
	const summary = errors
		.slice(0, 5)
		.map(e => `  - [${e.ErrorType}] ${e.Description} (path: ${(e.Path && e.Path.PartUri) || '?'})`)
		.join('\n')
	const more = errors.length > 5 ? `\n  ...(${errors.length - 5} more)` : ''
	assert(false, `${label}: ${errors.length} schema error(s):\n${summary}${more}`)
}

async function withTempDir (prefix, fn) {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix))
	try {
		return await fn(dir)
	} finally {
		try { fs.rmSync(dir, { recursive: true, force: true }) } catch (_) { /* ignore */ }
	}
}

module.exports = { assert, expectNoSchemaErrors, withTempDir }
