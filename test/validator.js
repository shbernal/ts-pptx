// Test-time helper: validate a `.pptx` Buffer against Microsoft's
// OpenXmlValidator (via the OOXML-Validator CLI binary installed under
// tools/ooxml-validator/bin/). Returns a Promise<Array> of validation
// errors (empty array on a clean file).

import { execFile as execFileCallback } from 'node:child_process'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const execFile = promisify(execFileCallback)

// The upstream release ships the binary as `OOXMLValidatorCLI` on
// Linux/macOS and `OOXMLValidatorCLI.exe` on Windows (install.sh probes
// both). Resolve to whichever this platform installed so `isInstalled`
// and `execFile` agree — a bare extensionless path never resolves on
// Windows, silently disabling schema validation there.
const BIN_DIR = path.resolve(__dirname, '..', 'tools', 'ooxml-validator', 'bin')
const VALIDATOR = path.join(BIN_DIR, process.platform === 'win32' ? 'OOXMLValidatorCLI.exe' : 'OOXMLValidatorCLI')

// The conformance target this project validates against, pinned HERE rather than
// inherited from the CLI's own default.
//
// Two different defaults are in play upstream and they disagree: the Open XML SDK's
// `new OpenXmlValidator()` defaults to `Office2007`, while the OOXMLValidatorCLI
// wrapper defaults to `Microsoft365`. Passing nothing meant the project's conformance
// bar was whichever the pinned wrapper release happened to choose — a bump in
// `tools/ooxml-validator/version.json` could have moved it silently.
//
// `Microsoft365` is also the STRONGEST available check, not merely the newest. The
// SDK's per-version schemas differ only in how much markup they model: an older
// version does not reject newer constructs, it skips them. Measured on this
// codebase (see `pnpm run schema:versions`), a chartEx deck reports 0 errors at
// Office2007/2010/2013 and 4 at Office2016+ — the older runs are blind, not
// permissive — while a corrupted core `<p:sp>` attribute is caught identically at
// every version. Error count is therefore monotonically non-decreasing in version,
// so validating anywhere below `Microsoft365` can only lose coverage.
const FILE_FORMAT = 'Microsoft365'

// Every value the CLI accepts, oldest first. Only `schema:versions` iterates these;
// the suites all validate at FILE_FORMAT. Order matters — the probe reads it as the
// coverage axis and flags any fixture whose error count decreases along it.
const FILE_FORMATS = [
	'Office2007',
	'Office2010',
	'Office2013',
	'Office2016',
	'Office2019',
	'Office2021',
	'Microsoft365',
]

async function isInstalled() {
	try {
		await fs.access(VALIDATOR)
		return true
	} catch {
		return false
	}
}

let noticeEmitted = false

// The gate every `test.skipIf(!validatorInstalled)` site should be built on,
// rather than calling `isInstalled()` directly.
//
// A missing validator silently skips a few hundred schema assertions, so a local
// `pnpm run verify` can be green while proving far less than it appears to. That
// is a reasonable local trade — the binary is a large download and is not
// committed — but it is never acceptable in CI, where installing it is a step of
// the job. So: hard failure under CI, and locally a single one-line notice so a
// green run cannot quietly be mistaken for a complete one.
async function validatorAvailable() {
	if (await isInstalled()) return true
	if (process.env.CI) {
		throw new Error(
			'OOXMLValidatorCLI not installed at ' +
				VALIDATOR +
				'\nSchema assertions must not be skipped in CI. Run: ./tools/ooxml-validator/install.sh'
		)
	}
	if (!noticeEmitted) {
		noticeEmitted = true
		// `process.stderr.write`, not `console.warn`: this fires while the module
		// graph is still being collected, and Vitest drops console output emitted
		// outside a running test. A notice nobody sees defeats the entire point.
		process.stderr.write(
			'\n[validator] OOXMLValidatorCLI not installed — schema assertions are being SKIPPED.\n' +
				'[validator] A green run here does NOT prove schema validity.\n' +
				'[validator] Install with: ./tools/ooxml-validator/install.sh\n\n'
		)
	}
	return false
}

async function runValidatorOnFile(filePath, fileFormat = FILE_FORMAT) {
	const args = [filePath]
	const env = {
		...process.env,
		DOTNET_BUNDLE_EXTRACT_BASE_DIR: process.env.DOTNET_BUNDLE_EXTRACT_BASE_DIR || os.tmpdir(),
	}
	// Always explicit — never fall through to the CLI's own default (see FILE_FORMAT).
	args.push(fileFormat)

	// The CLI prints a JSON array to stdout regardless of whether
	// errors were found; exit code is 0 in both cases.
	const { stdout } = await execFile(VALIDATOR, args, { env, maxBuffer: 32 * 1024 * 1024 })
	try {
		return JSON.parse(stdout || '[]')
	} catch {
		throw new Error('failed to parse OOXMLValidatorCLI output: ' + String(stdout).slice(0, 500))
	}
}

async function validateBuf(buf, fileFormat = FILE_FORMAT) {
	if (!(await isInstalled())) {
		throw new Error('OOXMLValidatorCLI not installed. Run ./tools/ooxml-validator/install.sh')
	}
	const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'TsPptx-schema-'))
	const tmp = path.join(tmpDir, 'fixture.pptx')
	await fs.writeFile(tmp, buf)
	try {
		return await runValidatorOnFile(tmp, fileFormat)
	} finally {
		await fs.rm(tmpDir, { recursive: true, force: true })
	}
}

export { isInstalled, validatorAvailable, validateBuf, runValidatorOnFile, VALIDATOR, FILE_FORMAT, FILE_FORMATS }
