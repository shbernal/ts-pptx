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

// ---------------------------------------------------------------------------
// Batching
//
// The CLI accepts a *directory* and validates every package in it in ONE
// process. That matters because the binary is a 110 MB self-contained .NET
// single-file app: measured here, a run costs ~0.40s of startup plus only
// ~0.048s per additional deck, and ~55 MB of RSS regardless. Validating one
// deck per process therefore paid the 0.40s and the 55 MB about 500 times per
// suite, and — because the spawn happened inside a `describe.concurrent` test —
// the process count scaled with `maxConcurrency` × the worker pool. That
// product, not the tests themselves, is what put the memory ceiling out of the
// repo's hands and into the host's core count.
//
// The queue below is a dataloader: requests accumulate while an invocation is
// in flight and go out as one batch when it returns. It is self-tuning — under
// load batches grow, when idle the delay is a single timer tick — and it bounds
// this process to at most ONE validator child at a time, so a fork's validator
// memory is ~55 MB flat instead of ~55 MB × maxConcurrency.
//
// The directory mode's contract, established empirically against this pinned
// binary (see docs/testing.md "Validator batching") rather than assumed:
//   - a package with >= 1 error is reported, keyed by absolute `FilePath`, with
//     the same error list single-file mode gives it;
//   - a CLEAN package is omitted from the output entirely;
//   - an unreadable or non-package file is NOT omitted — it is reported as an
//     `OpenXmlPackageException`. This is the property that makes "absent means
//     clean" safe. Were corrupt packages silently dropped instead, this batcher
//     would report the worst failures the suite exists to catch as passes.
//   - results come back in arbitrary order, so they are keyed by filename, never
//     by position;
//   - the scan is not recursive, so a flat batch directory is sufficient.
//
// Set TSPPTX_VALIDATOR_NO_BATCH=1 to bypass all of this and validate one deck
// per process, for when a batch failure needs to be pinned to a single fixture
// by hand.
// ---------------------------------------------------------------------------

const BATCH_DISABLED = process.env.TSPPTX_VALIDATOR_NO_BATCH === '1'
// Caps stdout size and batch-directory disk, not memory — the CLI's RSS does not
// grow meaningfully with file count. Well above the ~8 a concurrent suite offers.
const MAX_BATCH = 32

/** @type {{buf: Buffer, fileFormat: string, resolve: Function, reject: Function}[]} */
let queue = []
let flushScheduled = false
let inFlight = false

function scheduleFlush() {
	// While an invocation is in flight, new work simply accumulates; the flush
	// that is running re-arms this on the way out. That is what keeps the child
	// count at one without a semaphore.
	if (flushScheduled || inFlight) return
	flushScheduled = true
	setTimeout(() => {
		flushScheduled = false
		void flush()
	}, 0)
}

async function flush() {
	if (inFlight || queue.length === 0) return
	inFlight = true
	// `fileFormat` is an argument to the whole invocation, so a batch is only
	// ever one format; anything else waits for the next round. The queue is
	// non-empty here (guarded above), which `noUncheckedIndexedAccess` cannot see.
	const head = /** @type {{buf: Buffer, fileFormat: string, resolve: Function, reject: Function}} */ (queue[0])
	const fileFormat = head.fileFormat
	const batch = []
	const deferred = []
	for (const item of queue) {
		if (item.fileFormat === fileFormat && batch.length < MAX_BATCH) batch.push(item)
		else deferred.push(item)
	}
	queue = deferred
	try {
		await runBatch(batch, fileFormat)
	} finally {
		inFlight = false
		if (queue.length > 0) scheduleFlush()
	}
}

async function runBatch(batch, fileFormat) {
	const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'TsPptx-schema-batch-'))
	try {
		// Positional names, so a result can be mapped back to its request without
		// trusting either output order or the caller's own naming.
		const names = batch.map((_, i) => 'b' + i + '.pptx')
		await Promise.all(batch.map((item, i) => fs.writeFile(path.join(dir, names[i]), item.buf)))

		let rows
		try {
			const { stdout } = await execFile(VALIDATOR, [dir, fileFormat], {
				env: {
					...process.env,
					DOTNET_BUNDLE_EXTRACT_BASE_DIR: process.env.DOTNET_BUNDLE_EXTRACT_BASE_DIR || os.tmpdir(),
				},
				maxBuffer: 128 * 1024 * 1024,
			})
			rows = JSON.parse(stdout || '[]')
		} catch {
			// One bad deck must not turn into 32 indistinguishable failures. Re-run the
			// batch one deck per process so the error lands on the request that caused
			// it and its neighbours still get real verdicts. Costs a slow path only on
			// a failure that would otherwise have been unattributable.
			//
			// The batch-level error is deliberately dropped: it describes an invocation
			// covering 32 decks, so it can only be less specific than what the retry is
			// about to produce for each of them.
			await Promise.all(
				batch.map(async (item) => {
					try {
						item.resolve(await validateBufDirect(item.buf, item.fileFormat))
					} catch (individualErr) {
						item.reject(individualErr)
					}
				})
			)
			return
		}

		/** @type {Map<string, unknown[]>} */
		const byName = new Map()
		for (const row of rows) {
			byName.set(path.basename(row.FilePath), JSON.parse(row.ValidationErrors || '[]'))
		}
		// Absent means clean — safe only because corrupt packages are reported, not
		// dropped (see the contract above).
		batch.forEach((item, i) => item.resolve(byName.get(names[i]) ?? []))
	} finally {
		await fs.rm(dir, { recursive: true, force: true })
	}
}

/** One deck, one process — the pre-batching path, kept for the fallback and the opt-out. */
async function validateBufDirect(buf, fileFormat) {
	const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'TsPptx-schema-'))
	const tmp = path.join(tmpDir, 'fixture.pptx')
	await fs.writeFile(tmp, buf)
	try {
		return await runValidatorOnFile(tmp, fileFormat)
	} finally {
		await fs.rm(tmpDir, { recursive: true, force: true })
	}
}

async function validateBuf(buf, fileFormat = FILE_FORMAT) {
	if (!(await isInstalled())) {
		throw new Error('OOXMLValidatorCLI not installed. Run ./tools/ooxml-validator/install.sh')
	}
	if (BATCH_DISABLED) return validateBufDirect(buf, fileFormat)
	return new Promise((resolve, reject) => {
		queue.push({ buf, fileFormat, resolve, reject })
		scheduleFlush()
	})
}

export { isInstalled, validatorAvailable, validateBuf, runValidatorOnFile, VALIDATOR, FILE_FORMAT, FILE_FORMATS }
