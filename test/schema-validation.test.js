import { beforeAll, describe, test } from 'vitest'
import TsPptx from '../dist/node.js'
import { isInstalled, validateBuf } from './validator.js'
import cases from './schema-cases.js'

// Most fixtures run concurrently (see `describe.concurrent` below). Validate one
// minimal deck serially first: the oracle is a .NET single-file app whose first
// invocation self-extracts its bundle into DOTNET_BUNDLE_EXTRACT_BASE_DIR, and
// several processes racing on a cold extract directory can collide. It also gets
// the one-time download out of the way before anything concurrent starts. This is
// file-level rather than attached to one suite so both blocks below are covered
// even when only one of them is selected by `-t`.
beforeAll(async () => {
	if (!(await isInstalled())) {
		throw new Error(
			'the ooxml-validate oracle could not be obtained, so this suite would prove nothing.\n' +
				'It is fetched from GitHub Releases on first use and cached under ~/.cache/ooxml-validate;\n' +
				'see docs/testing.md if this machine cannot reach it.'
		)
	}
	const pres = new TsPptx()
	pres.addSlide()
	await validateBuf(/** @type {Uint8Array} */ (await pres.toBytes()))
})

/**
 * Does this fixture take over a process-global for the duration of its body?
 *
 * A handful of cases assert on emitted warnings by swapping `console.warn` for a
 * collector and restoring it in a `finally`. That is only sound if nothing else
 * runs in between — and under `describe.concurrent` something does: a neighbour's
 * `finally` restores the *original* `console.warn` while this fixture is still
 * inside its capture window, so its warnings go to the terminal and its assertion
 * fails. Nothing in the deck being built is wrong when that happens, which is what
 * makes it such an expensive failure to read.
 *
 * It is a race, so it was always latent — it surfaced when validator batching
 * (test/validator.js) changed how fixtures interleave, not because batching broke
 * anything. Routing these to a sequential block is the fix that matches the cause:
 * a fixture that needs exclusive access to a global must not share the process.
 *
 * Detection is by source inspection *plus* an explicit `exclusive: true` opt-out,
 * so a newly added warn-capturing fixture is quarantined automatically instead of
 * silently rejoining the concurrent block and reintroducing a load-dependent flake.
 * If a fixture ever captures warnings through a helper rather than inline, set the
 * flag on it — the sniff cannot see through a call boundary.
 *
 * The scalable alternative, if this set grows much past a handful, is to route
 * `setDiagnosticHandler` (src/diagnostics.ts) through an `AsyncLocalStorage` sink
 * so each concurrent fixture collects into its own store. That is a real fix for
 * concurrency rather than an avoidance of it, and it is not worth its cost for 7
 * of 151 fixtures.
 * @param {{ name: string, fn: Function, exclusive?: boolean }} fixture
 */
function needsExclusiveProcess(fixture) {
	return fixture.exclusive === true || /console\.warn\s*=/.test(String(fixture.fn))
}

const concurrentCases = cases.filter((f) => !needsExclusiveProcess(f))
const exclusiveCases = cases.filter(needsExclusiveProcess)

describe.concurrent('TsPptx schema validation fixtures', () => {
	for (const fixture of concurrentCases) {
		test(fixture.name, async () => {
			await fixture.fn()
		})
	}
})

// Sibling suites run in order, so nothing from the concurrent block above is still
// in flight once this one starts.
describe.sequential('TsPptx schema validation fixtures (exclusive process globals)', () => {
	for (const fixture of exclusiveCases) {
		test(fixture.name, async () => {
			await fixture.fn()
		})
	}
})
