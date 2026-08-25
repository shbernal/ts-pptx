import {
	setDiagnosticHandler,
	defineRegressionSuite,
	build,
	captureDiagnostics,
	assert,
	assertEqual,
} from '../../helpers.js'
import { resetDiagnosticState } from '../../../dist/node.js'

// Guards `test/setup-globals.js`, which is what makes `isolate: false` safe rather than
// merely fast (vitest.config.ts, and docs/testing.md "One module registry per worker").
//
// With one module registry per worker, a diagnostic handler left installed by one test is
// still installed for every test that runs after it — including tests in other files,
// minutes later, whose diagnostics it then swallows or throws on. That reads as a bug in
// the victim, which is what makes it worth a guard rather than a convention.
//
// The two cases below are a leak and its detector, deliberately in that source order:
// `sequence.shuffle` randomizes *files* but never tests within one, so this pair is
// ordered by construction rather than by luck. Verified sensitive — with the setup file
// removed, the second case fails with the first case's error.
//
// A cross-*file* version of this was used to establish the same thing during the isolation
// change, but it can only be ordered by pinning a shuffle seed, which is exactly the
// property the shuffle exists to deny. This is the deterministic half of that proof.

/** A deck that trips exactly one `warnOnce` condition (a table margin >= 1 reads as legacy points). */
function legacyMargin(p) {
	p.addSlide().addTable([[{ text: 'x' }]], { x: 1, y: 1, w: 4, margin: 10 })
}

/** A deck that trips exactly one known condition (`columns` must be 1-16). */
function badColumns(p) {
	p.addSlide().addText('x', { x: 1, y: 1, w: 2, h: 1, columns: 99 })
}

defineRegressionSuite('Global state is reset between tests', [
	{
		name: 'a test may install a diagnostic handler and leave it installed',
		fn: () => {
			setDiagnosticHandler(() => {
				throw new Error('handler leaked out of the test that installed it')
			})
		},
	},
	{
		name: '…and the next test still gets the default handler back',
		fn: async () => {
			// No handler installed here on purpose. If the neighbour's survived, building this
			// deck throws its error rather than warning; the default handler warns instead.
			const seen = []
			const originalConsoleWarn = console.warn
			console.warn = (msg) => seen.push(String(msg))
			try {
				await build(badColumns)
			} finally {
				console.warn = originalConsoleWarn
			}
			assert(
				seen.some((line) => line.startsWith('ts-pptx: ')),
				`expected the default console handler's prefixed line; got: ${JSON.stringify(seen)}`
			)
		},
	},
	{
		// The other global with the same lifetime, and the harder one to notice: a leaked
		// `warnOnce` key makes a test go GREEN when it should go red. Both cases below are in one
		// test on purpose — the dedupe set is what is under test, so it must not be reset between
		// the halves by the `afterEach` that exists precisely to reset it.
		name: 'warnOnce dedupes within a process, and resetDiagnosticState clears it',
		fn: async () => {
			const first = await captureDiagnostics(async () => await build(legacyMargin))
			assertEqual(
				first.codes.filter((c) => c === 'margin/legacy-points').length,
				1,
				'the condition reports the first time'
			)

			const second = await captureDiagnostics(async () => await build(legacyMargin))
			assertEqual(
				second.codes.filter((c) => c === 'margin/legacy-points').length,
				0,
				'and is silent the second time — that is what warnOnce means'
			)

			resetDiagnosticState()
			const third = await captureDiagnostics(async () => await build(legacyMargin))
			assertEqual(
				third.codes.filter((c) => c === 'margin/legacy-points').length,
				1,
				'until the dedupe set is cleared, which is what a host building a second deck needs'
			)
		},
	},
])
