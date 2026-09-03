import {
	defineRegressionSuite,
	build,
	assert,
	assertEqual,
	captureDiagnostics,
	setDiagnosticHandler,
} from '../../helpers.js'

// The diagnostics seam: every library warning is a structured `{ code, message }` delivered to a
// handler a consumer can install, rather than an unconditional `console.warn` line.
//
// What is pinned here is the CONTRACT, not the wording. A diagnostic's `code` is API — a consumer
// may branch on it — while its `message` is explicitly free to change, so the assertions below
// check codes and check that the message carries no `ts-pptx:` prefix of its own (the default
// console handler owns that).

/** A tiny deck that trips exactly one known condition. */
function badColumns(p) {
	// `columns` must be 1-16; 99 is rejected and the option ignored.
	p.addSlide().addText('x', { x: 1, y: 1, w: 2, h: 1, columns: 99 })
}

defineRegressionSuite('Diagnostics handler', [
	{
		name: 'an installed handler receives a structured diagnostic instead of a console line',
		fn: async () => {
			const seen = []
			const originalConsoleWarn = console.warn
			console.warn = (msg) => seen.push(String(msg))
			let captured
			try {
				captured = await captureDiagnostics(() => build(badColumns))
			} finally {
				console.warn = originalConsoleWarn
			}

			assert(
				captured.codes.includes('text/invalid-columns'),
				'expected the text/invalid-columns code; got: ' + JSON.stringify(captured.codes)
			)
			assertEqual(seen.length, 0, 'an installed handler must suppress the console fallback entirely')

			const diagnostic = captured.diagnostics.find((d) => d.code === 'text/invalid-columns')
			assertEqual(typeof diagnostic.message, 'string', 'a diagnostic carries a message string')
			assert(diagnostic.message.length > 0, 'the message is non-empty')
			// The prefix belongs to the default console handler, not to the message. A message that
			// carried its own would double up the moment the default handler printed it.
			assert(
				!diagnostic.message.startsWith('ts-pptx:'),
				'the message must not carry the console prefix; got: ' + diagnostic.message
			)
		},
	},
	{
		name: 'setDiagnosticHandler(null) restores the prefixed console default',
		fn: async () => {
			// Install and remove a handler, then confirm the console default is back.
			setDiagnosticHandler(() => {})
			setDiagnosticHandler(null)

			const seen = []
			const originalConsoleWarn = console.warn
			console.warn = (msg) => seen.push(String(msg))
			try {
				await build(badColumns)
			} finally {
				console.warn = originalConsoleWarn
			}

			assert(seen.length > 0, 'expected the console default to emit after the handler was removed')
			assert(
				seen.every((line) => line.startsWith('ts-pptx: ')),
				'the default handler stamps the library prefix; got: ' + JSON.stringify(seen)
			)
			assert(
				seen.some((line) => /columns/.test(line)),
				'expected the columns warning on the console; got: ' + JSON.stringify(seen)
			)
		},
	},
	{
		name: 'a throwing handler propagates out of the emitting call (the documented strict mode)',
		fn: async () => {
			class Escalated extends Error {}
			setDiagnosticHandler((d) => {
				if (d.code === 'text/invalid-columns') throw new Escalated(d.code)
			})
			let thrown = null
			try {
				await build(badColumns)
			} catch (err) {
				thrown = err
			} finally {
				setDiagnosticHandler(null)
			}
			assert(thrown instanceof Escalated, 'expected the handler’s throw to reach the caller; got: ' + thrown)
			assertEqual(thrown.message, 'text/invalid-columns', 'the handler saw the code it branched on')
		},
	},
	{
		name: 'warnOnce reports a repeated condition once but a different value again',
		fn: async () => {
			// `fontSize` out of range is a warnOnce condition, deduped on code + message — so the same
			// offending value is reported once no matter how many runs it appears in, while a different
			// value is a different message and reports on its own.
			const first = await captureDiagnostics(() =>
				build((p) => {
					const slide = p.addSlide()
					slide.addText('a', { x: 1, y: 1, w: 2, h: 1, fontSize: 4111 })
					slide.addText('b', { x: 1, y: 3, w: 2, h: 1, fontSize: 4111 })
				})
			)
			const firstHits = first.codes.filter((c) => c === 'font/size-out-of-range')
			assertEqual(firstHits.length, 1, 'the same offending value reports once; got: ' + JSON.stringify(first.messages))

			const second = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addText('c', { x: 1, y: 1, w: 2, h: 1, fontSize: 4222 })
				})
			)
			assert(
				second.codes.includes('font/size-out-of-range'),
				'a different offending value is a different message and reports again; got: ' + JSON.stringify(second.messages)
			)
		},
	},
	{
		name: 'codes are shared across entry points when the condition is the same',
		fn: async () => {
			// One condition, several authoring surfaces. Each of these is a separate call site in a
			// separate module; they agree on the code because a consumer branching on "this deck asked
			// for something unusable" should not have to know which surface reported it.
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					const slide = p.addSlide()
					slide.addText('x', { x: 1, y: 1, w: 2, h: 1, columns: 99 })
					slide.addText('y', { x: 1, y: 2, w: 2, h: 1, columnSpacing: -5 })
					p.addSection({})
				})
			)
			for (const expected of ['text/invalid-columns', 'text/invalid-column-spacing', 'section/missing-title']) {
				assert(codes.includes(expected), `expected ${expected}; got: ` + JSON.stringify(codes))
			}
			// Every code is `area/condition` — the shape consumers pattern-match on.
			for (const code of codes) {
				assert(/^[a-z0-9-]+\/[a-z0-9-]+$/.test(code), 'malformed diagnostic code: ' + code)
			}
		},
	},
	{
		name: 'every published subpath can install the handler, and they are the same one',
		fn: async () => {
			// A consumer of `ts-pptx/read` alone gets warnings from the read path -- a chart point
			// cache out of range, a picture whose relationship does not resolve -- and until the
			// diagnostic surface was republished the way the error taxonomy already was, there was
			// no supported way to intercept them: the handler was exported only by the three
			// authoring entries. Importing it from `.` did happen to work, because bundling puts
			// `diagnostics.js` in a shared chunk, but that is an artifact of chunking rather than
			// a promise -- and it pulls the whole write path in for a three-line function.
			//
			// Identity, not just presence: one installed handler must serve every subpath, which
			// is only true while they all resolve to the one module.
			const base = await import('../../../dist/node.js')
			for (const entry of ['read', 'measure', 'script', 'inspect', 'html', 'math', 'zip']) {
				const mod = await import(`../../../dist/${entry}.js`)
				assertEqual(typeof mod.setDiagnosticHandler, 'function', `ts-pptx/${entry} must publish setDiagnosticHandler`)
				assertEqual(typeof mod.resetDiagnosticState, 'function', `ts-pptx/${entry} must publish resetDiagnosticState`)
				assert(
					mod.setDiagnosticHandler === base.setDiagnosticHandler,
					`ts-pptx/${entry} must publish the SAME handler installer as the main entry`
				)
			}
		},
	},
])
