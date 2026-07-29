import { defineRegressionSuite, TsPptx, assert, assertEqual } from '../helpers.js'
import { InvalidOptionError } from '../../dist/node.js'

// The reality-checks in `gen/define/` that used to write straight to `console.error` /
// `console.log` and then carry on, so the caller got an unroutable line on stderr and a deck
// quietly missing what they asked for. Each one now either throws (nothing sane to draw) or
// emits a routed diagnostic (there is a fallback) -- never both, and never the console.
//
// What decides which: `addImage()` with an unusable source has no image to place, and
// `addTable()` given a row that is not an array has no cells to lay out, so both reject. The
// picture-bullet check is the counter-example and lives in text-definition.test.js: the run
// emitter falls back to a default glyph, so refusing the rel is recoverable and warns instead.

/** Call `fn` with `console.log`/`console.error` captured; returns `{ error, noise }`. */
function caughtQuietly(fn) {
	const origLog = console.log
	const origError = console.error
	const noise = []
	console.log = (...args) => noise.push(args.map(String).join(' '))
	console.error = (...args) => noise.push(args.map(String).join(' '))
	try {
		fn()
		return { error: null, noise }
	} catch (err) {
		return { error: err, noise }
	} finally {
		console.log = origLog
		console.error = origError
	}
}

/** Assert `fn` throws an `InvalidOptionError` carrying `code`, and says nothing on the console. */
function assertRejects(fn, code) {
	const { error, noise } = caughtQuietly(fn)
	assert(error instanceof InvalidOptionError, `expected an InvalidOptionError; got: ${String(error)}`)
	assertEqual(error.code, code, 'the code identifies the condition')
	assertEqual(noise.join('\n'), '', 'the condition must reach the caller only through the throw')
}

/** A one-slide deck to define objects on. */
const slide = () => new TsPptx().addSlide()

defineRegressionSuite('Definition reality-checks', [
	{
		// The four `addImage()` source checks, in the order the chain tests them. The two
		// `not-a-string` arms are only reachable from an untyped caller -- `ImageProps` types both
		// as `string` -- which is exactly the caller a runtime check exists for.
		name: 'addImage() rejects an unusable source instead of dropping the image',
		fn: () => {
			assertRejects(() => slide().addImage(/** @type {never} */ ({})), 'image/missing-source')
			assertRejects(() => slide().addImage(/** @type {never} */ ({ path: 42 })), 'image/path-not-a-string')
			assertRejects(() => slide().addImage(/** @type {never} */ ({ data: 42 })), 'image/data-not-a-string')
			assertRejects(() => slide().addImage({ data: 'iVBORw0KGgoAAAA==' }), 'image/missing-base64-header')
		},
	},
	{
		// `addTable()` already rejected a non-nested row 0 up front; every later row only logged and
		// pushed an empty row, so `[['a'], 'b']` built a deck with a blank second row. Same condition,
		// same code, wherever in the array it sits.
		name: 'addTable() rejects a row that is not an array of cells, at any index',
		fn: () => {
			assertRejects(() => slide().addTable(/** @type {never} */ (['a'])), 'table/rows-not-nested')
			assertRejects(() => slide().addTable(/** @type {never} */ ([['a'], 'b'])), 'table/rows-not-nested')
			assertRejects(() => slide().addTable(/** @type {never} */ ([['a'], ['b'], 'c'])), 'table/rows-not-nested')
		},
	},
	{
		// The rejections must not fire on input that is merely unusual: an empty row is a row.
		name: 'a well-formed table with an empty row is accepted',
		fn: () => {
			const { error, noise } = caughtQuietly(() => slide().addTable([[{ text: 'a' }], [], [{ text: 'c' }]]))
			assert(!error, `an empty row is well-formed; got: ${String(error)}`)
			assertEqual(noise.join('\n'), '', 'and it is not worth a console line either')
		},
	},
])
