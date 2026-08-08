import { defineRegressionSuite, TsPptx, assert, assertEqual } from '../../helpers.js'
import { InvalidOptionError } from '../../../dist/node.js'

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
	{
		// `addModel3d()`'s camera and scale end up as raw numbers in `am3d` attributes, so a `NaN`
		// or an out-of-range value would be emitted verbatim -- schema-valid-looking markup that
		// renders as nothing, in a subtree the schema validator does not check (it never descends
		// into an `mc:Choice`). Reject rather than coerce.
		name: 'addModel3d() rejects an unusable payload, camera, or scale instead of emitting NaN',
		fn: () => {
			const glb = 'Z2xURgIAAAA='
			assertRejects(() => slide().addModel3d(/** @type {never} */ ({})), 'model3d/missing-source')
			// Camera positions: every component of every vector is checked.
			for (const bad of [Number.NaN, Number.POSITIVE_INFINITY]) {
				assertRejects(
					() => slide().addModel3d({ data: glb, camera: { pos: { x: bad, y: 0, z: 1 } } }),
					'model3d/invalid-camera'
				)
				assertRejects(
					() => slide().addModel3d({ data: glb, camera: { lookAt: { x: 0, y: bad, z: 0 } } }),
					'model3d/invalid-camera'
				)
				assertRejects(
					() => slide().addModel3d({ data: glb, camera: { up: { x: 0, y: 0, z: bad } } }),
					'model3d/invalid-camera'
				)
			}
			// fov is an angle, not a scale factor: 0 and 180 are degenerate, not merely extreme.
			for (const fov of [Number.NaN, 0, -45, 180, 200]) {
				assertRejects(() => slide().addModel3d({ data: glb, camera: { fov } }), 'model3d/invalid-fov')
			}
			// A non-positive metres-per-unit inverts or collapses the whole scene.
			for (const scale of [Number.NaN, 0, -1, Number.POSITIVE_INFINITY]) {
				assertRejects(() => slide().addModel3d({ data: glb, meterPerModelUnit: scale }), 'model3d/invalid-scale')
			}
		},
	},
	{
		// The counterpart: values that are unusual but meaningful must pass. A camera below the
		// model (negative y) and a very small scale are both ordinary for a large model.
		name: 'addModel3d() accepts an unusual but valid camera and scale',
		fn: () => {
			const { error } = caughtQuietly(() =>
				slide().addModel3d({
					data: 'Z2xURgIAAAA=',
					preview: { data: 'image/png;base64,iVBORw0KGgo=' },
					camera: { pos: { x: -2, y: -3.5, z: 0.25 }, lookAt: { x: 0, y: 0, z: 0 }, fov: 179.9 },
					meterPerModelUnit: 1 / 10000,
				})
			)
			assert(!error, `an unusual camera is still a camera; got: ${String(error)}`)
		},
	},
])
