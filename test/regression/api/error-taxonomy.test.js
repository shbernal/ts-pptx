import { defineRegressionSuite, assert, assertEqual } from '../../helpers.js'
import TsPptx, {
	TsPptxError,
	InvalidOptionError,
	UnsupportedFeatureError,
	PackageReadError,
} from '../../../dist/node.js'
import { readZip } from '../../../dist/zip.js'
import { PackageReadError as PackageReadErrorFromZip, TsPptxError as TsPptxErrorFromZip } from '../../../dist/zip.js'
import { TsPptxError as TsPptxErrorFromRead } from '../../../dist/read.js'

// The error taxonomy: every failure the library raises is a `TsPptxError` carrying a stable
// `code`, so a consumer can classify it without matching on message substrings.
//
// What is pinned here is the CONTRACT, not the wording. The class and the `code` are API; the
// `message` is explicitly free to change in any release, so nothing below asserts on message text.

/** Run `fn`, returning whatever it threw (or `null` if it did not throw). */
async function caught(fn) {
	try {
		await fn()
		return null
	} catch (err) {
		return err
	}
}

defineRegressionSuite('Error taxonomy', [
	{
		name: 'an invalid option throws an InvalidOptionError with a stable code',
		fn: async () => {
			const err = await caught(() => {
				const pres = new TsPptx()
				pres.layout = 'NOT-A-LAYOUT'
			})

			assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
			assertEqual(err.code, 'layout/unknown', 'the code identifies the condition')
			assertEqual(err.name, 'InvalidOptionError', 'name is the subclass, so stack traces label it correctly')
		},
	},
	{
		name: 'every library error is a TsPptxError and an Error',
		fn: async () => {
			// `typeface` is required at the type level, so only an untyped caller reaches the runtime
			// guard — which is exactly the caller the taxonomy exists for.
			const err = await caught(() => new TsPptx().embedFont(/** @type {never} */ ({ data: 'AAAA' })))

			// A consumer catching broadly must still be able to catch this, and a consumer catching
			// only ts-pptx failures must be able to say so in one `instanceof`.
			assert(err instanceof Error, 'every library error remains an Error')
			assert(err instanceof TsPptxError, 'every library error is a TsPptxError')
			assert(err instanceof InvalidOptionError, 'a missing required option is an invalid option')
			assertEqual(err.code, 'font/missing-typeface')
		},
	},
	{
		name: 'unreadable input bytes throw a PackageReadError, not an InvalidOptionError',
		fn: async () => {
			// The distinction is the point of the taxonomy: bad *input bytes* are a different
			// failure from a bad *option*, and a consumer routes them differently.
			const err = await caught(() => readZip(new Uint8Array([1, 2, 3, 4])))

			assert(err instanceof PackageReadError, 'expected a PackageReadError; got: ' + String(err))
			assertEqual(err.code, 'zip/not-a-zip-archive')
			assert(!(err instanceof InvalidOptionError), 'a malformed package is not an invalid option')
		},
	},
	{
		name: 'an unsupported input type is an invalid option, not a package-read failure',
		fn: async () => {
			const err = await caught(() => readZip(/** @type {never} */ (42)))

			assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
			assertEqual(err.code, 'zip/unsupported-input')
		},
	},
	{
		name: 'the classes are identical across every entry point',
		fn: async () => {
			// Re-exported from ten entries, but resolved from one shared module — otherwise an
			// `instanceof` in a consumer that imports `ts-pptx/read` would silently fail against an
			// error thrown through `ts-pptx`.
			assert(TsPptxErrorFromZip === TsPptxError, 'ts-pptx/zip exports the same TsPptxError')
			assert(TsPptxErrorFromRead === TsPptxError, 'ts-pptx/read exports the same TsPptxError')
			assert(PackageReadErrorFromZip === PackageReadError, 'ts-pptx/zip exports the same PackageReadError')

			const err = await caught(() => readZip(new Uint8Array([1, 2, 3, 4])))
			assert(err instanceof PackageReadErrorFromZip, 'the thrown error matches the subpath-imported class')
		},
	},
	{
		name: 'a cause is preserved so the underlying failure is not lost',
		fn: async () => {
			const err = await caught(() => readZip(new Uint8Array([1, 2, 3, 4])))

			// Wrapping must not discard what fflate reported — the taxonomy adds classification on
			// top of the original failure rather than replacing it.
			assert(err.cause !== undefined, 'the originating error is preserved as `cause`')
		},
	},
	{
		name: 'UnsupportedFeatureError is exported and distinct from the other classes',
		fn: () => {
			// No Node-reachable site trips this synchronously (the sites are a missing optional peer
			// and a browser-only filesystem gap), so the contract pinned here is the shape.
			const err = new UnsupportedFeatureError('math/missing-optional-peer', 'x')
			assert(err instanceof TsPptxError, 'it is a TsPptxError')
			assert(!(err instanceof InvalidOptionError), 'it is not an InvalidOptionError')
			assertEqual(err.name, 'UnsupportedFeatureError')
			assertEqual(err.code, 'math/missing-optional-peer')
		},
	},
	{
		name: 'detail is carried when a site supplies it, and absent otherwise',
		fn: () => {
			const bare = new InvalidOptionError('coord/non-finite', 'x')
			assertEqual(bare.detail, undefined, 'no detail key when a site gives none')

			const detailed = new InvalidOptionError('coord/non-finite', 'x', { detail: { value: NaN } })
			assert(Number.isNaN(detailed.detail.value), 'structured detail survives the constructor')
		},
	},
])
