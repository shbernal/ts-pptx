// Browser-path coverage for readFileAsBytes in src/zip.ts: when `node:fs` is
// unavailable (a browser bundle), the lazy `import('node:fs/promises')` fails
// and a string path must throw the clear "filesystem access requires Node"
// error rather than an opaque failure.
//
// This branch cannot be reached in a real Node run — the import always
// succeeds — so we force the dynamic import to reject with vi.doMock. It lives
// in its own file so mocking the fs module can't leak into the other read
// tests. Because the mock must be active before the module under test loads,
// dist/zip.js is imported dynamically *after* the mock is registered.

import { describe, test, vi, beforeEach, afterEach } from 'vitest'
import { assert } from '../helpers.js'

describe('readFileAsBytes when node:fs is unavailable (browser build)', () => {
	beforeEach(() => {
		vi.resetModules()
		// Make the lazy `import('node:fs/promises')` reject, standing in for a
		// browser environment where the Node builtin does not resolve.
		vi.doMock('node:fs/promises', () => {
			throw new Error('simulated: node:fs/promises unavailable')
		})
	})

	afterEach(() => {
		vi.doUnmock('node:fs/promises')
		vi.resetModules()
	})

	test('a string path throws a Node-required error, not the opaque zip error', async () => {
		const { readZip } = await import('../../dist/zip.js')
		let error = null
		try {
			await readZip('/any/deck.pptx')
		} catch (err) {
			error = err
		}
		assert(error, 'a string path throws when the filesystem is unavailable')
		assert(
			error.message.includes('filesystem access requires Node'),
			`error explains the missing filesystem; got: ${error.message}`
		)
		assert(error.message.includes('/any/deck.pptx'), `error names the path; got: ${error.message}`)
		assert(
			!error.message.includes('Not a valid ZIP archive'),
			'a missing filesystem is not misreported as a corrupt archive'
		)
	})
})
