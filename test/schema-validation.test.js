import { beforeAll, describe, test } from 'vitest'
import TsPptx from '../dist/node.js'
import { isInstalled, validateBuf, VALIDATOR } from './validator.js'
import cases from './schema-cases.js'

// The fixtures run concurrently (see `describe.concurrent` below), so each one
// spawns its own OOXMLValidatorCLI process. That binary is a .NET single-file
// app: the first invocation self-extracts its bundle into
// DOTNET_BUNDLE_EXTRACT_BASE_DIR, and several processes racing on a cold
// extract directory can collide. Validate one minimal deck serially first so
// the bundle is already extracted by the time the concurrent fixtures start.
describe.concurrent('TsPptx schema validation fixtures', () => {
	beforeAll(async () => {
		if (!(await isInstalled())) {
			throw new Error('OOXMLValidatorCLI not installed at ' + VALIDATOR + '\nRun: ./tools/ooxml-validator/install.sh')
		}
		const pres = new TsPptx()
		pres.addSlide()
		await validateBuf(/** @type {Uint8Array} */ (await pres.stream()))
	})

	for (const fixture of cases) {
		test(fixture.name, async () => {
			await fixture.fn()
		})
	}
})
