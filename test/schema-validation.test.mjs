import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { beforeAll, describe, test } from 'vitest'

const require = createRequire(import.meta.url)
const testDir = path.dirname(fileURLToPath(import.meta.url))
const { isInstalled, VALIDATOR } = require('./validator')
const cases = require(path.join(testDir, 'schema.test.js'))

describe('PptxGenJS schema validation fixtures', () => {
	beforeAll(() => {
		if (!isInstalled()) {
			throw new Error('OOXMLValidatorCLI not installed at ' + VALIDATOR + '\nRun: ./tools/ooxml-validator/install.sh')
		}
	})

	for (const fixture of cases) {
		test(fixture.name, async () => {
			await fixture.fn()
		})
	}
})
