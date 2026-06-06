import { beforeAll, describe, test } from 'vitest'
import { isInstalled, VALIDATOR } from './validator.js'
import cases from './schema.test.js'

describe('PptxGenJS schema validation fixtures', () => {
	beforeAll(async () => {
		if (!(await isInstalled())) {
			throw new Error('OOXMLValidatorCLI not installed at ' + VALIDATOR + '\nRun: ./tools/ooxml-validator/install.sh')
		}
	})

	for (const fixture of cases) {
		test(fixture.name, async () => {
			await fixture.fn()
		})
	}
})
