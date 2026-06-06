import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'

const require = createRequire(import.meta.url)
const testDir = path.dirname(fileURLToPath(import.meta.url))
const files = fs
	.readdirSync(testDir)
	.filter((file) => /^bug-\d+\.test\.js$/.test(file))
	.sort()

describe('PptxGenJS regression fixtures', () => {
	for (const file of files) {
		const cases = require(path.join(testDir, file))
		describe(file, () => {
			for (const fixture of cases) {
				test(fixture.name, async () => {
					await fixture.fn()
				})
			}
		})
	}
})
