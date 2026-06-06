import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { describe, test } from 'vitest'

const testDir = path.dirname(fileURLToPath(import.meta.url))
const files = fs
	.readdirSync(testDir)
	.filter((file) => /^bug-\d+\.test\.js$/.test(file))
	.sort()
const fixtures = await Promise.all(
	files.map(async (file) => ({
		file,
		cases: (await import(pathToFileURL(path.join(testDir, file)).href)).default,
	}))
)

describe('PptxGenJS regression fixtures', () => {
	for (const { file, cases } of fixtures) {
		describe(file, () => {
			for (const fixture of cases) {
				test(fixture.name, async () => {
					await fixture.fn()
				})
			}
		})
	}
})
