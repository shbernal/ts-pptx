#!/usr/bin/env node
/**
 * Emit load() -> save() round-trip output for each read fixture so the saved
 * decks can be opened in PowerPoint to confirm there is no repair prompt (the
 * manual check in test/read/fixtures/README.md).
 *
 * Output goes to .tmp/roundtrip/ (gitignored) by default; override with the
 * first CLI arg or PPTXGENJS_READ_EMIT_DIR. Assumes a current build — the
 * test:read:emit script runs `pnpm run build` first.
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT } from './script-utils.mjs'

const fixturesDir = path.join(ROOT, 'test', 'read', 'fixtures')
const outDir = process.argv[2] || process.env.PPTXGENJS_READ_EMIT_DIR || path.join(ROOT, '.tmp', 'roundtrip')

const readEntry = path.join(ROOT, 'dist', 'read.js')
try {
	await fs.access(readEntry)
} catch {
	console.error(
		`Missing ${path.relative(ROOT, readEntry)}. Run \`pnpm run build\` first (or use \`pnpm run test:read:emit\`).`
	)
	process.exit(1)
}
const { OpcPackage } = await import(pathToFileURL(readEntry).href)

const entries = await fs.readdir(fixturesDir)
const fixtures = entries.filter((name) => name.endsWith('.pptx')).sort()
if (fixtures.length === 0) {
	console.error(`No .pptx fixtures found in ${path.relative(ROOT, fixturesDir)}`)
	process.exit(1)
}

await fs.mkdir(outDir, { recursive: true })
for (const fixture of fixtures) {
	const input = await fs.readFile(path.join(fixturesDir, fixture))
	const pkg = await OpcPackage.load(input)
	const output = await pkg.save()
	const outName = fixture.replace(/\.pptx$/, '.roundtrip.pptx')
	const outPath = path.join(outDir, outName)
	await fs.writeFile(outPath, output)
	console.log(`${fixture}: ${input.length} -> ${output.length} bytes  ${path.relative(ROOT, outPath)}`)
}

console.log(`\nOpen the files in ${path.relative(ROOT, outDir)}/ in PowerPoint and confirm no repair prompt.`)
console.log('Record the result in test/read/fixtures/README.md.')
