#!/usr/bin/env node
/**
 * Emit load() -> save() round-trip output for each read fixture so the saved
 * decks can be opened in PowerPoint to confirm there is no repair prompt (the
 * manual check in test/read/fixtures/README.md).
 *
 * Output goes to .tmp/roundtrip/ (gitignored) by default; override with the
 * first CLI arg or TSPPTX_READ_EMIT_DIR. Assumes a current build — the
 * test:read:emit script ensures `dist/` is current first.
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { FIXTURES_DIR, ROOT, corpusDecks, parseCliOrExit, requireDist } from './script-utils.mjs'

const { positionals } = parseCliOrExit(process.argv.slice(2), {
	usage: `Emit load() -> save() output for each read fixture, for the manual PowerPoint check.

  pnpm run test:read:emit
  pnpm run test:read:emit -- <out-dir>

This is a GENERATOR, not a gate: it writes decks for a human to open in desktop
PowerPoint and confirm no repair prompt. Nothing here asserts on the output.

Arguments:
  <out-dir>   where to write (default .tmp/roundtrip, or $TSPPTX_READ_EMIT_DIR)

Options:
  -h, --help  show this message`,
	allowPositionals: true,
	options: {},
})

const outDir = positionals[0] || process.env.TSPPTX_READ_EMIT_DIR || path.join(ROOT, '.tmp', 'roundtrip')

const { OpcPackage } = await requireDist('read.js', 'test:read:emit')

const fixtures = await corpusDecks()

await fs.mkdir(outDir, { recursive: true })
for (const fixture of fixtures) {
	const input = await fs.readFile(path.join(FIXTURES_DIR, fixture))
	const pkg = await OpcPackage.load(input)
	const output = await pkg.save()
	const outName = fixture.replace(/\.pptx$/, '.roundtrip.pptx')
	const outPath = path.join(outDir, outName)
	await fs.writeFile(outPath, output)
	console.log(`${fixture}: ${input.length} -> ${output.length} bytes  ${path.relative(ROOT, outPath)}`)
}

console.log(`\nOpen the files in ${path.relative(ROOT, outDir)}/ in PowerPoint and confirm no repair prompt.`)
console.log('Record the result in test/read/fixtures/README.md.')
