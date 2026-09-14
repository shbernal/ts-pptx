import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import { afterAll, beforeAll, describe, expect, test } from 'vitest'
import { Presentation } from '../../../dist/read.js'
import { validateBuf, validatorInstalled } from '../../validator.js'

/**
 * The complete program on "Your first deck" runs, and writes a deck that opens.
 *
 * It is the page a new reader copies from, so a renamed call or a retyped option there breaks the
 * example every later page assumes already worked. The program is read from the page itself, from
 * the fence after the marker comment, not from a copy kept here that could drift from it. Its one
 * `"pptx-ts"` import is pointed at this checkout's `dist/node.js`, and Node runs it with its own
 * TypeScript type stripping, the way a reader would run `node deck.ts`.
 */

const REPO = fileURLToPath(new URL('../../../', import.meta.url))
const PAGE = path.join(REPO, 'docs', 'getting-started', 'first-deck.md')
const MARKER = '<!-- first-deck:program -->'

/** The source inside the `ts` fence that follows the marker comment. */
function programFromPage() {
	const markdown = readFileSync(PAGE, 'utf8')
	const at = markdown.indexOf(MARKER)
	if (at === -1) throw new Error(`${PAGE} has no ${MARKER} comment`)
	const fence = /^```ts\r?\n([\s\S]*?)^```/m.exec(markdown.slice(at + MARKER.length))
	if (!fence) throw new Error(`no ts fence follows ${MARKER} in ${PAGE}`)
	return fence[1]
}

let dir = ''
/** @type {Buffer} */
let deck

beforeAll(() => {
	dir = mkdtempSync(path.join(os.tmpdir(), 'first-deck-'))
	const source = programFromPage()
	// Exactly one import to repoint. A second would be a dependency the page never tells the
	// reader to install, and none would mean the program no longer uses the package at all.
	const imports = source.match(/from "pptx-ts"/g) ?? []
	if (imports.length !== 1) {
		throw new Error(`the program on ${PAGE} should import "pptx-ts" once, not ${imports.length} times`)
	}
	const entry = pathToFileURL(path.join(REPO, 'dist', 'node.js')).href
	writeFileSync(path.join(dir, 'deck.mts'), source.replace('from "pptx-ts"', `from ${JSON.stringify(entry)}`))
	try {
		execFileSync(process.execPath, ['deck.mts'], { cwd: dir, stdio: 'pipe' })
	} catch (error) {
		// The program's own stderr is the whole diagnosis (a renamed method, a type Node could not
		// strip); the default message says only that the command failed.
		throw new Error(`the program on ${PAGE} failed:\n${String(error.stderr)}`, { cause: error })
	}
	deck = readFileSync(path.join(dir, 'quarterly-summary.pptx'))
})

afterAll(() => rmSync(dir, { force: true, recursive: true }))

describe('Your first deck', () => {
	test('the program writes a deck that opens, with the four slides the page builds', async () => {
		const pres = await Presentation.load(deck)
		expect(pres.slides).toHaveLength(4)
	})

	test.skipIf(!validatorInstalled)('the deck it writes is schema-valid', async () => {
		expect(await validateBuf(deck)).toEqual([])
	})
})
