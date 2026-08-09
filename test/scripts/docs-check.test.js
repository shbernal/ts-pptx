// The docs gate's link rules.
//
// These decide what VitePress will and will not resolve, and getting them wrong is not
// visible locally: a relative link to a repo file *outside* `docs/` points at a file that
// exists on disk, so nothing looks broken until `vitepress build` fails on the dead link.
// That is not hypothetical — the gate used to skip those links as "not ours to check", and
// one of them stalled the Pages deploy for three days. These pin both edges: the escaping
// link is an error, and the asset links that are genuinely not routes stay quiet.

import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterAll, beforeAll, describe, expect, test } from 'vitest'
import { checkLinks } from '../../scripts/docs-check.mjs'

/** A throwaway repo shaped like this one: a `docs/` tree with sibling files outside it. */
let root = ''
let docsDir = ''
const routes = new Set(['/', '/guide', '/guide/start'])

beforeAll(() => {
	root = mkdtempSync(path.join(os.tmpdir(), 'docs-check-'))
	docsDir = path.join(root, 'docs')
	mkdirSync(path.join(docsDir, 'guide'), { recursive: true })
	mkdirSync(path.join(root, 'test', 'read', 'fixtures'), { recursive: true })
	writeFileSync(path.join(root, 'test', 'read', 'fixtures', 'README.md'), '# fixtures\n')
	writeFileSync(path.join(docsDir, 'guide', 'start.md'), '# start\n')
	// A sibling of `docs/` whose name shares its leading characters — a prefix test would
	// wrongly call this one "inside the docs tree".
	mkdirSync(path.join(root, 'docs-extra'), { recursive: true })
	writeFileSync(path.join(root, 'docs-extra', 'notes.md'), '# notes\n')
})

afterAll(() => rmSync(root, { force: true, recursive: true }))

/** The errors reported for a page whose body is `markdown`. */
function check(markdown) {
	writeFileSync(path.join(docsDir, 'page.md'), markdown)
	return checkLinks(docsDir, 'page.md', routes)
}

describe('links that leave the docs tree', () => {
	test('a relative link to a repo file outside docs/ is an error, even though the file exists', () => {
		const errors = check('See [the fixtures](../test/read/fixtures/README.md).')
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/points outside docs\//)
	})

	test('a sibling directory sharing the `docs` prefix is still outside', () => {
		expect(check('See [notes](../docs-extra/notes.md).')).toHaveLength(1)
	})

	test('the absolute URL is how those files are meant to be linked', () => {
		expect(check('See [the fixtures](https://github.com/o/r/blob/master/test/read/fixtures/README.md).')).toEqual([])
	})
})

describe('links inside the docs tree', () => {
	test('a relative link to a real page passes', () => {
		expect(check('See [start](./guide/start.md).')).toEqual([])
	})

	test('a relative link to a missing page is an error', () => {
		const errors = check('See [gone](./guide/gone.md).')
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/broken relative link/)
	})

	test('an absolute site route must be one the site serves', () => {
		expect(check('See [start](/guide/start).')).toEqual([])
		expect(check('See [nope](/guide/nope)')).toHaveLength(1)
	})
})

describe('what is not a route, and so not this gate’s business', () => {
	test('a non-markdown asset, wherever it resolves', () => {
		expect(check('![shot](./images/shot.png) and [raw](../package.json)')).toEqual([])
	})

	test('an anchor on the page itself', () => {
		expect(check('See [below](#details).')).toEqual([])
	})

	test('a mail or protocol-relative target', () => {
		expect(check('[mail](mailto:x@example.com) and [cdn](//example.com/a.md)')).toEqual([])
	})
})
