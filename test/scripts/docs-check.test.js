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
import { checkDocsJson, checkLinks } from '../../scripts/docs-check.mjs'

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

// A repository-only page is read on GitHub and never built, so the two kinds of page resolve
// links in different places: a served page's relative link to one is a dead link in the site,
// and a repository-only page's site route is a dead link on GitHub.
describe('the repository-only tree', () => {
	const repoOnly = { repoOnly: ['contributing'], blobBase: 'https://github.com/o/r/blob/master/' }

	beforeAll(() => {
		mkdirSync(path.join(docsDir, 'contributing'), { recursive: true })
		writeFileSync(path.join(docsDir, 'contributing', 'testing.md'), '# testing\n')
	})

	/** The errors reported for the page at docs-relative `rel` whose body is `markdown`. */
	function checkAt(rel, markdown) {
		writeFileSync(path.join(docsDir, rel), markdown)
		return checkLinks(docsDir, rel, routes, repoOnly)
	}

	test('a served page linking relatively to a repository-only page is told the GitHub URL', () => {
		const errors = checkAt('page.md', 'See [testing](./contributing/testing.md#suites).')
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/repository-only/)
		expect(errors[0]).toContain('https://github.com/o/r/blob/master/docs/contributing/testing.md#suites')
	})

	test('the GitHub URL is how a served page links one', () => {
		expect(
			checkAt('page.md', 'See [testing](https://github.com/o/r/blob/master/docs/contributing/testing.md).')
		).toEqual([])
	})

	test('a repository-only page links relatively to any page, served or not', () => {
		expect(checkAt('contributing/page.md', 'See [start](../guide/start.md) and [testing](./testing.md).')).toEqual([])
	})

	test('a repository-only page linking a site route is an error, since GitHub has no such route', () => {
		const errors = checkAt('contributing/page.md', 'See [start](/guide/start).')
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/link the page relatively/)
	})
})

describe('navigation and the repository-only tree', () => {
	/** The navigation errors for `pages` under a docs.json holding `config`. */
	function checkNav(config, pages) {
		writeFileSync(path.join(docsDir, 'docs.json'), JSON.stringify(config))
		return checkDocsJson(docsDir, pages)
	}

	const pages = ['index.md', 'contributing/testing.md']

	test('a repository-only page needs no navigation group', () => {
		expect(checkNav({ repoOnly: ['contributing'], navigation: [{ group: 'Start', pages: ['index'] }] }, pages)).toEqual(
			[]
		)
	})

	test('without the repoOnly entry, the same page is an orphan', () => {
		const errors = checkNav({ navigation: [{ group: 'Start', pages: ['index'] }] }, pages)
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/in no navigation group/)
	})

	test('a navigation entry naming a repository-only page is an error, because the site would 404', () => {
		const errors = checkNav(
			{ repoOnly: ['contributing'], navigation: [{ group: 'Start', pages: ['index', 'contributing/testing'] }] },
			pages
		)
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/repository-only/)
	})

	test('repoOnly must be a list of directory names', () => {
		const errors = checkNav({ repoOnly: 'contributing', navigation: [{ group: 'Start', pages: ['index'] }] }, [
			'index.md',
		])
		expect(errors).toHaveLength(1)
		expect(errors[0]).toMatch(/`repoOnly` must be a list/)
	})
})
