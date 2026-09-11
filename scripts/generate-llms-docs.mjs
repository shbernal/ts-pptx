#!/usr/bin/env node
// Generate `docs/public/llms.txt` (a linked list of the docs pages) and `llms-full.txt` (every
// page's body) from the docs tree.
//
// Runs in `docs:prepare` after `docs:index`, because the flat index that script writes is one of
// the pages listed here. It used to run first and saw the index only when an earlier run had left
// one behind, so a fresh checkout, which is what CI builds, published an `llms.txt` without it and
// a second local run published one with it.
//
// The page set, the frontmatter, the site's base URL and the page routes all come from
// `docs-frontmatter.mjs`, where `docs-check.mjs` reads them too, so the published list and the
// check that verifies its URLs cannot disagree about what a page is or where it is served.

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

import { canonicalBase, parseFrontmatter, routeForPage, walkDocs } from './docs-frontmatter.mjs'
import { ROOT, parseCliOrExit, repoRel } from './script-utils.mjs'

// No flags, but `--help` still has to answer and `--bogus` still has to report itself in one
// line -- and both have to happen BEFORE the generator writes anything.
parseCliOrExit(process.argv.slice(2), {
	usage: `Generate docs/public/llms.txt and llms-full.txt from the docs tree.

  pnpm run docs:llms

Environment:
  DOCS_BASE_URL   site base URL for the emitted links

Options:
  -h, --help      show this message`,
	options: {},
})

const docsDir = path.join(ROOT, 'docs')
const publicDir = path.join(docsDir, 'public')
const docsConfig = JSON.parse(readFileSync(path.join(docsDir, 'docs.json'), 'utf8'))
const pkg = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8'))

/**
 * The published site's base URL. Exits when it cannot be derived, rather than advertising links
 * under a guess.
 * @returns {string}
 */
function siteBase() {
	const { base, errors } = canonicalBase(docsDir)
	if (base && errors.length === 0) return base
	for (const error of errors) console.error(`docs:llms: ${error}`)
	process.exit(1)
}

const baseUrl = siteBase()

/**
 * Whether VitePress serves a docs page at a route of its own. A README beside an `index.md` is
 * superseded by it.
 * @param {string} rel - docs-relative POSIX path of the page
 * @returns {boolean}
 */
function isServedPage(rel) {
	if (path.posix.basename(rel) !== 'README.md') return true
	return !existsSync(path.join(docsDir, path.posix.dirname(rel), 'index.md'))
}

/**
 * @param {string} body
 * @param {string} fallback
 * @returns {string}
 */
function titleFromBody(body, fallback) {
	const heading = body.match(/^#\s+(.+)$/m)
	return heading?.[1] ? heading[1].trim() : fallback
}

/**
 * A frontmatter value, when it is a single string.
 * @param {unknown} value
 * @returns {string | undefined}
 */
function scalar(value) {
	return typeof value === 'string' ? value : undefined
}

/**
 * One page as the generated indexes consume it.
 * @typedef {{body: string, rel: string, summary: string, title: string, url: string}} PageRecord
 */

/**
 * @param {string} rel - docs-relative POSIX path of the page
 * @returns {PageRecord}
 */
function pageRecord(rel) {
	const { data, body: rawBody } = parseFrontmatter(path.join(docsDir, rel))
	const body = rawBody.trimStart()
	return {
		body,
		rel,
		summary: scalar(data.summary) ?? '',
		title: scalar(data.title) ?? titleFromBody(body, rel.replace(/\.md$/, '')),
		url: new URL(routeForPage(rel), baseUrl).toString(),
	}
}

/**
 * @param {PageRecord[]} records
 * @returns {PageRecord[]}
 */
function docsNavigationOrder(records) {
	const byRouteKey = new Map(
		records.map(/** @param {PageRecord} record */ (record) => [record.rel.replace(/\.md$/, ''), record])
	)
	const ordered = []
	for (const group of docsConfig.navigation ?? []) {
		for (const page of group.pages ?? []) {
			const record = byRouteKey.get(page)
			if (record) ordered.push(record)
		}
	}
	for (const record of records) {
		if (!ordered.includes(record)) ordered.push(record)
	}
	return ordered
}

const records = docsNavigationOrder(walkDocs(docsDir).filter(isServedPage).map(pageRecord))

const llms = [
	`# ${docsConfig.name}`,
	'',
	`> ${docsConfig.description}`,
	'',
	`Package: ${pkg.name}`,
	`Version: ${pkg.version}`,
	'',
	'## Docs',
	'',
	...records.map((record) => `- [${record.title}](${record.url})${record.summary ? `: ${record.summary}` : ''}`),
	'',
].join('\n')

const llmsFull = [
	`# ${docsConfig.name}`,
	'',
	`> ${docsConfig.description}`,
	'',
	...records.flatMap((record) => [
		`## ${record.title}`,
		'',
		`URL: ${record.url}`,
		'',
		record.summary ? `${record.summary}\n` : '',
		record.body.trim(),
		'',
	]),
].join('\n')

mkdirSync(publicDir, { recursive: true })
writeFileSync(path.join(publicDir, 'llms.txt'), llms, 'utf8')
writeFileSync(path.join(publicDir, 'llms-full.txt'), llmsFull, 'utf8')

console.log(`generated ${repoRel(path.join(publicDir, 'llms.txt'))}`)
console.log(`generated ${repoRel(path.join(publicDir, 'llms-full.txt'))}`)
