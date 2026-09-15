// Frontmatter parsing and docs-tree walking, shared by the `docs:*` scripts.
//
// These were previously five Python scripts, three of which carried a byte-identical copy of the
// parser below — one of them with a comment explaining that the repo "keeps these small scripts
// self-contained rather than sharing a module". That trade stops paying once the scripts are in
// the same language as the rest of the tooling and can just import each other.
//
// The parser is deliberately NOT a general YAML parser, and this repo deliberately carries no
// YAML library to make it one — do not add one for this. It recognizes exactly the shape the docs
// kit writes — scalars, block lists, and inline `[a, b]` lists — and ignores anything else. A real
// parser would reject frontmatter this one tolerates, which would turn `docs:check` from a lint
// into a gate on YAML pedantry; the checks that matter are asserted explicitly in `docs-check.mjs`.

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import path from 'node:path'

import { ROOT } from './script-utils.mjs'

/** Subtrees that are documentation-adjacent but not part of the checked docs set. */
export const EXCLUDED_DIRS = new Set(['archive', 'changelog-archive', 'research'])

/** The `doc_type` values the frontmatter schema allows. */
export const ALLOWED_DOC_TYPES = [
	'architecture',
	'decision',
	'guide',
	'overview',
	'reference',
	'runbook',
	'troubleshooting',
]

/**
 * Drop one layer of matching surrounding quotes, if present.
 * @param {string} value
 * @returns {string}
 */
function stripQuotes(value) {
	const trimmed = value.trim()
	if (trimmed.length >= 2 && trimmed[0] === trimmed[trimmed.length - 1] && (trimmed[0] === "'" || trimmed[0] === '"')) {
		return trimmed.slice(1, -1)
	}
	return trimmed
}

/**
 * Non-empty, trimmed string values only — the shape every list field is consumed as.
 * @param {unknown} values
 * @returns {string[]}
 */
export function compactStrings(values) {
	if (!Array.isArray(values)) return []
	/** @type {string[]} */
	const out = []
	for (const value of values) {
		if (value === null || value === undefined) continue
		const text = String(value).trim()
		if (text) out.push(text)
	}
	return out
}

/**
 * Parse an inline `[a, b]` list, tolerating single quotes. Returns `[]` when it is not one.
 * @param {string} value
 * @returns {string[]}
 */
function parseInlineList(value) {
	let parsed
	try {
		parsed = JSON.parse(value.replaceAll("'", '"'))
	} catch {
		return []
	}
	return Array.isArray(parsed) ? compactStrings(parsed) : []
}

/**
 * Parse a markdown file's frontmatter block.
 * @param {string} filePath - absolute path to the markdown file
 * @returns {{ data: Record<string, unknown>, error: string | null, body: string }} `body` is the
 *   page after its frontmatter block, or the whole file when it has none
 */
export function parseFrontmatter(filePath) {
	const raw = readFileSync(filePath, 'utf8')
	if (!raw.startsWith('---\n') && !raw.startsWith('---\r\n'))
		return { data: {}, error: 'missing front matter', body: raw }

	const lines = raw.split(/\r?\n/)
	let endIndex = -1
	for (let i = 1; i < lines.length; i++) {
		const trimmed = (lines[i] ?? '').trim()
		if (trimmed === '---' || trimmed === '...') {
			endIndex = i
			break
		}
	}
	if (endIndex < 0) return { data: {}, error: 'unterminated front matter', body: raw }

	/** @type {Record<string, unknown>} */
	const data = {}
	/** @type {string | null} */
	let collecting = null
	for (const rawLine of lines.slice(1, endIndex)) {
		const line = rawLine.trim()
		if (!line) continue

		if (collecting && line.startsWith('- ')) {
			const bucket = data[collecting]
			if (Array.isArray(bucket)) bucket.push(stripQuotes(line.slice(2).trim()))
			continue
		}

		collecting = null
		const colon = line.indexOf(':')
		if (colon < 0) continue

		const key = line.slice(0, colon).trim()
		const value = line.slice(colon + 1).trim()
		if (!value) {
			data[key] = []
			collecting = key
		} else if (value.startsWith('[') && value.endsWith(']')) {
			data[key] = parseInlineList(value)
		} else {
			data[key] = stripQuotes(value)
		}
	}

	return { data, error: null, body: lines.slice(endIndex + 1).join('\n') }
}

/**
 * Every checked markdown page under `docsDir`, as repo-relative POSIX paths, sorted.
 * Dotted and excluded subtrees are skipped, as is `skipName` when given (the generated index
 * must not index itself).
 * @param {string} docsDir - absolute path to the docs directory
 * @param {string} [skipName] - a basename to omit
 * @returns {string[]} docs-relative POSIX paths
 */
export function walkDocs(docsDir, skipName) {
	/** @type {string[]} */
	const found = []
	/**
	 * @param {string} dir
	 * @param {string[]} relParts
	 * @returns {void}
	 */
	const walk = (dir, relParts) => {
		for (const entry of readdirSync(dir, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : 1))) {
			if (entry.name.startsWith('.') || EXCLUDED_DIRS.has(entry.name)) continue
			const abs = path.join(dir, entry.name)
			if (entry.isDirectory()) walk(abs, [...relParts, entry.name])
			else if (entry.isFile() && entry.name.endsWith('.md') && entry.name !== skipName) {
				found.push([...relParts, entry.name].join('/'))
			}
		}
	}
	walk(docsDir, [])
	return found.sort()
}

/**
 * The directories `docs.json` marks as repository-only under `repoOnly`. Their pages keep the
 * docs schema and are read on GitHub, but the site does not build them and the generated indexes
 * do not list them. `docs/.vitepress/config.mts` reads the same key, so no consumer names a
 * directory itself.
 * @param {unknown} config - the parsed `docs/docs.json`
 * @returns {string[]} docs-relative directory paths, without surrounding slashes
 */
export function repoOnlyDirs(config) {
	const value =
		config && typeof config === 'object' ? /** @type {Record<string, unknown>} */ (config).repoOnly : undefined
	return compactStrings(value).map((dir) => dir.replace(/^\/+|\/+$/g, ''))
}

/**
 * {@link repoOnlyDirs}, read from `docs.json` under `docsDir`; empty when there is no such file.
 * @param {string} docsDir - absolute path to the docs directory
 * @returns {string[]}
 */
export function readRepoOnlyDirs(docsDir) {
	const configPath = path.join(docsDir, 'docs.json')
	return existsSync(configPath) ? repoOnlyDirs(JSON.parse(readFileSync(configPath, 'utf8'))) : []
}

/**
 * Whether a docs-relative page path or page key sits under a repository-only directory.
 * @param {string} rel
 * @param {string[]} dirs - from {@link repoOnlyDirs}
 * @returns {boolean}
 */
export function isRepoOnly(rel, dirs) {
	return dirs.some((dir) => rel.startsWith(`${dir}/`))
}

/**
 * Resolve the docs directory, exiting with the script's own message when it is absent.
 * @param {string} label the calling script's name, for the error message
 * @returns {string}
 */
export function requireDocsDir(label) {
	const docsDir = path.join(ROOT, 'docs')
	let ok
	try {
		ok = statSync(docsDir).isDirectory()
	} catch {
		ok = false
	}
	if (!ok) {
		console.error(`${label}: missing docs directory at ${docsDir}.`)
		process.exit(1)
	}
	return docsDir
}

/**
 * The GitHub owner and repository `package.json` names, as a match whose groups 1 and 2 are them.
 * @param {string} docsDir - absolute path to the docs directory
 * @returns {RegExpMatchArray | null}
 */
function githubSlug(docsDir) {
	const pkg = JSON.parse(readFileSync(path.join(docsDir, '..', 'package.json'), 'utf8'))
	return String(pkg.repository?.url ?? pkg.homepage ?? '').match(/github\.com\/([^/]+)\/([^/.#]+)/)
}

/**
 * The prefix a file on the default branch is browsed at on GitHub, which is where a
 * repository-only page is read and so how a served page has to link one.
 * @param {string} docsDir - absolute path to the docs directory
 * @returns {string | null} `https://github.com/<owner>/<repo>/blob/master/`
 */
export function githubBlobBase(docsDir) {
	const slug = githubSlug(docsDir)
	return slug ? `https://github.com/${slug[1]}/${slug[2]}/blob/master/` : null
}

/**
 * The URL prefix the published site actually answers on, derived rather than restated: the
 * GitHub Pages host comes from `repository`, the path from the VitePress `base`. Writing the
 * value out a second time is what let `llms.txt` drift to a host that never existed, which is
 * why the llms generator and `docs-check.mjs` both read it from here. `DOCS_BASE_URL` overrides it.
 * @param {string} docsDir - absolute path to the docs directory
 * @returns {{base: string | null, errors: string[]}}
 */
export function canonicalBase(docsDir) {
	if (process.env.DOCS_BASE_URL) return { base: process.env.DOCS_BASE_URL.replace(/\/?$/, '/'), errors: [] }

	const slug = githubSlug(docsDir)
	if (!slug) return { base: null, errors: ['package.json: cannot derive the GitHub Pages host from `repository`'] }

	const config = readFileSync(path.join(docsDir, '.vitepress', 'config.mts'), 'utf8')
	const configured = config.match(/^\s*base:.*?'([^']+)'/m)
	if (!configured) return { base: null, errors: ['docs/.vitepress/config.mts: cannot read the `base` option'] }

	// An unreadable `base` collapses to `/`, which the slug check below then rejects loudly.
	const base = (process.env.VITEPRESS_BASE ?? configured[1] ?? '').replace(/\/?$/, '/')
	const errors = []
	// A project Pages site is served under /<repo>/, so a `base` that disagrees means the whole
	// site 404s no matter how well-formed the routes underneath it are.
	if (!process.env.VITEPRESS_BASE && base !== `/${slug[2]}/`) {
		errors.push(`docs/.vitepress/config.mts: \`base\` is \`${base}\`, but Pages serves this repo at \`/${slug[2]}/\``)
	}
	return { base: `https://${slug[1]}.github.io${base}`, errors }
}

/**
 * The route VitePress serves a docs page at, relative to the site base.
 *
 * VitePress builds with `cleanUrls`, which emits `tables.html` rather than `tables/index.html`.
 * GitHub Pages serves that as `/tables` and 404s on `/tables/`, so only a directory index may
 * carry the trailing slash. A README is not one of those: VitePress leaves it at `README.html`.
 * @param {string} rel - docs-relative POSIX path of the page
 * @returns {string} `''` for the site index, `dir/` for a directory index, the bare path otherwise
 */
export function routeForPage(rel) {
	const withoutExt = rel.replace(/\.md$/, '')
	if (withoutExt === 'index') return ''
	if (withoutExt.endsWith('/index')) return withoutExt.slice(0, withoutExt.lastIndexOf('/') + 1)
	return withoutExt
}

/**
 * The built file that serves a route, relative to the site's output directory: the inverse of
 * {@link routeForPage}, and the mapping the server applies under `cleanUrls`.
 * @param {string} route - a route relative to the site base
 * @returns {string}
 */
export function pageForRoute(route) {
	if (route === '') return 'index.html'
	return route.endsWith('/') ? `${route}index.html` : `${route}.html`
}

// Built from code points, so the file holds no raw control or combining characters.
const CONTROL_CHARACTERS = new RegExp(`[${String.fromCodePoint(0)}-${String.fromCodePoint(0x1f)}]`, 'g')
const COMBINING_MARKS = new RegExp(`[${String.fromCodePoint(0x300)}-${String.fromCodePoint(0x36f)}]`, 'g')
const TYPOGRAPHIC_QUOTES = new RegExp(`[${String.fromCodePoint(0x201c, 0x201d, 0x2018, 0x2019)}]`, 'g')

/**
 * A heading's anchor as the site generates it: VitePress's default `slugify` from
 * `@mdit-vue/shared`. Whitespace, ASCII punctuation and typographic quotes collapse to one
 * hyphen, so "`check:core` runs" becomes `check-core-runs` and `element_` becomes `element`.
 * @param {string} text the heading's plain text
 * @returns {string}
 */
export function siteHeadingSlug(text) {
	return text
		.normalize('NFKD')
		.replace(COMBINING_MARKS, '')
		.replace(CONTROL_CHARACTERS, '')
		.replace(TYPOGRAPHIC_QUOTES, '-')
		.replace(/[\s~`!@#$%^&*()\-_+=[\]{}|\\;:"'<>,.?/]+/g, '-')
		.replace(/-{2,}/g, '-')
		.replace(/^-+|-+$/g, '')
		.replace(/^(\d)/, '_$1')
		.toLowerCase()
}

/**
 * A heading's anchor as GitHub generates it, which is where a repository-only page is read.
 * Punctuation is dropped rather than turned into a hyphen, so "`check:core` runs" becomes
 * `checkcore-runs`, and each space is its own hyphen.
 * @param {string} text the heading's plain text
 * @returns {string}
 */
export function githubHeadingSlug(text) {
	return text
		.toLowerCase()
		.replace(/[^\p{L}\p{M}\p{N}\p{Pc} -]/gu, '')
		.replace(/ /g, '-')
}
