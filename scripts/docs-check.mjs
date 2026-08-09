#!/usr/bin/env node
// Validate the docs tree: frontmatter schema, balanced code fences, navigation entries that
// resolve to a real page, and internal links that resolve to a real route or file.
//
// Runs in `docs:check` and again inside `docs:build`, so a broken link fails the docs build
// rather than shipping a 404.
//
// With `--dist`, checks the built output instead: every URL in the generated llms.txt files
// must name a page VitePress actually emitted. That is a separate pass because it can only run
// *after* `vitepress build`, and the source checks above only need the markdown tree.

import { existsSync, readFileSync } from 'node:fs'
import path from 'node:path'

import { ALLOWED_DOC_TYPES, compactStrings, parseFrontmatter, requireDocsDir, walkDocs } from './docs-frontmatter.mjs'

const REQUIRED_FIELDS = ['doc-schema-version', 'doc_type', 'read_when', 'summary', 'title']
// Markdown inline links, excluding images (`!` prefix) and tolerating a trailing "title".
const MARKDOWN_LINK_RE = /(?<!!)\[[^\]]+\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g

/** Normalize a site route: strip query/hash and surrounding slashes, keeping a leading `/`. */
function normalizeRoute(route) {
	const bare = (route.split('#', 1)[0]?.split('?', 1)[0] ?? '').replace(/^\/+|\/+$/g, '')
	return bare ? `/${bare}` : '/'
}

/** Every route a docs page is reachable at (a page, plus `/` for index pages). */
function routesFor(relativePath) {
	const stem = relativePath.replace(/\.md$/, '')
	const routes = new Set([normalizeRoute(stem)])
	if (stem === 'index') routes.add('/')
	if (stem.endsWith('/index')) routes.add(normalizeRoute(stem.slice(0, -'/index'.length)))
	return routes
}

/** Every `pages:` entry anywhere in a docs.json navigation tree. */
function collectNavPages(value) {
	const pages = []
	if (Array.isArray(value)) {
		for (const item of value) pages.push(...collectNavPages(item))
	} else if (value && typeof value === 'object') {
		for (const [key, item] of Object.entries(value)) {
			if (key === 'pages' && Array.isArray(item)) {
				for (const page of item) {
					if (typeof page === 'string') pages.push(page)
					else pages.push(...collectNavPages(page))
				}
			} else {
				pages.push(...collectNavPages(item))
			}
		}
	}
	return pages
}

/** A link target that leaves the docs site entirely. */
function isExternal(target) {
	return /^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('//')
}

function checkFrontmatter(docsDir, rel) {
	const { data, error } = parseFrontmatter(path.join(docsDir, rel))
	if (error) return [`${rel}: ${error}`]

	const errors = []
	for (const field of REQUIRED_FIELDS) {
		if (!Object.hasOwn(data, field)) errors.push(`${rel}: missing frontmatter field \`${field}\``)
	}
	if (String(data['doc-schema-version'] ?? '').trim() !== '1') errors.push(`${rel}: \`doc-schema-version\` must be 1`)
	for (const field of ['title', 'summary', 'doc_type']) {
		if (!String(data[field] ?? '').trim()) errors.push(`${rel}: \`${field}\` must be non-empty`)
	}
	const docType = String(data.doc_type ?? '').trim()
	if (docType && !ALLOWED_DOC_TYPES.includes(docType)) {
		errors.push(`${rel}: \`doc_type\` must be one of: ${ALLOWED_DOC_TYPES.join(', ')}`)
	}
	if (compactStrings(data.read_when).length === 0) errors.push(`${rel}: \`read_when\` must contain at least one hint`)
	return errors
}

function checkCodeFences(docsDir, rel) {
	const text = readFileSync(path.join(docsDir, rel), 'utf8')
	const fences = text.split(/\r?\n/).filter((line) => line.trimStart().startsWith('```')).length
	return fences % 2 ? [`${rel}: unbalanced fenced code block`] : []
}

function checkDocsJson(docsDir, relPaths) {
	const configPath = path.join(docsDir, 'docs.json')
	if (!existsSync(configPath)) return ['docs/docs.json: missing docs navigation file']

	let config
	try {
		config = JSON.parse(readFileSync(configPath, 'utf8'))
	} catch (error) {
		return [`docs/docs.json: invalid JSON: ${error instanceof Error ? error.message : String(error)}`]
	}

	const pageKeys = new Set(relPaths.map((rel) => rel.replace(/\.md$/, '')))
	const errors = []
	for (const page of collectNavPages(config.navigation ?? [])) {
		const key = page.trim().replace(/^\/+|\/+$/g, '')
		if (!pageKeys.has(key)) errors.push(`docs/docs.json: navigation page \`${page}\` has no matching docs page`)
	}
	return errors
}

function checkLinks(docsDir, rel, routes) {
	const filePath = path.join(docsDir, rel)
	const text = readFileSync(filePath, 'utf8')
	const errors = []
	const docsRoot = path.resolve(docsDir)

	for (const match of text.matchAll(MARKDOWN_LINK_RE)) {
		const target = (match[1] ?? '').trim()
		if (!target || target.startsWith('#') || isExternal(target)) continue

		const targetPath = target.split('#', 1)[0]?.split('?', 1)[0] ?? ''
		if (targetPath.startsWith('/')) {
			if (!routes.has(normalizeRoute(targetPath))) errors.push(`${rel}: broken docs route \`${target}\``)
			continue
		}
		const resolved = path.resolve(path.dirname(filePath), targetPath)
		if (!resolved.startsWith(docsRoot)) continue // points outside the docs tree; not ours to check
		if (targetPath.endsWith('.md') && !existsSync(resolved)) errors.push(`${rel}: broken relative link \`${target}\``)
	}
	return errors
}

/**
 * The URL prefix the published site actually answers on, derived rather than restated: the
 * GitHub Pages host comes from `repository`, the path from the VitePress `base`. Writing the
 * value out a second time is what let `llms.txt` drift to a host that never existed.
 */
function canonicalBase(docsDir) {
	if (process.env.DOCS_BASE_URL) return { base: process.env.DOCS_BASE_URL.replace(/\/?$/, '/'), errors: [] }

	const pkg = JSON.parse(readFileSync(path.join(docsDir, '..', 'package.json'), 'utf8'))
	const slug = String(pkg.repository?.url ?? pkg.homepage ?? '').match(/github\.com\/([^/]+)\/([^/.#]+)/)
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

// How each generated file writes the URLs it advertises. These have to stay per-file rather than
// being tried against both: llms-full.txt embeds whole page bodies, so a link pattern applied to
// it would also match the ordinary markdown links *inside* those bodies.
const ADVERTISED_URL_RE = {
	'llms-full.txt': /^URL: (\S+)$/gm,
	'llms.txt': /^- \[[^\]]+\]\((\S+?)\)/gm,
}

/** The URLs a generated llms file advertises. */
function advertisedUrls(text, pattern) {
	const urls = [...text.matchAll(pattern)].map((match) => match[1] ?? '')
	return new Set(urls.map((url) => url.split('#', 1)[0]?.split('?', 1)[0] ?? ''))
}

/**
 * Every advertised URL must map to a file in the build. VitePress runs with `cleanUrls`, so a
 * leaf page is `x.html` and only a directory index is `x/index.html` — the mapping below is the
 * one the server applies, so a route the generator invents cannot pass by looking plausible.
 */
function checkGeneratedUrls(docsDir) {
	const dist = path.join(docsDir, '.vitepress', 'dist')
	if (!existsSync(dist)) return ['docs/.vitepress/dist: missing; run `vitepress build docs` before `--dist`']

	const { base, errors } = canonicalBase(docsDir)
	if (!base) return errors

	for (const [name, pattern] of Object.entries(ADVERTISED_URL_RE)) {
		const file = path.join(dist, name)
		if (!existsSync(file)) {
			errors.push(`${name}: not present in the build output`)
			continue
		}
		const urls = advertisedUrls(readFileSync(file, 'utf8'), pattern)
		if (urls.size === 0) {
			errors.push(`${name}: advertises no URLs, so nothing was verified`)
			continue
		}
		for (const url of [...urls].sort()) {
			if (!url.startsWith(base)) {
				errors.push(`${name}: \`${url}\` is not under the published base \`${base}\``)
				continue
			}
			const rest = url.slice(base.length)
			const page = rest === '' ? 'index.html' : rest.endsWith('/') ? `${rest}index.html` : `${rest}.html`
			if (!existsSync(path.join(dist, page))) errors.push(`${name}: \`${url}\` has no page in the build (${page})`)
		}
	}
	return errors
}

const docsDir = requireDocsDir('docs:check')

if (process.argv.includes('--dist')) {
	const errors = checkGeneratedUrls(docsDir)
	if (errors.length > 0) {
		for (const error of errors) console.error(`docs:check: ${error}`)
		process.exit(1)
	}
	console.log('docs:check: ok (generated llms URLs all resolve to a built page)')
	process.exit(0)
}

const relPaths = walkDocs(docsDir)

const routes = new Set()
for (const rel of relPaths) for (const route of routesFor(rel)) routes.add(route)

const errors = [...checkDocsJson(docsDir, relPaths)]
for (const rel of relPaths) {
	errors.push(...checkFrontmatter(docsDir, rel), ...checkCodeFences(docsDir, rel), ...checkLinks(docsDir, rel, routes))
}

if (errors.length > 0) {
	for (const error of errors) console.error(`docs:check: ${error}`)
	process.exit(1)
}
console.log(`docs:check: ok (${relPaths.length} docs page(s))`)
