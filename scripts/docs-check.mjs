#!/usr/bin/env node
// Validate the docs tree: frontmatter schema, balanced code fences, navigation entries that
// resolve to a real page, and internal links that resolve to a real route or file.
//
// Runs in `docs:check` and again inside `docs:build`, so a broken link fails the docs build
// rather than shipping a 404.

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

const docsDir = requireDocsDir('docs:check')
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
