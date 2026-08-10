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

import { readdirSync, readFileSync, statSync } from 'node:fs'
import path from 'node:path'

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

/** Drop one layer of matching surrounding quotes, if present. */
function stripQuotes(value) {
	const trimmed = value.trim()
	if (trimmed.length >= 2 && trimmed[0] === trimmed[trimmed.length - 1] && (trimmed[0] === "'" || trimmed[0] === '"')) {
		return trimmed.slice(1, -1)
	}
	return trimmed
}

/** Non-empty, trimmed string values only — the shape every list field is consumed as. */
export function compactStrings(values) {
	if (!Array.isArray(values)) return []
	const out = []
	for (const value of values) {
		if (value === null || value === undefined) continue
		const text = String(value).trim()
		if (text) out.push(text)
	}
	return out
}

/** Parse an inline `[a, b]` list, tolerating single quotes. Returns `[]` when it is not one. */
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
 * @returns {{ data: Record<string, unknown>, error: string | null }}
 */
export function parseFrontmatter(filePath) {
	const raw = readFileSync(filePath, 'utf8')
	if (!raw.startsWith('---\n') && !raw.startsWith('---\r\n')) return { data: {}, error: 'missing front matter' }

	const lines = raw.split(/\r?\n/)
	let endIndex = -1
	for (let i = 1; i < lines.length; i++) {
		const trimmed = (lines[i] ?? '').trim()
		if (trimmed === '---' || trimmed === '...') {
			endIndex = i
			break
		}
	}
	if (endIndex < 0) return { data: {}, error: 'unterminated front matter' }

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

	return { data, error: null }
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
	const found = []
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

/** Resolve the docs directory, exiting with the script's own message when it is absent. */
export function requireDocsDir(label) {
	const docsDir = path.resolve('docs')
	let ok
	try {
		ok = statSync(docsDir).isDirectory()
	} catch {
		ok = false
	}
	if (!ok) {
		console.error(`${label}: missing docs directory. Run from repo root.`)
		process.exit(1)
	}
	return docsDir
}
