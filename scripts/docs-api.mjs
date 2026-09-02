#!/usr/bin/env node
import { spawnSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import path from 'node:path'

import { parseCliOrExit } from './script-utils.mjs'

// No flags, but `--help` still has to answer and `--bogus` still has to report itself in one
// line -- and both have to happen BEFORE the generator writes anything.
parseCliOrExit(process.argv.slice(2), {
	usage: `Generate the TypeDoc markdown reference into docs/reference/api.

  pnpm run docs:api

Options:
  -h, --help   show this message`,
	options: {},
})

const root = process.cwd()
const outDir = path.join(root, 'docs', 'reference', 'api')
// TypeDoc lives in the tools/api-docs workspace package, not at the root, because it needs
// a TypeScript 6 that the root no longer has: TypeScript 7 ships a native binary and no JS
// compiler API, so TypeDoc cannot import it. See tools/api-docs/README.md. `cwd` below stays
// the repo root, which is what keeps every root-relative path in typedoc.docs.json working,
// and TypeDoc resolves its markdown plugin relative to its own install rather than to cwd.
const typedocBin = path.join(
	root,
	'tools',
	'api-docs',
	'node_modules',
	'.bin',
	process.platform === 'win32' ? 'typedoc.cmd' : 'typedoc'
)

/**
 * Every `.md` file under `dir`, recursively, sorted.
 * @param {string} dir
 * @returns {string[]}
 */
function walkMarkdown(dir) {
	/** @type {string[]} */
	const out = []
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const filePath = path.join(dir, entry.name)
		if (entry.isDirectory()) {
			out.push(...walkMarkdown(filePath))
		} else if (entry.isFile() && entry.name.endsWith('.md')) {
			out.push(filePath)
		}
	}
	return out.sort()
}

/**
 * @param {string} markdown
 * @param {string} filePath
 * @returns {string}
 */
function titleFromMarkdown(markdown, filePath) {
	const heading = markdown.match(/^#\s+(.+)$/m)
	if (heading?.[1]) return heading[1].replace(/\s+\|.*$/, '').trim()
	const basename = path.basename(filePath, '.md')
	return basename === 'index' ? 'Public API Reference' : basename
}

/**
 * @param {string} filePath
 * @param {string} markdown
 * @returns {string}
 */
function frontmatterFor(filePath, markdown) {
	const rel = path.relative(outDir, filePath).split(path.sep).join('/')
	const title = rel === 'index.md' ? 'Public API Reference' : titleFromMarkdown(markdown, filePath)
	const summary =
		rel === 'index.md'
			? 'Generated TypeDoc reference for the public TsPptx package exports.'
			: `Generated TypeDoc reference for ${title}.`

	return [
		'---',
		'doc-schema-version: 1',
		`title: "${title.replaceAll('"', '\\"')}"`,
		`summary: "${summary.replaceAll('"', '\\"')}"`,
		'read_when:',
		'  - Looking up public TsPptx API details',
		'  - Verifying generated TypeScript API documentation',
		'doc_type: "reference"',
		'---',
		'',
	].join('\n')
}

/**
 * @param {string} markdown
 * @returns {string}
 */
function escapeVueUnsafeHtml(markdown) {
	let inFence = false
	return markdown
		.split('\n')
		.map(
			/** @param {string} line */ (line) => {
				if (line.trimStart().startsWith('```')) {
					inFence = !inFence
					return line
				}
				if (inFence) return line

				return line
					.split('`')
					.map(
						/** @param {string} segment @param {number} index */ (segment, index) => {
							if (index % 2 === 1) return segment
							return segment
								.replaceAll(/<\/([A-Za-z][A-Za-z0-9:._-]*)>/g, '&lt;/$1&gt;')
								.replaceAll(/<([A-Za-z][A-Za-z0-9:._-]*)(\s[^>\n]*)?>/g, '&lt;$1$2&gt;')
						}
					)
					.join('`')
			}
		)
		.join('\n')
}

rmSync(outDir, { force: true, recursive: true })
mkdirSync(outDir, { recursive: true })

// Invoke as a single shell string (not an args array) so Windows can run the
// .cmd shim without shell:true's arg-escaping deprecation warning (DEP0190).
const typedoc = spawnSync(`"${typedocBin}" --options typedoc.docs.json`, {
	cwd: root,
	stdio: 'inherit',
	shell: true,
})

if (typedoc.status !== 0) {
	process.exit(typedoc.status ?? 1)
}

const readmePath = path.join(outDir, 'README.md')
const indexPath = path.join(outDir, 'index.md')
if (!existsSync(indexPath) && existsSync(readmePath)) {
	copyFileSync(readmePath, indexPath)
}

for (const filePath of walkMarkdown(outDir)) {
	const markdown = readFileSync(filePath, 'utf8')
	const body = markdown.startsWith('---\n') ? markdown.replace(/^---\n[\s\S]*?\n---\n+/, '') : markdown
	const safeBody = escapeVueUnsafeHtml(body.trimStart())
	writeFileSync(filePath, `${frontmatterFor(filePath, safeBody)}${safeBody}`, 'utf8')
}
