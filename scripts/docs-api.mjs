#!/usr/bin/env node
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import path from 'node:path'

import { ROOT, parseCliOrExit, repoRel, runNodeBin } from './script-utils.mjs'

// No flags, but `--help` still has to answer and `--bogus` still has to report itself in one
// line -- and both have to happen BEFORE the generator writes anything.
parseCliOrExit(process.argv.slice(2), {
	usage: `Generate the TypeDoc markdown reference into docs/reference/api.

  pnpm run docs:api

Options:
  -h, --help   show this message`,
	options: {},
})

const root = ROOT
const outDir = path.join(root, 'docs', 'reference', 'api')

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
	// Un-escape the heading before it becomes a YAML scalar. TypeDoc writes markdown, so a generic
	// in the public surface arrives as `ComposeOptions\<Fs\>`, and `\<` is not a legal escape inside
	// a double-quoted YAML string -- vitepress refuses to parse the frontmatter and the whole site
	// build fails on a page nobody hand-wrote.
	if (heading?.[1])
		return heading[1]
			.replace(/\s+\|.*$/, '')
			.replaceAll(/\\([^A-Za-z0-9])/g, '$1')
			.trim()
	const basename = path.basename(filePath, '.md')
	return basename === 'index' ? 'API reference' : basename
}

/**
 * @param {string} filePath
 * @param {string} markdown
 * @returns {string}
 */
function frontmatterFor(filePath, markdown) {
	const rel = path.relative(outDir, filePath).split(path.sep).join('/')
	const title = rel === 'index.md' ? 'API reference' : titleFromMarkdown(markdown, filePath)
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

// TypeDoc lives in the tools/api-docs workspace package, not at the root, because it needs
// a TypeScript 6 that the root no longer has: TypeScript 7 ships a native binary and no JS
// compiler API, so TypeDoc cannot import it. See tools/api-docs/README.md. `cwd` stays the
// repo root, which is what keeps every root-relative path in typedoc.docs.json working, and
// TypeDoc resolves its markdown plugin relative to its own install rather than to cwd.
try {
	await runNodeBin('typedoc', ['--options', 'typedoc.docs.json'], {
		from: path.join(root, 'tools', 'api-docs'),
		cwd: root,
	})
} catch (error) {
	console.error(error instanceof Error ? error.message : String(error))
	process.exit(1)
}

/**
 * Subpaths `package.json#exports` publishes that the generated reference deliberately leaves out.
 * Any other subpath without a module directory fails the run, and so does an entry here that has
 * one, so the list cannot quietly grant slack.
 *
 * `./families` is out because its family objects are typed by the write path's internal plumbing
 * (`PresSlideInternal`, `RenderContext`, `PresentationAuthorContext`, the slide class): documenting
 * it would mean exporting those or silencing each of them.
 */
const UNDOCUMENTED_SUBPATHS = new Set(['./families'])

/**
 * The file one `package.json#exports` entry resolves to under the `default` condition, following
 * nested condition objects (the bare `.` nests `browser`/`node`/`default`, each with its own
 * `types`/`default`).
 * @param {unknown} target
 * @returns {string | undefined}
 */
function defaultTarget(target) {
	if (typeof target === 'string') return target
	if (target && typeof target === 'object' && 'default' in target) return defaultTarget(target.default)
	return undefined
}

/**
 * The one-line description of an entry point on the landing page: the first sentence of the
 * entry file's module comment, which is the leading doc comment tagged `@module` and the same
 * comment TypeDoc renders at the top of that module's page. So the description is edited in
 * `src/`, never here. A `{@link X}` becomes a code span, because the link target is resolved
 * relative to the module and would not resolve from the landing page.
 * @param {string} sourcePath - absolute path of the entry file
 * @returns {string}
 */
function moduleSummary(sourcePath) {
	const comment = readFileSync(sourcePath, 'utf8').match(/^\/\*\*([\s\S]*?)\*\//)?.[1]
	if (comment === undefined || !/^\s*\*\s*@module\s*$/m.test(comment)) {
		throw new Error(`${repoRel(sourcePath)} does not open with a doc comment tagged @module`)
	}
	/** @type {string[]} */
	const paragraph = []
	for (const line of comment.split('\n').map((raw) => raw.replace(/^\s*\*?/, '').trim())) {
		if (line.startsWith('@')) break
		if (line === '') {
			if (paragraph.length > 0) break
			continue
		}
		paragraph.push(line)
	}
	const text = paragraph.join(' ').replaceAll(/\{@link\s+([^\s|}]+)[^}]*\}/g, '`$1`')
	let inCode = false
	for (let i = 0; i < text.length; i++) {
		if (text[i] === '`') inCode = !inCode
		else if (!inCode && '.!?'.includes(text[i] ?? '') && (i + 1 === text.length || text[i + 1] === ' ')) {
			return text.slice(0, i + 1)
		}
	}
	throw new Error(`the module comment of ${repoRel(sourcePath)} has no first sentence ending in a full stop`)
}

// The landing page lists every entry point `package.json#exports` publishes, rather than TypeDoc's
// flat list of module names, so a reader starts from the specifier they import.
const pkg = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'))
/** @type {string[]} */
const landingEntries = []
/** @type {string[]} */
const problems = []
for (const [subpath, target] of Object.entries(pkg.exports)) {
	const file = defaultTarget(target)
	// `./package.json` is exported as a file, not a module.
	if (!file?.endsWith('.js')) continue
	const name = file.match(/^\.\/dist\/([\w-]+)\.js$/)?.[1]
	if (!name) {
		problems.push(`exports["${subpath}"] resolves to ${file}, not to a dist/<name>.js the entry file can be found from`)
		continue
	}
	const specifier = subpath === '.' ? pkg.name : `${pkg.name}${subpath.slice(1)}`
	const generated = existsSync(path.join(outDir, name, 'README.md'))
	if (UNDOCUMENTED_SUBPATHS.has(subpath)) {
		if (generated) problems.push(`exports["${subpath}"] now has a generated module; drop it from UNDOCUMENTED_SUBPATHS`)
		continue
	}
	if (!generated) {
		problems.push(
			`exports["${subpath}"] (${specifier}) has no generated module directory "${name}": add src/${name}.ts to the entryPoints in typedoc.docs.json`
		)
		continue
	}
	try {
		landingEntries.push(
			`- [\`${specifier}\`](${name}/README.md): ${moduleSummary(path.join(root, 'src', `${name}.ts`))}`
		)
	} catch (error) {
		problems.push(error instanceof Error ? error.message : String(error))
	}
}
for (const subpath of UNDOCUMENTED_SUBPATHS) {
	if (!(subpath in pkg.exports))
		problems.push(`UNDOCUMENTED_SUBPATHS names ${subpath}, which package.json does not export`)
}
if (problems.length > 0) {
	for (const problem of problems) console.error(`docs:api: ${problem}`)
	process.exit(1)
}

writeFileSync(
	path.join(outDir, 'index.md'),
	['# API reference', '', 'Each entry point, by the specifier you import it from.', '', ...landingEntries, ''].join(
		'\n'
	),
	'utf8'
)

for (const filePath of walkMarkdown(outDir)) {
	const markdown = readFileSync(filePath, 'utf8')
	const body = markdown.startsWith('---\n') ? markdown.replace(/^---\n[\s\S]*?\n---\n+/, '') : markdown
	const safeBody = escapeVueUnsafeHtml(body.trimStart())
	writeFileSync(filePath, `${frontmatterFor(filePath, safeBody)}${safeBody}`, 'utf8')
}
