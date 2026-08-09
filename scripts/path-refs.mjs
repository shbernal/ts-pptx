#!/usr/bin/env node
/**
 * Path-citation gate — a backticked repo path must name a file that exists.
 *
 * This repo cites files constantly, and it does it in backticks rather than as
 * markdown links: a doc says "see `test/regression/shape/group-shapes.test.js`", a
 * source comment says "the read side (`src/read/api/ops/table-styles.ts`) merges …".
 * Those citations carry real weight — they are usually the evidence for the claim
 * beside them, naming the test that proves it or the module that owns it.
 *
 * Nothing checked them. `docs-check.mjs` validates markdown links, but a backticked
 * path is not a link, and three quarters of these live in `src/` and `test/` where a
 * docs gate never looks. So they rotted silently: seven citations were found pointing
 * at files that had moved or been deleted, two of them still describing the upstream
 * demo layout this project replaced. A citation that resolves to nothing is worse than
 * no citation — it costs a reader the lookup and then strands them.
 *
 *   node scripts/path-refs.mjs          # check (exit 1 on any dead citation)
 *   node scripts/path-refs.mjs --list   # every citation found, resolved or not
 *
 * ## What counts as a citation
 *
 * A backticked token containing a `/` and ending in a source-ish extension. The `/` is
 * load-bearing: bare `package.json` or `index.ts` appear in prose constantly and name
 * nothing in particular, so demanding a directory separator is what keeps the signal up.
 *
 * ## How one resolves
 *
 * Any of: relative to the repo root, relative to the citing file, or as a suffix of some
 * file's repo-relative path. That last rule is deliberately loose — comments cite
 * `read/api/rel-types.ts` and `gen/oxml/el.ts` without the `src/` prefix, and demanding
 * full paths would fail honest citations. A `.js`/`.mjs` token also resolves against its
 * `.ts`/`.mts` source, because ESM specifiers in comments name the emitted file.
 *
 * ## What is not checked
 *
 * - `dist/`, `coverage/`, `.tmp/` and demo `output/` — build artifacts. `RELEASING.md`
 *   in particular lists `dist/pptxgen.*` files *on purpose*, as the negative space of
 *   what this package refuses to ship; those must never resolve.
 * - `CHANGELOG.md` — a release log describes the tree as it stood at the time, so a
 *   path that has since moved is a correct historical record, not a dead citation.
 * - The `ALLOWLIST` below: citations that are deliberately unresolvable. Each carries
 *   its reason, and an entry that stops firing is itself reported — a stale exemption is
 *   the same disease this gate exists to catch.
 */

import { readdirSync, readFileSync, existsSync } from 'node:fs'
import path from 'node:path'

import { isMain, ROOT, runCli } from './script-utils.mjs'

/** Trees worth scanning. Everything else is either generated or third-party. */
const SCAN_ROOTS = ['docs', 'src', 'test', 'scripts', 'tools', 'demos', '.github']

/** Files at the repo root that carry citations. `CHANGELOG.md` is excluded on purpose. */
const ROOT_FILES = ['AGENTS.md', 'README.md', 'CONTRIBUTING.md', 'SECURITY.md', 'CLAUDE.md']

/**
 * Directory names never walked, wherever they appear. Deliberately *not* including
 * `test/read/fixtures/` — its README is the densest concentration of citations in the repo
 * and held three of the seven dead ones. The binary decks beside it are skipped by extension,
 * not by directory, so there is nothing to gain by excluding the tree.
 */
const SKIP_DIRS = new Set(['node_modules', 'dist', 'coverage', '.git', '.tmp', '.vitepress', 'output'])

/** Only these carry citations worth resolving; a `.png` or `.pptx` is an asset, not a claim. */
const CITED_EXT = /\.(ts|mts|tsx|js|mjs|cjs|jsx|md|json|jsonc|yml|yaml|html|css|tsv)$/

/** A backticked token that looks like a path, `./`- and `../`-relative forms included. */
const CITATION_RE = /`((?:\.{1,2}\/)*[A-Za-z0-9_@][A-Za-z0-9_./@-]*\.[a-z]+)`/g

/**
 * This file, which cannot be scanned by itself: its header and `ALLOWLIST` quote dead paths
 * as examples of the thing it rejects, and every one of them would be a finding.
 */
const SELF = 'scripts/path-refs.mjs'

/**
 * Directories that hold build output; a citation into one is about an artifact, not a source
 * file. Matched anywhere in the token so that `../../../dist/node.js` — a relative specifier
 * inside a code sample — is recognised as the same thing as `dist/node.js`.
 *
 * `reference/api/` is the TypeDoc output under `docs/`, and is matched on its two-segment tail
 * because the citations to it are relative (`reference/api/index.md`, from `docs/reference/`);
 * a bare `api/` would be broad enough to swallow real source directories.
 */
const GENERATED = /(^|\/)(dist|coverage|\.tmp|output|node_modules|reference\/api)\//

/**
 * Generated *files*, which the directory test above cannot catch. These are skipped on the same
 * grounds — a citation naming one is about an artifact — but the reason they must be skipped
 * rather than allowlisted is the ordering: `path-refs:check` runs before `docs:build` in both
 * `check:static` and `verify`, so in a clean checkout the file does not exist yet, while on a
 * tree that has built docs it does. An `ALLOWLIST` entry would fail in the second case, since a
 * citation that starts resolving is reported as a stale exemption.
 */
const GENERATED_FILES = new Set(['docs/doc-index.md'])

/**
 * Citations that are meant not to resolve. `where` is `<file>:<token>` — the line is left
 * out so ordinary edits above them do not churn this list.
 */
const ALLOWLIST = [
	{
		where: 'docs/testing.md:scripts/demo-smoke.mjs',
		why: 'names the deleted script whose job this section explains taking over',
	},
	{
		where: 'src/ooxml/rel-types.ts:gen/oxml/schema-uris.ts',
		why: 'names one of the two pre-merge copies, to explain why this module exists',
	},
	{
		where: 'src/ooxml/rel-types.ts:read/api/rel-types.ts',
		why: 'the other pre-merge copy, same sentence',
	},
	{
		where: 'scripts/docs-check.mjs:x/index.html',
		why: 'illustrative — `x.html` vs `x/index.html` is the cleanUrls mapping being described',
	},
	{
		where: 'scripts/generate-llms-docs.mjs:tables/index.html',
		why: 'illustrative, same mapping',
	},
	{
		where: 'scripts/docs-new.mjs:guides/my-page.md',
		why: 'illustrative — the usage example for an argument the user supplies',
	},
	{
		where: 'scripts/README.md:./scripts/x.mjs',
		why: 'illustrative — a stand-in filename in the convention description',
	},
]

/** Every file under `dir`, recursively, skipping the generated and vendored trees. */
function walk(dir, out = []) {
	if (!existsSync(dir)) return out
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		if (SKIP_DIRS.has(entry.name)) continue
		const full = path.join(dir, entry.name)
		if (entry.isDirectory()) walk(full, out)
		else out.push(full)
	}
	return out
}

/** Repo-relative, forward-slashed — the form every message and allowlist key uses. */
function rel(file) {
	return path.relative(ROOT, file).split(path.sep).join('/')
}

/**
 * Does `token` name a real file? Three ways, plus the ESM `.js` → `.ts` swap.
 * @param {string} token the cited path
 * @param {string} from the citing file, absolute
 * @param {Set<string>} known every repo-relative path in the scanned trees
 */
function resolves(token, from, known) {
	const candidates = [token]
	// A comment citing `./pattern-fill.js` means the module whose source is `pattern-fill.ts`.
	if (/\.m?js$/.test(token)) candidates.push(token.replace(/\.js$/, '.ts').replace(/\.mjs$/, '.mts'))

	for (const candidate of candidates) {
		if (existsSync(path.resolve(ROOT, candidate))) return true
		if (existsSync(path.resolve(path.dirname(from), candidate))) return true
		// Suffix match, on a `/` boundary so `types.ts` cannot satisfy `shapes/types.ts`.
		const bare = candidate.replace(/^(\.{1,2}\/)+/, '')
		if (known.has(bare)) return true
		for (const file of known) if (file.endsWith(`/${bare}`)) return true
	}
	return false
}

/** Every citation in the scanned trees, with whether it resolved. */
function collect() {
	const files = SCAN_ROOTS.flatMap((r) => walk(path.join(ROOT, r)))
	for (const name of ROOT_FILES) {
		const full = path.join(ROOT, name)
		if (existsSync(full)) files.push(full)
	}

	const known = new Set(files.map(rel))
	const citations = []
	for (const file of files) {
		if (!CITED_EXT.test(file)) continue
		const relFile = rel(file)
		if (relFile === SELF) continue
		const lines = readFileSync(file, 'utf8').split(/\r?\n/)
		lines.forEach((line, index) => {
			for (const match of line.matchAll(CITATION_RE)) {
				const token = match[1] ?? ''
				if (!token.includes('/') || !CITED_EXT.test(token)) continue
				if (token.startsWith('@') || token.startsWith('node:')) continue
				if (GENERATED.test(token) || GENERATED_FILES.has(token)) continue
				citations.push({ file: relFile, line: index + 1, token, ok: resolves(token, file, known) })
			}
		})
	}
	return citations
}

/** @param {string[]} argv */
function main(argv) {
	const citations = collect()

	if (argv.includes('--list')) {
		for (const c of citations) console.log(`${c.ok ? 'ok  ' : 'DEAD'} ${c.file}:${c.line}  ${c.token}`)
		console.log(`\n${citations.length} citation(s), ${citations.filter((c) => !c.ok).length} unresolved`)
		return 0
	}

	const exempt = new Map(ALLOWLIST.map((e) => [e.where, e]))
	const used = new Set()
	const dead = []
	for (const c of citations.filter((x) => !x.ok)) {
		const key = `${c.file}:${c.token}`
		if (exempt.has(key)) used.add(key)
		else dead.push(c)
	}

	let failed = false
	if (dead.length > 0) {
		failed = true
		console.error('path-refs: citation(s) naming a file that does not exist:\n')
		for (const c of dead) console.error(`  ${c.file}:${c.line}  \`${c.token}\``)
		console.error('\nRepoint each at where the file actually is. If it is meant not to resolve')
		console.error('(a historical mention, an illustrative name), add it to ALLOWLIST in')
		console.error('scripts/path-refs.mjs with the reason.')
	}

	const unused = ALLOWLIST.filter((e) => !used.has(e.where))
	if (unused.length > 0) {
		failed = true
		console.error(`${dead.length ? '\n' : ''}path-refs: ALLOWLIST entr(ies) that no longer fire — drop them:\n`)
		for (const e of unused) console.error(`  ${e.where}  (${e.why})`)
	}

	if (failed) return 1
	console.log(`path-refs: ok (${citations.length} citation(s) resolve, ${ALLOWLIST.length} exempt)`)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
