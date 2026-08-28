#!/usr/bin/env node
/**
 * Raw-XML ratchet — hand-concatenated OOXML can only ever decrease.
 *
 * `src/gen/oxml/el.ts` exists to end hand-built XML strings: its header names
 * escaping, attribute-order and child-sequence bugs as the motivation. Most of
 * `src/gen/` now builds through it, but a plain "no raw XML anywhere" rule still
 * cannot be turned on. Two kinds of holdout remain.
 *
 * The bulk of it is simply not migrated yet — `src/gen/pres/theme.ts` alone is
 * more than half the total, with the rest across `src/gen/drawingml/`,
 * `src/gen/slide/`, `src/embedded-fonts.ts`, `src/math.ts` and two files under
 * `src/read/api/`. Those are work, not obstacles.
 *
 * The small remainder cannot be migrated at all. Seven occurrences under
 * `src/gen/chart/` pad the space between a tag name and an attribute —
 * `<c:showOutline    val="0"/>`, ` <c:baseTimeUnit  val="days"/>`, and the
 * `<a:defRPr ${sizeAttr} b=…>` in `genXmlTitle` that emits two spaces when no
 * font size is set. `el()` writes exactly one space before an attribute, by
 * design, and its `fmt` describes whitespace *around* an element rather than
 * inside its open tag. Dropping the padding would be the real fix and no XML
 * consumer could see the difference, but it is still an output change, and
 * AGENTS.md is explicit that a whitespace-only diff is a stop rather than a
 * cleanup — so each one stays hand-written with the reason at its call site.
 *
 * Without *some* gate the migration silently un-does itself: nothing stops a new
 * emitter from being written as a template literal, and nothing stops a migrated
 * file from growing one back.
 *
 * So this is a ratchet, not a ban. `budget.json` freezes today's per-file count.
 * A file may go down or vanish; it may never go up, and a file not in the budget
 * must be at zero. Lowering a number is the whole point — the check prints the
 * new value to paste in.
 *
 *   node scripts/raw-xml-ratchet.mjs            # check (exit 1 on any regression)
 *   node scripts/raw-xml-ratchet.mjs --freeze   # rewrite budget.json from source
 *   node scripts/raw-xml-ratchet.mjs --list     # every occurrence, with line numbers
 *
 * What counts: an XML tag delimiter — `<ns:name` or `</ns:name` — inside a string
 * or template literal in `src/`. The scan is over the TypeScript AST rather than
 * the file text, so the `<a:bodyPr>` in a doc comment is not a finding while the
 * one in a template literal is. Three further exemptions:
 *
 * - `src/gen/oxml/` — emitting those delimiters is its job.
 * - Arguments to a message sink (`warn`, `notes.note`, `new *Error`). A warning
 *   that names the element it is talking about is prose, and prose that has to
 *   dodge the gate would get written worse to pass it.
 * - A declaration marked `@raw-xml-asset`. This is the difference between XML this
 *   library *builds* and XML it *ships*: a payload captured verbatim from PowerPoint
 *   and stored as a constant carries none of the escaping, attribute-order or
 *   child-sequence risk the ratchet exists to contain — there is nothing to
 *   interpolate and nothing to migrate, because `el()` would only re-encode a byte
 *   sequence that must stay exactly as Office wrote it. Counting them anyway put the
 *   single largest number in the budget (`chartex-style.ts`, 542) on the one file
 *   that will never move, which buried the real holdouts and made the total read as
 *   a migration backlog nearly five times its actual size.
 *
 *   The marker is deliberately narrow: it exempts one declaration, must be written
 *   on it, and does not apply to anything that interpolates. If you find yourself
 *   reaching for it on something built at runtime, it is the wrong tool.
 */

import fs from 'node:fs'
import path from 'node:path'
// `typescript-6`, not `typescript`: this script walks a real syntax tree
// (`createSourceFile`, `forEachChild`, the `isXxx` predicates), and TypeScript 7 is the
// native Go compiler — its npm package ships a `tsc` binary and no JavaScript compiler
// API at all, so the bare specifier resolves to something with none of these on it. The
// root stays on 7 for `tsc`; this one consumer takes an aliased 6. Unlike TypeDoc, which
// hard-codes the bare specifier and therefore needs the tools/api-docs workspace package,
// an alias is enough here because the import is ours to name.
import ts from 'typescript-6'
import { ROOT, isMain, parseCli, runCli } from './script-utils.mjs'

const SRC = path.join(ROOT, 'src')
const BUDGET = path.join(ROOT, 'scripts', 'raw-xml-budget.json')
/** The builder itself, and anything else under it, is where these delimiters belong. */
const EXEMPT_DIR = 'src/gen/oxml/'
/** `<ns:name` or `</ns:name` — a namespaced XML tag delimiter. */
const TAG_DELIMITER = /<\/?[a-zA-Z][\w.-]*:[a-zA-Z]/g
/** Callees whose string arguments are messages for a human, never emitted bytes. */
const MESSAGE_SINKS = new Set(['warn', 'note'])
/** Marks a declaration holding XML captured verbatim from Office rather than built here. */
const ASSET_MARKER = '@raw-xml-asset'

/** @param {string} dir @returns {string[]} */
function tsFilesUnder(dir) {
	return fs
		.readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) =>
			entry.isDirectory()
				? tsFilesUnder(path.join(dir, entry.name))
				: entry.name.endsWith('.ts')
					? [path.join(dir, entry.name)]
					: []
		)
}

/**
 * Is this literal an argument to the nearest enclosing call, and is that call a
 * message sink? Nearest, so `warn('code', buildXml())` still counts what
 * `buildXml` builds — only the literal handed straight to the sink is exempt.
 * @param {ts.Node} node
 * @returns {boolean}
 */
export function isMessageArgument(node) {
	for (let child = node, parent = node.parent; parent; child = parent, parent = parent.parent) {
		if (!ts.isCallExpression(parent) && !ts.isNewExpression(parent)) continue
		if (!parent.arguments?.some((arg) => arg === child)) return false
		const callee = parent.expression
		const name = ts.isPropertyAccessExpression(callee) ? callee.name.text : ts.isIdentifier(callee) ? callee.text : ''
		return MESSAGE_SINKS.has(name) || (ts.isNewExpression(parent) && name.endsWith('Error'))
	}
	return false
}

/**
 * Is this literal the initializer of a declaration marked {@link ASSET_MARKER}?
 *
 * Walks out to the enclosing statement and reads its leading trivia, so the marker can sit in
 * the JSDoc that already documents the constant. Only a plain literal qualifies: a template
 * with substitutions is something being *built*, whatever it is labelled.
 * @param {ts.Node} node
 * @param {string} text full file text, for reading comment ranges
 * @returns {boolean}
 */
export function isCapturedAsset(node, text) {
	if (ts.isTemplateExpression(node.parent)) return false
	for (let current = node; current; current = current.parent) {
		if (!ts.isVariableStatement(current) && !ts.isPropertyDeclaration(current)) continue
		const ranges = ts.getLeadingCommentRanges(text, current.getFullStart()) ?? []
		return ranges.some((range) => text.slice(range.pos, range.end).includes(ASSET_MARKER))
	}
	return false
}

/**
 * Every tag delimiter inside a string or template literal in one file.
 * @param {string} file absolute path
 * @returns {Array<{ line: number, text: string }>}
 */
export function findingsIn(file) {
	return scanSource(fs.readFileSync(file, 'utf8'), file)
}

/**
 * The scan, over source text rather than a path.
 *
 * Separated from {@link findingsIn} so the exemption rules — the message-sink and
 * captured-asset carve-outs, which are where this gate can silently stop counting —
 * can be exercised against a literal snippet in `test/scripts/raw-xml-ratchet.test.js`
 * instead of a fixture file on disk.
 * @param {string} text TypeScript source
 * @param {string} [fileName] name used for the synthetic source file
 * @returns {Array<{ line: number, text: string }>}
 */
export function scanSource(text, fileName = 'input.ts') {
	const source = ts.createSourceFile(fileName, text, ts.ScriptTarget.Latest, true)
	/** @type {Array<{ line: number, text: string }>} */
	const found = []

	/** @param {ts.Node} node */
	function visit(node) {
		// Template *spans* are visited as their own literal nodes, so a multi-part
		// template literal is covered piece by piece rather than as one blob.
		if (
			(ts.isStringLiteralLike(node) || ts.isTemplateLiteralToken(node)) &&
			!isMessageArgument(node) &&
			!isCapturedAsset(node, text)
		) {
			for (const match of node.text.matchAll(TAG_DELIMITER)) {
				const { line } = source.getLineAndCharacterOfPosition(node.getStart(source))
				found.push({ line: line + 1, text: match[0] })
			}
		}
		ts.forEachChild(node, visit)
	}
	visit(source)
	return found
}

/**
 * Scan all of `src/`, minus the exempt directory.
 * @returns {Map<string, Array<{ line: number, text: string }>>} keyed by repo-relative path
 */
export function collectFindings() {
	const findings = new Map()
	for (const file of tsFilesUnder(SRC)) {
		const rel = path.relative(ROOT, file).split(path.sep).join('/')
		if (rel.startsWith(EXEMPT_DIR)) continue
		const found = findingsIn(file)
		if (found.length) findings.set(rel, found)
	}
	return findings
}

// ---------------------------------------------------------------- CLI

const USAGE = `Raw-XML ratchet — hand-concatenated OOXML can only ever decrease.

  node scripts/raw-xml-ratchet.mjs            check (exit 1 on any regression)
  node scripts/raw-xml-ratchet.mjs --freeze   rewrite the budget from source
  node scripts/raw-xml-ratchet.mjs --list     every occurrence, with line numbers

Options:
  --freeze    record today's per-file counts as the new budget
  --list      print every occurrence rather than just the budget comparison
  -h, --help  show this message`

/** @param {string[]} argv @returns {number} process exit code */
export function main(argv) {
	const { values } = parseCli(argv, {
		usage: USAGE,
		options: {
			freeze: { type: 'boolean', default: false },
			list: { type: 'boolean', default: false },
		},
	})

	const findings = collectFindings()
	/** @param {string} file @returns {number} */
	const found = (file) => findings.get(file)?.length ?? 0

	const ordered = [...findings.keys()].sort((a, b) => found(b) - found(a) || a.localeCompare(b))
	const total = ordered.reduce((sum, file) => sum + found(file), 0)

	if (values.list) {
		for (const file of ordered) {
			console.log(`${file} (${found(file)})`)
			for (const { line, text } of findings.get(file) ?? []) console.log(`  ${file}:${line}  ${text}`)
		}
		return 0
	}

	if (values.freeze) {
		/** @type {Record<string, number>} */
		const frozen = {}
		for (const file of [...ordered].sort()) frozen[file] = found(file)
		fs.writeFileSync(BUDGET, JSON.stringify(frozen, null, '\t') + '\n')
		console.log(`raw-xml ratchet: froze ${ordered.length} file(s), ${total} occurrence(s) -> ${path.basename(BUDGET)}`)
		return 0
	}

	/** @type {Record<string, number>} */
	const budgetFile = JSON.parse(fs.readFileSync(BUDGET, 'utf8'))
	/** @param {string} file @returns {number} */
	const budget = (file) => budgetFile[file] ?? 0

	const relBudget = path.relative(ROOT, BUDGET).split(path.sep).join('/')
	const overBudget = ordered.filter((file) => found(file) > budget(file))
	const underBudget = ordered.filter((file) => found(file) < budget(file))
	const cleared = Object.keys(budgetFile).filter((file) => !findings.has(file))

	if (overBudget.length) {
		console.error('raw-xml ratchet FAILED — hand-built XML grew:\n')
		for (const file of overBudget) {
			console.error(`  ${file}: ${found(file)} (budget ${budget(file)})`)
			for (const { line, text } of findings.get(file) ?? []) console.error(`    ${file}:${line}  ${text}`)
		}
		console.error(`\nBuild these through \`el()\`/\`voidEl()\` from ${EXEMPT_DIR}el.ts.`)
		console.error('If the growth is genuinely unavoidable, raise the number in')
		console.error(`${relBudget} in the same commit, with the reason in the message.`)
		return 1
	}

	if (underBudget.length || cleared.length) {
		console.log(`raw-xml ratchet: below budget — lower ${relBudget} in the same commit:\n`)
		for (const file of underBudget) console.log(`  ${file}: ${budget(file)} -> ${found(file)}`)
		for (const file of cleared) console.log(`  ${file}: ${budget(file)} -> 0 (drop the entry)`)
		console.log('\n  pnpm run raw-xml:freeze')
		return 1
	}

	console.log(`raw-xml ratchet: ok (${total} occurrence(s) in ${ordered.length} file(s), none above budget)`)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
