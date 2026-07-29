#!/usr/bin/env node
/**
 * Raw-XML ratchet — hand-concatenated OOXML can only ever decrease.
 *
 * `src/gen/oxml/el.ts` exists to end hand-built XML strings: its header names
 * escaping, attribute-order and child-sequence bugs as the motivation. Most of
 * `src/gen/` now builds through it, but the chart emitters still concatenate, so
 * a plain "no raw XML anywhere" rule cannot be turned on today. Without *some*
 * gate the migration silently un-does itself: nothing stops a new emitter from
 * being written as a template literal, and nothing stops a migrated file from
 * growing one back.
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
 * one in a template literal is. Two further exemptions:
 *
 * - `src/gen/oxml/` — emitting those delimiters is its job.
 * - Arguments to a message sink (`warn`, `notes.note`, `new *Error`). A warning
 *   that names the element it is talking about is prose, and prose that has to
 *   dodge the gate would get written worse to pass it.
 */

import fs from 'node:fs'
import path from 'node:path'
import ts from 'typescript'
import { ROOT } from './script-utils.mjs'

const SRC = path.join(ROOT, 'src')
const BUDGET = path.join(ROOT, 'scripts', 'raw-xml-budget.json')
/** The builder itself, and anything else under it, is where these delimiters belong. */
const EXEMPT_DIR = 'src/gen/oxml/'
/** `<ns:name` or `</ns:name` — a namespaced XML tag delimiter. */
const TAG_DELIMITER = /<\/?[a-zA-Z][\w.-]*:[a-zA-Z]/g
/** Callees whose string arguments are messages for a human, never emitted bytes. */
const MESSAGE_SINKS = new Set(['warn', 'note'])

const argv = process.argv.slice(2)
const freeze = argv.includes('--freeze')
const list = argv.includes('--list')

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
function isMessageArgument(node) {
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
 * Every tag delimiter inside a string or template literal in one file.
 * @param {string} file absolute path
 * @returns {Array<{ line: number, text: string }>}
 */
function findingsIn(file) {
	const text = fs.readFileSync(file, 'utf8')
	const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true)
	/** @type {Array<{ line: number, text: string }>} */
	const found = []

	/** @param {ts.Node} node */
	function visit(node) {
		// Template *spans* are visited as their own literal nodes, so a multi-part
		// template literal is covered piece by piece rather than as one blob.
		if ((ts.isStringLiteralLike(node) || ts.isTemplateLiteralToken(node)) && !isMessageArgument(node)) {
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

/** @type {Map<string, Array<{ line: number, text: string }>>} */
const findings = new Map()
for (const file of tsFilesUnder(SRC)) {
	const rel = path.relative(ROOT, file).split(path.sep).join('/')
	if (rel.startsWith(EXEMPT_DIR)) continue
	const found = findingsIn(file)
	if (found.length) findings.set(rel, found)
}

/** @param {string} file @returns {number} */
const found = (file) => findings.get(file)?.length ?? 0

const ordered = [...findings.keys()].sort((a, b) => found(b) - found(a) || a.localeCompare(b))
const total = ordered.reduce((sum, file) => sum + found(file), 0)

if (list) {
	for (const file of ordered) {
		console.log(`${file} (${found(file)})`)
		for (const { line, text } of findings.get(file) ?? []) console.log(`  ${file}:${line}  ${text}`)
	}
	process.exit(0)
}

if (freeze) {
	/** @type {Record<string, number>} */
	const frozen = {}
	for (const file of [...ordered].sort()) frozen[file] = found(file)
	fs.writeFileSync(BUDGET, JSON.stringify(frozen, null, '\t') + '\n')
	console.log(`raw-xml ratchet: froze ${ordered.length} file(s), ${total} occurrence(s) -> ${path.basename(BUDGET)}`)
	process.exit(0)
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
	process.exit(1)
}

if (underBudget.length || cleared.length) {
	console.log(`raw-xml ratchet: below budget — lower ${relBudget} in the same commit:\n`)
	for (const file of underBudget) console.log(`  ${file}: ${budget(file)} -> ${found(file)}`)
	for (const file of cleared) console.log(`  ${file}: ${budget(file)} -> 0 (drop the entry)`)
	console.log('\n  pnpm run raw-xml:freeze')
	process.exit(1)
}

console.log(`raw-xml ratchet: ok (${total} occurrence(s) in ${ordered.length} file(s), none above budget)`)
