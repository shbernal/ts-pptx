/**
 * Turn a probe's `build` function back into the code a reader can read.
 *
 * The comparison's snippets are not written twice. `scripts/comparison/probes.mjs` already
 * holds each intent as executable code in both libraries' idioms, and that code is what the
 * measurement ran -- so the page shows the function itself, recovered here with
 * `Function.prototype.toString()` and carried into the snapshot beside the outcome it
 * produced. A snippet therefore cannot drift from its row: the two were written by the same
 * run.
 *
 * Two things stand between a function object and a snippet worth printing.
 *
 * The first is the wrapper. A build takes the presentation and returns nothing, so
 * `(pres) => { ... }` is scaffolding a reader has to look past on every one of twenty-one
 * snippets; the page states that frame once and prints the bodies. An arrow shape this
 * module does not recognise throws rather than rendering something almost right -- the
 * corpus is small and hand-written, and a probe formatted a new way is a thing to notice,
 * not to absorb.
 *
 * The second is the corpus constants. A body that reads `PNG_1PX_URL` is complete for the
 * harness and incomplete for a reader, so {@link renderSource} prints the declaration of
 * every constant a body names. Those declarations are formatted from the live values rather
 * than transcribed, because a transcription is a comment that can quietly become false.
 * Long strings are elided, and a path under the repository root is printed relative to it --
 * both because an absolute path is noise to a reader and because a committed snapshot must
 * not carry the directory layout of whichever machine measured it.
 */
import path from 'node:path'
import { ROOT } from '../script-utils.mjs'

/** Strings longer than this are printed truncated, with an ellipsis marking the cut. */
const ELIDE_OVER = 60

/** `(pres) => {` … `}`, optionally `async`. The only shape a probe body may take. */
const BLOCK_ARROW = /^(?:async\s+)?\(\s*\w*\s*\)\s*=>\s*\{\r?\n([\s\S]*?)\r?\n[\t ]*\}$/

/** A key an object literal can carry unquoted. */
const BARE_KEY = /^[A-Za-z_$][A-Za-z0-9_$]*$/

/**
 * Drop the indentation a nested function carries, keeping relative depth.
 * @param {string} text
 * @returns {string}
 */
function dedent(text) {
	const lines = text.split(/\r?\n/)
	const indents = lines.filter((line) => line.trim() !== '').map((line) => (/^[\t ]*/.exec(line) ?? [''])[0].length)
	const common = indents.length > 0 ? Math.min(...indents) : 0
	return lines.map((line) => (line.trim() === '' ? '' : line.slice(common))).join('\n')
}

/**
 * The body of a block-bodied arrow function, dedented.
 * @param {Function} fn
 * @returns {string}
 */
export function functionBody(fn) {
	const text = fn.toString()
	const match = BLOCK_ARROW.exec(text)
	if (!match || match[1] === undefined)
		throw new Error(
			'a probe build must be a block-bodied arrow function taking the presentation, so its body ' +
				'can be printed on the syntax page; got: ' +
				text.split(/\r?\n/)[0]
		)
	return dedent(match[1])
}

/**
 * A value as JavaScript source, elided where it is too long to read.
 *
 * Deliberately narrow: it covers the shapes the corpus actually declares, and throws on
 * anything else rather than printing `[object Object]` into a documentation page.
 *
 * `elideOver` is raised by the one caller that is not writing for a reader: the bundle
 * corpus compiles its own preamble, and a data URL truncated there would be source that no
 * longer carries what the measurement fed the library.
 * @param {unknown} value
 * @param {{elideOver?: number}} [options]
 * @returns {string}
 */
export function literal(value, options = {}) {
	const elideOver = options.elideOver ?? ELIDE_OVER
	if (typeof value === 'string') {
		const shown = value.startsWith(ROOT) ? path.relative(ROOT, value).split(path.sep).join('/') : value
		const cut = shown.length > elideOver ? shown.slice(0, elideOver) + '…' : shown
		return "'" + cut.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'"
	}
	if (typeof value === 'number' || typeof value === 'boolean' || value === null) return String(value)
	if (Array.isArray(value)) return '[' + value.map((entry) => literal(entry, options)).join(', ') + ']'
	if (typeof value === 'object' && value !== null)
		return (
			'{ ' +
			Object.entries(value)
				.map(([key, entry]) => (BARE_KEY.test(key) ? key : "'" + key + "'") + ': ' + literal(entry, options))
				.join(', ') +
			' }'
		)
	throw new Error('no literal form for a corpus constant of type ' + typeof value)
}

/**
 * One probe arm as printable source: the constants it names, then its body.
 *
 * Constants come out in declaration order rather than order of use, so two arms that need
 * the same pair of values cannot differ in the preamble alone -- the page compares these
 * strings to report how many probes call both libraries identically, and an ordering that
 * followed the body would make that number an artefact of where a name first appeared.
 * @param {Function} fn
 * @param {Record<string, unknown>} constants - every value the corpus declares once
 * @param {{elideOver?: number}} [options] - passed through to {@link literal}
 * @returns {string}
 */
export function renderSource(fn, constants, options = {}) {
	const body = functionBody(fn)
	const used = Object.entries(constants).filter(([name]) => new RegExp('\\b' + name + '\\b').test(body))
	const preamble = used.map(([name, value]) => 'const ' + name + ' = ' + literal(value, options))
	return preamble.length > 0 ? preamble.join('\n') + '\n\n' + body : body
}
