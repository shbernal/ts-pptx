/**
 * {@link IrValue} → TypeScript source text.
 *
 * This module is deliberately the dumbest part of the subsystem. It knows nothing about
 * decks, shapes, or units — only how a value is spelled. Every semantic decision was made
 * upstream in the read half, which is what lets the printed text be checked by comparing
 * IRs rather than by reading source: if this file were allowed to interpret, the same
 * construct could mean two things depending on which half you asked.
 *
 * Two consequences worth stating, because they are easy to erode later:
 *
 * - **Numbers are printed, not reformatted.** The IR already holds geometry either as a
 *   `` `${n}emu` `` string (exact) or as a six-decimal inch number (the proven minimum for
 *   an EMU-exact round-trip). Rounding here to suppress `0.5000000001`-style noise would
 *   turn a cosmetic problem into a real geometry loss, so `printNumber` only refuses
 *   values JavaScript cannot spell back.
 * - **Layout is a function of the value alone.** Whether an object prints on one line or
 *   several depends only on its rendered width, never on where it came from, so the same
 *   deck prints byte-identical text every time.
 */
import type { AssetRef, IrValue } from '../ir.js'
import { isAssetRef } from '../ir.js'
import { InvalidOptionError } from '../../errors.js'

/** Matches `.oxfmtrc.jsonc`: `printWidth: 120`. */
const MAX_WIDTH = 120

/**
 * Columns a tab is assumed to occupy when measuring against {@link MAX_WIDTH}. The
 * formatter's default `tabWidth`, since the repo sets `useTabs` but not `tabWidth`.
 */
const TAB_COLUMNS = 2

/** A bare object key needs no quotes only if it is a plain identifier. */
const IDENTIFIER = /^[A-Za-z_$][A-Za-z0-9_$]*$/

/**
 * Resolves an {@link AssetRef} to the source text that stands in for it — a `const`
 * identifier when assets are written as files, an inline `data:` literal otherwise. The
 * printer takes this as a callback rather than deciding itself, because the choice belongs
 * to the caller's output layout, not to how a value is spelled.
 */
export type AssetPrinter = (ref: AssetRef) => string

/**
 * A string literal in the repo's style (`singleQuote: true`).
 *
 * Deck text is arbitrary user content, so every case below is reachable in practice: a
 * quote, a backslash, a hard line break inside a run. `JSON.stringify` would handle those
 * correctly — it is avoided here for the quote style, not for correctness.
 */
export function printString(value: string): string {
	let out = "'"
	for (const char of value) {
		switch (char) {
			case '\\':
				out += '\\\\'
				break
			case "'":
				out += "\\'"
				break
			case '\n':
				out += '\\n'
				break
			case '\r':
				out += '\\r'
				break
			case '\t':
				out += '\\t'
				break
			default: {
				const code = char.codePointAt(0) ?? 0
				// C0/C1 controls have no visible spelling, so a raw one is invisible corruption.
				// U+2028/U+2029 have been legal inside a string literal since ES2019 and do not
				// strictly need escaping for a modern engine — they are escaped anyway because
				// they remain line terminators everywhere *else* in the grammar, and a file that
				// depends on that distinction is a poor thing to hand to an unknown toolchain.
				// Everything else rides through, so the script stays readable in the deck's language.
				const escape = code < 0x20 || (code >= 0x7f && code <= 0x9f) || code === 0x2028 || code === 0x2029
				out += escape ? '\\u' + code.toString(16).padStart(4, '0') : char
			}
		}
	}
	return `${out}'`
}

/**
 * A number literal. `String` gives the shortest text that reads back as the same double,
 * which is exactly the guarantee needed: the printed script must produce the same value the
 * IR held. Non-finite values cannot satisfy that (`NaN` and `Infinity` are identifiers, not
 * literals) and never reach here from a well-formed IR, so they throw rather than emit
 * source that would not run.
 */
export function printNumber(value: number): string {
	if (!Number.isFinite(value))
		throw new InvalidOptionError('script/non-finite-literal', `Cannot print a non-finite number: ${value}`)
	return String(value)
}

/** An object key, bare when it is a plain identifier and quoted when it is not. */
function printKey(key: string): string {
	return IDENTIFIER.test(key) ? key : printString(key)
}

/** The value on a single line, however long that line comes out. */
function inline(value: IrValue, asset: AssetPrinter): string {
	if (value === null) return 'null'
	if (typeof value === 'boolean') return String(value)
	if (typeof value === 'number') return printNumber(value)
	if (typeof value === 'string') return printString(value)
	if (isAssetRef(value)) return asset(value)
	if (Array.isArray(value)) return `[${value.map((item) => inline(item, asset)).join(', ')}]`

	const keys = Object.keys(value)
	if (keys.length === 0) return '{}'
	return `{ ${keys.map((key) => `${printKey(key)}: ${inline(value[key] as IrValue, asset)}`).join(', ')} }`
}

/**
 * The value as source text, broken across lines only where it would otherwise overflow
 * {@link MAX_WIDTH}.
 *
 * `indentLevel` is the depth the value *starts* at, so a caller that has already written a
 * prefix (`slide1.addText(`) passes the level its closing brace should line up with.
 */
export function printValue(value: IrValue, indentLevel: number, asset: AssetPrinter): string {
	const flat = inline(value, asset)
	if (indentLevel * TAB_COLUMNS + flat.length <= MAX_WIDTH) return flat

	// Only containers can be broken up; a single long string or number stays long.
	const pad = '\t'.repeat(indentLevel)
	const inner = '\t'.repeat(indentLevel + 1)

	if (Array.isArray(value)) {
		if (value.length === 0) return '[]'
		const items = value.map((item) => `${inner}${printValue(item, indentLevel + 1, asset)},`)
		return `[\n${items.join('\n')}\n${pad}]`
	}

	if (typeof value === 'object' && value !== null && !isAssetRef(value)) {
		const keys = Object.keys(value)
		if (keys.length === 0) return '{}'
		const entries = keys.map(
			(key) => `${inner}${printKey(key)}: ${printValue(value[key] as IrValue, indentLevel + 1, asset)},`
		)
		return `{\n${entries.join('\n')}\n${pad}}`
	}

	return flat
}

/**
 * A call's argument list, either inline after `prefix` or one argument per line.
 *
 * The decision is made on the *whole* call rather than per argument, so a call either reads
 * as one line or as a list — never as a short first argument followed by a hanging brace,
 * which is what a per-argument rule produces.
 */
export function printArguments(prefix: string, args: IrValue[], indentLevel: number, asset: AssetPrinter): string {
	const pad = '\t'.repeat(indentLevel)
	const flat = `${pad}${prefix}(${args.map((arg) => inline(arg, asset)).join(', ')})`
	if (indentLevel * TAB_COLUMNS + flat.length - pad.length <= MAX_WIDTH) return flat

	// No trailing comma after the final argument: `trailingComma: 'es5'` adds them inside
	// objects and arrays but not to an argument list, so one here would be reformatted away.
	const inner = '\t'.repeat(indentLevel + 1)
	const printed = args.map((arg) => `${inner}${printValue(arg, indentLevel + 1, asset)}`)
	return `${pad}${prefix}(\n${printed.join(',\n')}\n${pad})`
}
