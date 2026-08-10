/**
 * Central sink for library diagnostics.
 *
 * Every user-facing warning is routed through {@link warn} / {@link warnOnce} rather than calling
 * `console.warn` directly, so there is one place to apply a consistent prefix and one place a
 * consumer can take over — see {@link setDiagnosticHandler}. This module has no *runtime* imports
 * so it can be used from anywhere without risking an import cycle; `codes.ts` is types only and is
 * erased at compile time, so it does not cost that property.
 *
 * **A diagnostic's `code` is API; its `message` is not.** The code is a stable identifier a
 * consumer may branch on (`if (d.code === 'chart/non-finite-value')`); the wording behind it is
 * free to improve in any release. The vocabulary lives in `codes.ts` and is shared with the thrown
 * surface (`errors.ts`), so a condition that can both warn and throw reads the same either way.
 */

import type { DiagnosticCode } from './codes.js'

export type { DiagnosticCode }

/** One diagnostic emitted by the library. */
export interface Diagnostic {
	/** Stable condition identifier — see {@link DiagnosticCode}. Branch on this, not on the text. */
	readonly code: DiagnosticCode
	/** Human-readable explanation. Free to change in any release; do not parse it. */
	readonly message: string
	/**
	 * Structured context for the condition, when a site has any to give. No site populates it
	 * today; it exists so one can start to without a breaking signature change.
	 */
	readonly detail?: Readonly<Record<string, unknown>>
}

/** Receives every {@link Diagnostic} the library emits. See {@link setDiagnosticHandler}. */
export type DiagnosticHandler = (diagnostic: Diagnostic) => void

/** Prefix stamped on every warning so consumers can attribute console noise to this library. */
const WARN_PREFIX = 'ts-pptx'

/** The default handler: one `console.warn` line per diagnostic, prefixed. */
const consoleHandler: DiagnosticHandler = (diagnostic) => {
	console.warn(`${WARN_PREFIX}: ${diagnostic.message}`)
}

let handler: DiagnosticHandler = consoleHandler

/**
 * Route every library diagnostic to `handler`, or pass `null` to restore the default
 * (one prefixed `console.warn` per diagnostic).
 *
 * This is process-global rather than per-presentation, deliberately: the emitting code is a tree
 * of free functions in `gen/**` with no presentation in scope, and threading a handler through
 * every signature would be a far larger and worse change than the problem warrants. The practical
 * consequence is that a process building several decks concurrently cannot attribute a diagnostic
 * to one of them; if that matters, correlate on `code` or set the handler around each build.
 *
 * A throwing handler propagates out of whatever library call emitted the diagnostic, which is the
 * supported way to make a specific condition fatal:
 *
 * ```ts
 * setDiagnosticHandler((d) => {
 *   if (d.code === 'coord/bare-number-is-inches') throw new Error(d.message)
 * })
 * ```
 * @param next - the receiver, or `null` to restore the console default
 */
export function setDiagnosticHandler(next: DiagnosticHandler | null): void {
	handler = next ?? consoleHandler
}

/**
 * Emit a library diagnostic.
 * @param code - the stable condition identifier ({@link DiagnosticCode})
 * @param message - explanation, without a `ts-pptx:` prefix (the default handler adds one)
 * @param detail - optional structured context
 */
export function warn(code: DiagnosticCode, message: string, detail?: Readonly<Record<string, unknown>>): void {
	handler(detail === undefined ? { code, message } : { code, message, detail })
}

// Track diagnostics already emitted so a recurring condition (e.g. the same out-of-range fontSize
// on every cell of a table) is reported once instead of flooding the handler. Keyed on code AND
// message, so a *different* offending value under the same code still reports.
const seen = new Set<string>()

/**
 * Emit a library diagnostic at most once per distinct code+message for the life of the process.
 * @param code - the stable condition identifier ({@link DiagnosticCode})
 * @param message - explanation (see {@link warn})
 * @param detail - optional structured context
 */
export function warnOnce(code: DiagnosticCode, message: string, detail?: Readonly<Record<string, unknown>>): void {
	const key = `${code}\0${message}`
	if (seen.has(key)) return
	seen.add(key)
	warn(code, message, detail)
}
