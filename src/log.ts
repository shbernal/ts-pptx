/**
 * Central sink for library diagnostics.
 *
 * Every user-facing warning is routed through {@link warn} / {@link warnOnce} rather than calling
 * `console.warn` directly, so the library has one place to apply a consistent prefix and, later,
 * mute or redirect output (a warning handler) without editing call sites. This module intentionally
 * has no internal imports so it can be used from anywhere without risking an import cycle.
 */

/** Prefix stamped on every warning so consumers can attribute console noise to this library. */
const WARN_PREFIX = 'ts-pptx'

/**
 * Emit a library warning to the console.
 * @param {string} message - warning text (without a `Warning:`/`ts-pptx:` prefix; the prefix is added here)
 */
export function warn(message: string): void {
	console.warn(`${WARN_PREFIX}: ${message}`)
}

// Track messages already emitted so a recurring condition (e.g. the same out-of-range fontSize on
// every cell of a table) warns once instead of flooding the console.
const warnedMessages = new Set<string>()

/**
 * Emit a library warning at most once per distinct message for the life of the process.
 * @param {string} message - warning text (see {@link warn})
 */
export function warnOnce(message: string): void {
	if (warnedMessages.has(message)) return
	warnedMessages.add(message)
	warn(message)
}
