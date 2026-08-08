/**
 * Validating a value against an `ST_` enumeration, under the two policies the library uses.
 *
 * **Why validate at all.** A value outside its `ST_` union makes the containing part
 * schema-invalid, and PowerPoint reports that as a *corrupt file* rather than as a mis-set
 * option — so the cost of letting one through is a deck that will not open, attributed to
 * nothing in particular. Both halves therefore check before a value reaches the XML.
 *
 * **Why two policies.** They differ deliberately, and the difference is about where the value
 * came from:
 *
 * - {@link checkEnumOrWarn} (write path) reports and **drops** the attribute, leaving the
 *   renderer on the schema default. An option here comes from a deck being built, and losing one
 *   attribute is better than failing an entire build over it.
 * - {@link checkEnumOrThrow} (read/edit path) **throws**. A value here comes from a caller
 *   editing one attribute of a loaded deck, so silently doing nothing would leave them looking at
 *   an unchanged file with no explanation.
 *
 * Before this module the two policies were spread across three near-identical private validators
 * (two in `gen/`, one in `read/`), each with its own wording; adding an enum meant writing a
 * fourth. The policy split is worth keeping — the triplication was not.
 */

import type { DiagnosticCode, InvalidOptionErrorCode } from '../codes.js'
import { warnOnce } from '../diagnostics.js'
import { InvalidOptionError } from '../errors.js'

/**
 * Check a value against its `ST_` union, reporting and dropping it when outside (the write-path
 * policy — see the module note).
 *
 * `undefined`/`null` returns `null` without reporting: an unset option is not an invalid one, and
 * it is by far the common case.
 *
 * The offending value is interpolated into the message deliberately. {@link warnOnce} dedups on
 * code **and** message, so keeping the value in the text is what lets a second, *different* bad
 * value under the same code still get reported.
 * @param value - the value the caller supplied, if any
 * @param valid - the enum's members
 * @param code - the condition to report
 * @param label - how to name the offending option in the message, e.g. ``'table cell: cell3D `preset`'``
 * @param detail - extra structured context to merge into the diagnostic's `detail`
 * @returns the value when legal, else `null`
 */
export function checkEnumOrWarn<T extends string>(
	value: string | undefined | null,
	valid: readonly T[],
	code: DiagnosticCode,
	label: string,
	detail?: Readonly<Record<string, unknown>>
): T | null {
	if (value === undefined || value === null) return null
	if ((valid as readonly string[]).includes(value)) return value as T
	warnOnce(code, `${label} value \`${String(value)}\` is not valid and is ignored — use one of ${valid.join(', ')}.`, {
		received: value,
		valid,
		...detail,
	})
	return null
}

/**
 * Check a value against its `ST_` union, throwing when outside (the read/edit-path policy — see
 * the module note).
 * @param value - the value the caller asked for
 * @param valid - the enum's members
 * @param attribute - the attribute's name, for the message
 * @param code - the condition to report
 * @returns `value`, when it is legal
 */
export function checkEnumOrThrow<T extends string>(
	value: string,
	valid: readonly T[],
	attribute: string,
	code: InvalidOptionErrorCode
): T {
	if ((valid as readonly string[]).includes(value)) return value as T
	throw new InvalidOptionError(
		code,
		`Invalid ${attribute}: ${JSON.stringify(value)}. Expected one of: ${valid.join(', ')}.`
	)
}
