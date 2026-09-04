/**
 * The one shape a measurement takes when it could not be taken.
 *
 * Every network fetch in this directory records `{ unavailable: <reason> }` instead of
 * aborting the run, because a release must not be blockable by a rate limit or an
 * offline afternoon. That leniency needs a counterweight or a snapshot full of holes
 * gets committed by reflex, so {@link findUnavailable} walks the finished snapshot and
 * `measure.mjs` exits non-zero on anything it finds unless `--allow-unavailable` says
 * the holes are intended.
 *
 * The renderer omits an unavailable row rather than printing an empty cell: a blank in a
 * comparison table reads as a measured zero, and there is no honest way to draw the
 * difference between "we asked and got nothing" and "the answer is nothing" in a cell.
 */

/**
 * @typedef {object} Unavailable
 * @property {string} unavailable - why, in a phrase a reader can act on
 */

/**
 * A measurement that could not be taken.
 * @param {string} reason
 * @returns {Unavailable}
 */
export function unavailable(reason) {
	return { unavailable: reason }
}

/**
 * Is this value an {@link Unavailable} marker?
 * @param {unknown} value
 * @returns {value is Unavailable}
 */
export function isUnavailable(value) {
	return (
		typeof value === 'object' &&
		value !== null &&
		typeof (/** @type {{unavailable?: unknown}} */ (value).unavailable) === 'string'
	)
}

/**
 * Every unavailable marker in a snapshot, by the path it sits at.
 *
 * Dotted paths rather than a count, so the failure message names what is missing. A run
 * that reports "3 measurements unavailable" and stops sends the reader back to the JSON
 * to find out which; naming them is the difference between a gate and an obstacle.
 * @param {unknown} value - any node of the snapshot
 * @param {string} [at] - the dotted path of `value`
 * @returns {Array<{path: string, reason: string}>}
 */
export function findUnavailable(value, at = '') {
	if (isUnavailable(value)) return [{ path: at || '(root)', reason: value.unavailable }]
	if (Array.isArray(value)) return value.flatMap((item, index) => findUnavailable(item, at + '[' + index + ']'))
	if (typeof value === 'object' && value !== null)
		return Object.entries(value).flatMap(([key, item]) => findUnavailable(item, at ? at + '.' + key : key))
	return []
}
