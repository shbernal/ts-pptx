/**
 * Layout arithmetic shared by the showcase decks.
 *
 * Both decks are 16:9 (`LAYOUT_WIDE`, 13.333 x 7.5 inches) and both lay content out on
 * the same 12-column grid, so the numbers that decide where things sit live here once
 * instead of being re-derived (and re-rounded differently) on every slide.
 */

/** `LAYOUT_WIDE` in inches. */
export const WIDE = { w: 13.333, h: 7.5 };

/** Horizontal margin used by every content slide in both decks. */
export const MARGIN = 0.75;

/** Width of the content column between the two margins. */
export const CONTENT_W = WIDE.w - MARGIN * 2;

/**
 * Left edges and width for `count` equal columns spanning the content area.
 *
 * @param {number} count number of columns
 * @param {object} [opts]
 * @param {number} [opts.gap=0.35] gutter between columns, in inches
 * @param {number} [opts.left=MARGIN] left edge of the first column
 * @param {number} [opts.width=CONTENT_W] total width to divide
 * @returns {{ x: number[], w: number }} column left edges and the shared column width
 */
export function columns(count, opts = {}) {
	const { gap = 0.35, left = MARGIN, width = CONTENT_W } = opts;
	const w = (width - gap * (count - 1)) / count;
	const x = Array.from({ length: count }, (_, i) => left + i * (w + gap));
	return { x, w };
}

/**
 * Centre `count` items of width `w` (plus `gap`) horizontally on the slide.
 * Used where a row is narrower than the content column and should sit centred
 * rather than left-aligned — e.g. the KPI band and the contents thumbnails.
 */
export function centeredRow(count, w, gap) {
	const total = count * w + (count - 1) * gap;
	const left = (WIDE.w - total) / 2;
	return Array.from({ length: count }, (_, i) => left + i * (w + gap));
}

/** Format a number as a signed percentage string, e.g. `+12.4%`. */
export function signedPct(value, digits = 1) {
	return `${value >= 0 ? "+" : "−"}${Math.abs(value).toFixed(digits)}%`;
}
