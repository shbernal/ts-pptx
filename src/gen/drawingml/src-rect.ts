/**
 * ts-pptx: the explicit `<a:srcRect>` source crop
 *
 * A leaf module on purpose. A picture object (`addImage`) and a native picture fill
 * (`fill: { type: 'image' }`) crop a source image the same way, but they live on opposite
 * sides of a bundling boundary: `gen/drawingml/fill.ts` is reached from every shape, table
 * and chart emitter, while `gen/drawingml/image.ts` pulls the image-format sniffing tables
 * in behind it through `media/image-size.ts`. Sharing the crop from here rather than from
 * `image.ts` is what keeps ~24 kB of format headers out of the `read` and `inspect` entry
 * points, which never measure an image.
 */

import { FIXED_PCT_PER_PERCENT } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { InvalidOptionError } from '../../errors.js'

/** Percentage edge insets trimmed from the source image; an omitted edge is `0`. */
export type ImageCrop = { l?: number; t?: number; r?: number; b?: number }

/**
 * The four insets, defaulted and checked. Rejects anything that is not a percentage and
 * any pair that would leave no source area at all — a degenerate `<a:srcRect>` renders as
 * an empty shape rather than as an error, so it has to fail here.
 * @param crop - edge insets in percent; omitted edges default to 0
 * @param label - the option's owner, as an error message names it (`addImage`, `image fill`)
 * @param where - trailing context for an error message, or `''`
 */
function normalizeImageCrop(
	crop: ImageCrop,
	label: string,
	where: string
): { l: number; t: number; r: number; b: number } {
	const edges = { l: crop.l ?? 0, t: crop.t ?? 0, r: crop.r ?? 0, b: crop.b ?? 0 }
	for (const [name, val] of Object.entries(edges)) {
		if (typeof val !== 'number' || !Number.isFinite(val) || val < 0 || val > 100) {
			throw new InvalidOptionError(
				'image/crop-inset-out-of-range',
				`${label} crop.${name} must be a percentage between 0 and 100 (got ${String(val)})${where}.`
			)
		}
	}
	if (edges.l + edges.r >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`${label} crop: left+right insets (${edges.l}%+${edges.r}%) must be < 100%${where}.`
		)
	if (edges.t + edges.b >= 100)
		throw new InvalidOptionError(
			'image/crop-insets-exceed-extent',
			`${label} crop: top+bottom insets (${edges.t}%+${edges.b}%) must be < 100%${where}.`
		)
	return edges
}

/**
 * Build an explicit `<a:srcRect>` crop from percentage edge insets (0–100), emitted verbatim.
 * Each inset is the percent of the *source* image trimmed from that edge; values serialize in
 * 1000ths of a percent (DrawingML `ST_Percentage`, where 100% = 100000). Unlike `sizing: 'crop'`
 * — which derives the rect from displayed inches and the natural pixel size — this maps a
 * sub-region of the source directly, so it works for SVG and other unmeasurable formats.
 * @param crop - edge insets in percent; omitted edges default to 0
 * @param label - the option's owner, as an error message names it
 * @param where - trailing context for an error message, or `''`
 * @returns the `<a:srcRect>` element on its own; the caller supplies the fill directive
 */
export function genXmlImageCropRect(crop: ImageCrop, label: string, where: string): string {
	const edges = normalizeImageCrop(crop, label, where)
	const v = (perc: number): number => Math.round(perc * FIXED_PCT_PER_PERCENT)
	// NOTE: attribute order here is l/t/r/b, where the sizing modes in `image.ts` emit l/r/t/b.
	// Attribute order is byte-significant, so the two orderings are kept as they are rather
	// than unified.
	return voidEl('a:srcRect', { l: v(edges.l), t: v(edges.t), r: v(edges.r), b: v(edges.b) })
}

/**
 * `<a:stretch><a:fillRect/></a:stretch>` — the blipFill child that scales the source to the
 * shape's box, which every picture-bearing emitter writes and none parameterizes.
 *
 * Here rather than in `image.ts` for the same bundling reason the crop is: `fill.ts` reaches
 * it from every shape, table and chart emitter, and must not pull the image-format tables in.
 */
export const STRETCH_FILL_RECT = el('a:stretch', null, raw(voidEl('a:fillRect')))
