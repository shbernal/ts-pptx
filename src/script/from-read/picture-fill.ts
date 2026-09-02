/**
 * `a:blipFill` on a *surface* → the write API's `fill: { type: 'image', image: { data } }`.
 *
 * Shared by a shape's fill and a table cell's, which are the same construct read from two
 * containers: `TableCellProps.fill` and `ShapeProps.fill` are both `ShapeFillProps`, so
 * the mapping is identical and only the note wording differs. Keeping one implementation
 * is what stops the two drifting into different answers for the same `a:blipFill`.
 *
 * The bytes, the blip's opacity and the source crop carry. Nothing else does, and that is a
 * property of the write path rather than of this mapper: `genXmlImageFill` emits every
 * picture fill as `<a:blipFill dpi="0" rotWithShape="1">…<a:stretch><a:fillRect/></a:stretch>`
 * with only its `<a:srcRect>` under the caller's control, so a tiled fill still comes back
 * stretched. That is a visible change, so it is noted when — and only when — the source
 * actually uses it.
 */
import type { FillRect, PictureFill } from '../../read/api/picture-fill.js'
import type { NoteScope } from '../fidelity.js'
import type { AssetRef, IrValue } from '../ir.js'
import type { AssetResolver } from './context.js'
import { FIXED_PCT_PER_PERCENT, PERCENT_SCALE } from '../../units.js'
import { alphaToTransparency, compact, compactRequired } from './values.js'

/** How a note names the surface it is about; the mapping itself does not vary. */
export interface PictureFillSubject {
	/** Note construct for a fill whose bytes cannot be carried at all. */
	construct: string
	/** The surface as a note's detail names it — "this shape's surface", "this table cell". */
	subject: string
	/** Where the fill sits in the source XML, for a note's detail. */
	element: string
}

/**
 * A picture fill as `ShapeFillProps`, or `undefined` when its bytes cannot be carried —
 * in which case a `dropped` note has been recorded and the surface comes out unfilled.
 */
export function pictureFillOption(
	picture: PictureFill,
	assets: AssetResolver,
	notes: NoteScope,
	subject: PictureFillSubject
): IrValue | undefined {
	const asset = pictureAsset(picture, assets, notes, subject)
	if (!asset) return undefined

	const crop = cropOption(picture.srcRect)
	noteUncarriedGeometry(picture, notes, subject, crop !== undefined)

	return compactRequired({
		type: 'image',
		image: compactRequired({ crop, data: asset }),
		transparency: alphaToTransparency(picture.alpha),
	})
}

/** The fill's image as an asset, or `null` with a note saying why it could not be one. */
function pictureAsset(
	picture: PictureFill,
	assets: AssetResolver,
	notes: NoteScope,
	subject: PictureFillSubject
): AssetRef | null {
	const partName = picture.partName
	if (partName === null) {
		notes.note(
			subject.construct,
			'dropped',
			'unsupported',
			`${subject.subject} is filled with an image (${subject.element}) that embeds no image part — an external or linked blip — so there are no bytes to carry and it comes out unfilled`
		)
		return null
	}

	// The write path refuses an SVG picture fill outright (`src/gen/define/image.ts` warns
	// and falls back to `type: 'none'`), so emitting one would produce a script that paints
	// nothing while claiming a fill. Checked *before* `assetFor`, which would otherwise
	// register bytes that no call ends up referencing.
	if (assets.contentTypeOf(partName) === 'image/svg+xml') {
		notes.note(
			subject.construct,
			'dropped',
			'unwritable',
			`${subject.subject} is filled with an SVG image (${subject.element} → ${partName}); a picture fill accepts raster bytes only on the write side, so it comes out unfilled`
		)
		return null
	}

	const asset = assets.assetFor(partName)
	if (!asset) {
		notes.note(
			subject.construct,
			'dropped',
			'unsupported',
			`the image filling ${subject.subject} (${subject.element} → ${partName}) is not present in the package, so there are no bytes to carry`
		)
		return null
	}
	return asset
}

/** `true` when a rect actually insets an edge; an explicit but empty rect reports zeros. */
function isRectSet(rect: FillRect | null): boolean {
	return rect !== null && (rect.left !== 0 || rect.top !== 0 || rect.right !== 0 || rect.bottom !== 0)
}

/**
 * The source crop as the write API's `image.crop`, or `undefined` when this rect cannot be
 * one — in which case the caller notes it as uncarried instead.
 *
 * An all-zero rect is `undefined` rather than four zeros: `<a:srcRect/>` is what the write
 * path emits with no `crop` at all, so carrying zeros would add an option that changes
 * nothing and take the fill off its byte-identical path for no reason.
 *
 * Two shapes of rect are legal in a source and not expressible as `crop`, and both belong
 * to the write path rather than to this mapper (`gen/drawingml/src-rect.ts` throws on
 * each): a **negative** inset, which is how a `contain`-style fill bleeds the source past
 * the surface, and a pair of opposite insets summing to 100% or more, which leaves no
 * source area. Carrying either would emit a script that throws when it is run, so they
 * stay uncarried and keep their note.
 */
function cropOption(srcRect: FillRect | null): IrValue | undefined {
	if (!isRectSet(srcRect) || srcRect === null) return undefined
	// `readRect` divided the source's thousandths-of-a-percent by PERCENT_SCALE; undo exactly
	// that, so the integer the writer re-derives is the one the source held and the crop
	// survives the trip byte for byte rather than to within a rounding step.
	const pct = (fraction: number): number => Math.round(fraction * PERCENT_SCALE) / FIXED_PCT_PER_PERCENT
	const edges = { b: pct(srcRect.bottom), l: pct(srcRect.left), r: pct(srcRect.right), t: pct(srcRect.top) }
	if (Object.values(edges).some((v) => v < 0 || v > 100)) return undefined
	if (edges.l + edges.r >= 100 || edges.t + edges.b >= 100) return undefined
	// A zero edge is the option's own default, so it is left out rather than spelled.
	return compact({
		b: edges.b || undefined,
		l: edges.l || undefined,
		r: edges.r || undefined,
		t: edges.t || undefined,
	})
}

/**
 * Note the parts of the fill's geometry the write path's fixed stretch cannot express.
 *
 * Only the ones the source actually uses: a note on every picture fill would say nothing,
 * and the round-trip check reads an unmatched note as a candidate for review. `cropCarried`
 * is the one that is conditional on more than the source — the crop is expressible as
 * `image.crop` unless its insets fall outside what that option accepts, so whether it is
 * lost is the caller's answer to give.
 */
function noteUncarriedGeometry(
	picture: PictureFill,
	notes: NoteScope,
	subject: PictureFillSubject,
	cropCarried: boolean
): void {
	const lost: string[] = []
	if (picture.mode === 'tile' || picture.tile !== null) lost.push('tiling (a:tile)')
	if (isRectSet(picture.srcRect) && !cropCarried)
		lost.push('the source crop (a:srcRect), whose insets fall outside the 0–100% `image.crop` accepts')
	if (isRectSet(picture.fillRect)) lost.push('the destination inset (a:stretch/a:fillRect)')
	if (picture.dpi !== null && picture.dpi !== 0) lost.push(`the render resolution (@dpi="${picture.dpi}")`)
	// Only the explicit `0`. Unset means "whatever the application defaults to", which is
	// not a claim this note can make about a difference.
	if (picture.rotWithShape === false) lost.push('not rotating with the shape (@rotWithShape="0")')
	if (lost.length === 0) return

	notes.note(
		`${subject.construct}.geometry`,
		'approximated',
		'unwritable',
		`the image filling ${subject.subject} is re-embedded, but the write path emits every picture fill as a stretched blip at a fixed dpi="0" rotWithShape="1", with only its <a:srcRect> under the caller's control, so ${lost.join(', ')} ${lost.length === 1 ? 'does' : 'do'} not carry`
	)
}
