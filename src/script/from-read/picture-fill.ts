/**
 * `a:blipFill` on a *surface* → the write API's `fill: { type: 'image', image: { data } }`.
 *
 * Shared by a shape's fill and a table cell's, which are the same construct read from two
 * containers: `TableCellProps.fill` and `ShapeProps.fill` are both `ShapeFillProps`, so
 * the mapping is identical and only the note wording differs. Keeping one implementation
 * is what stops the two drifting into different answers for the same `a:blipFill`.
 *
 * The bytes and the blip's opacity carry. Nothing else does, and that is a property of the
 * write path rather than of this mapper: `genXmlImageFill` emits every picture fill as
 * `<a:blipFill dpi="0" rotWithShape="1">…<a:srcRect/><a:stretch><a:fillRect/></a:stretch>`,
 * so a tiled fill comes back stretched and a cropped one comes back whole. That is a
 * visible change, so it is noted when — and only when — the source actually uses it.
 */
import type { FillRect, PictureFill } from '../../read/api/picture-fill.js'
import type { NoteScope } from '../fidelity.js'
import type { AssetRef, IrValue } from '../ir.js'
import type { AssetResolver } from './shape.js'
import { alphaToTransparency, compactRequired } from './values.js'

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

	noteUncarriedGeometry(picture, notes, subject)

	return compactRequired({
		type: 'image',
		image: { data: asset },
		// Fully opaque passes as `undefined` rather than `0`; see `alphaToTransparency`.
		transparency: picture.alpha === 1 ? undefined : alphaToTransparency(picture.alpha),
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
 * Note the parts of the fill's geometry the write path's fixed stretch cannot express.
 *
 * Only the ones the source actually uses: a note on every picture fill would say nothing,
 * and the round-trip check reads an unmatched note as a candidate for review.
 */
function noteUncarriedGeometry(picture: PictureFill, notes: NoteScope, subject: PictureFillSubject): void {
	const lost: string[] = []
	if (picture.mode === 'tile' || picture.tile !== null) lost.push('tiling (a:tile)')
	if (isRectSet(picture.srcRect)) lost.push('the source crop (a:srcRect)')
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
		`the image filling ${subject.subject} is re-embedded, but the write path emits every picture fill as a plain stretched blip (dpi="0" rotWithShape="1", <a:srcRect/><a:stretch><a:fillRect/></a:stretch>), so ${lost.join(', ')} ${lost.length === 1 ? 'does' : 'do'} not carry`
	)
}
