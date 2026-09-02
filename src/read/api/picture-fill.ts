/**
 * Shared picture-fill reader for DrawingML fill-bearing containers: a shape's
 * `p:spPr`, a table cell's `a:tcPr`, and a slide's `p:bg/p:bgPr`. Decodes an
 * `a:blipFill` into the embedded image reference plus the geometry that decides
 * how the image covers the surface — stretch vs tile, the source crop, and the
 * tile offsets/scale.
 *
 * A picture fill is not a colour, so the solid-fill accessors
 * (`AutoShape.resolvedFill`, `TableCell.resolvedFill`) report `null` for one and
 * a surface filled with an image reads as unfilled without this.
 */
import { attr, boolValue, firstChild, numberValue, pctAttr, type Element } from '../oxml/dom.js'
import type { Relationships } from '../opc/relationships.js'

/** A rectangle expressed as per-edge fractions (`a:srcRect`/`a:fillRect`), `0.1` = 10 %. */
export interface FillRect {
	left: number
	top: number
	right: number
	bottom: number
}

/**
 * The tiling geometry of a tiled picture fill (`a:blipFill/a:tile`). Offsets stay
 * in EMU (the attribute's own unit); scales are fractions (`1` = 100 %), matching
 * the read API's convention elsewhere.
 */
export interface PictureFillTile {
	/** Horizontal tile offset in EMU (`@tx`), `0` when unset. */
	offsetXEmu: number
	/** Vertical tile offset in EMU (`@ty`), `0` when unset. */
	offsetYEmu: number
	/** Horizontal tile scale as a fraction (`@sx`, thousandths of a percent), `1` when unset. */
	scaleX: number
	/** Vertical tile scale as a fraction (`@sy`), `1` when unset. */
	scaleY: number
	/** Tile mirroring (`@flip`: `none`/`x`/`y`/`xy`), or `null` when unset. */
	flip: string | null
	/** Tile alignment within the surface (`@algn`, e.g. `tl`), or `null` when unset. */
	align: string | null
}

/**
 * A picture (image) fill (`a:blipFill`) read from a fill-bearing container. The
 * image itself is {@link relId} / {@link partName}; the rest is the geometry a
 * faithful replica needs — {@link mode} (`stretch` or `tile`), the source crop
 * {@link srcRect}, the destination inset {@link fillRect}, and {@link tile}.
 *
 * Rect fields follow the read API's fraction convention (see
 * {@link import('./shapes.js').Picture.crop}): the raw attributes are
 * thousandths of a percent and are divided by 100000, so `0.1` is 10 % and a
 * negative value (a `fillRect` that bleeds past the edge) stays negative.
 */
export interface PictureFill {
	/** Relationship id of the embedded image (`a:blip/@r:embed`), or `null` when the blip embeds none. */
	relId: string | null
	/** Absolute partname of the embedded image, or `null` when it has no rel id / no resolvable target. */
	partName: string | null
	/** `stretch` (`a:stretch`) or `tile` (`a:tile`); `null` when the fill states neither. */
	mode: 'stretch' | 'tile' | null
	/** Source crop as per-edge fractions (`a:srcRect`), or `null` when absent. An explicit `<a:srcRect/>` reports zeros. */
	srcRect: FillRect | null
	/** Stretch destination rect as per-edge fractions (`a:stretch/a:fillRect`), or `null` when absent. */
	fillRect: FillRect | null
	/** Tiling geometry (`a:tile`), or `null` when the fill is not tiled. */
	tile: PictureFillTile | null
	/** Blip opacity 0–1 (`a:blip/a:alphaModFix/@amt`), or `null` when the blip sets none. */
	alpha: number | null
	/** DPI override the image is rendered at (`@dpi`), or `null` when unset. `0` means "use the image's own". */
	dpi: number | null
	/** Whether the fill rotates with its shape (`@rotWithShape`), or `null` when unset. */
	rotWithShape: boolean | null
}

/** Per-edge fractions of a rect element, absent attributes reading as `0`. */
function readRect(rect: Element): FillRect {
	const edge = (name: string): number => pctAttr(rect, name) ?? 0
	return { left: edge('l'), top: edge('t'), right: edge('r'), bottom: edge('b') }
}

/** Decode an `a:tile`: offsets stay in EMU, scales become fractions defaulting to `1`. */
function readTile(tile: Element): PictureFillTile {
	const scale = (name: string): number => pctAttr(tile, name) ?? 1
	return {
		offsetXEmu: numberValue(attr(tile, 'tx')) ?? 0,
		offsetYEmu: numberValue(attr(tile, 'ty')) ?? 0,
		scaleX: scale('sx'),
		scaleY: scale('sy'),
		flip: attr(tile, 'flip'),
		align: attr(tile, 'algn'),
	}
}

/** Resolve a rel id to an absolute partname, `null` for an absent/external/dangling one. */
function resolvePartName(relId: string | null, rels: Relationships | null): string | null {
	if (!relId || !rels) return null
	const rel = rels.get(relId)
	if (!rel || rel.targetMode === 'External') return null
	return rels.resolveTarget(relId)
}

/**
 * Read an `a:blipFill` from a container (`p:spPr`, `a:tcPr`, or `p:bgPr`) into a
 * {@link PictureFill}. `rels` — the *owning part's* relationships — resolves the
 * blip's `r:embed` to an absolute partname; pass `null` (or omit) to read the
 * geometry without resolving the image. `null` when the container's fill is not
 * a picture.
 */
export function readPictureFill(container: Element, rels: Relationships | null = null): PictureFill | null {
	const blipFill = firstChild(container, 'a:blipFill')
	if (!blipFill) return null

	const blip = firstChild(blipFill, 'a:blip')
	const relId = blip ? attr(blip, 'r:embed') : null
	const alphaMod = blip && firstChild(blip, 'a:alphaModFix')
	const amt = alphaMod ? pctAttr(alphaMod, 'amt') : null

	const srcRect = firstChild(blipFill, 'a:srcRect')
	const stretch = firstChild(blipFill, 'a:stretch')
	const fillRect = stretch && firstChild(stretch, 'a:fillRect')
	const tile = firstChild(blipFill, 'a:tile')

	return {
		relId,
		partName: resolvePartName(relId, rels),
		mode: stretch ? 'stretch' : tile ? 'tile' : null,
		srcRect: srcRect ? readRect(srcRect) : null,
		fillRect: fillRect ? readRect(fillRect) : null,
		tile: tile ? readTile(tile) : null,
		// `a:alphaModFix` omits `amt` to mean fully opaque (the schema default).
		alpha: alphaMod ? (amt ?? 1) : null,
		dpi: numberValue(attr(blipFill, 'dpi')),
		rotWithShape: boolValue(attr(blipFill, 'rotWithShape')),
	}
}
