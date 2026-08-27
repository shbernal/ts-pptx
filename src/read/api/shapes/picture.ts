/**
 * A picture (`p:pic`).
 *
 * A picture's drawn data is its sibling `p:blipFill`, not a fill of `p:spPr`, which is why the
 * blip plumbing (content-type → extension, the SVG extension blip, the recolour effects) lives
 * here rather than on the base: no other shape kind carries it.
 */

import {
	ELEMENT_NODE,
	attr,
	firstChild,
	firstChildElement,
	getElements,
	getOrAddChild,
	intValue,
	removeChildrenByQName,
	setAttr,
	OOXML_NS,
	type Element,
} from '../../oxml/dom.js'
import { fitSrcRectPercents, getImageSizeFromBytes } from '../../../media/image-size.js'
import { warn } from '../../../diagnostics.js'
import { relativePartName } from '../../opc/partnames.js'
import { Shape } from './base.js'
import { childElements } from './oxml.js'
import type { Recolor, RecolorColor } from './types.js'
import { IMAGE_REL } from '../../../ooxml/rel-types.js'
import { InvalidOptionError } from '../../../errors.js'
import { PERCENT_SCALE } from '../../../units.js'

// Microsoft's SVG blip extension namespace (a:blip/a:extLst/a:ext/asvg:svgBlip).
const ASVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'

// Schema successors within p:pic (CT_Picture: nvPicPr, blipFill, spPr, style?)
// and within a:blipFill (blip?, srcRect?, (tile|stretch)?), used to keep a
// get-or-added p:blipFill / a:blip in document order.
const PIC_AFTER_BLIPFILL = ['p:spPr', 'p:style']
const BLIPFILL_AFTER_BLIP = ['a:srcRect', 'a:tile', 'a:stretch']

/** Known content-type → file-extension map for image media parts. */
const IMAGE_EXTENSION_BY_CONTENT_TYPE: Readonly<Record<string, string>> = Object.freeze({
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
	'image/webp': 'webp',
	'image/svg+xml': 'svg',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
})

/**
 * Default a media-part file extension from a content type. Known image types use
 * an explicit map; otherwise fall back to the content-type subtype (before any
 * `+suffix`, with a leading `x-` stripped), e.g. `image/x-foo` → `foo`.
 */
function extFromContentType(contentType: string): string {
	const known = IMAGE_EXTENSION_BY_CONTENT_TYPE[contentType.toLowerCase()]
	if (known) return known
	const subtype = contentType.toLowerCase().split('/')[1] ?? ''
	const ext = (subtype.split('+')[0] ?? '').replace(/^x-/, '')
	if (!ext)
		throw new InvalidOptionError(
			'image/undeterminable-extension',
			`Cannot derive a file extension from content type "${contentType}"; pass { extension }`
		)
	return ext
}

/** Convert a DrawingML colour element to a {@link RecolorColor}, or `null` when it is not an `a:` colour element. */
function recolorColorOf(color: Element | null): RecolorColor | null {
	if (!color || color.namespaceURI !== OOXML_NS.a) return null
	return {
		color: color.localName === 'srgbClr' ? attr(color, 'val') : null,
		schemeColor: color.localName === 'schemeClr' ? attr(color, 'val') : null,
		presetColor: color.localName === 'prstClr' ? attr(color, 'val') : null,
	}
}

/** A picture (`p:pic`). */
export class Picture extends Shape {
	readonly shapeType = 'picture' as const

	// A picture's image is its sibling `p:blipFill`, not a fill of `p:spPr`, so a
	// solid `spPr` fill would not clobber the image. v1 still omits fill setters
	// here — recolouring a picture surface is rarely what a caller means — and
	// exposes only the border via `lineColor`. Reads of `fillColor` stay valid.
	protected override get supportsFill(): boolean {
		return false
	}

	/** Relationship id of the embedded image (`p:blipFill/a:blip/@r:embed`), or `null`. */
	get imageRelId(): string | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		return blip ? attr(blip, 'r:embed') : null
	}

	/**
	 * Repoint the blip at a relationship id already present in the owning part's
	 * relationships, without minting a new media part. The caller owns ensuring
	 * the id exists and targets an image; use {@link setImage} to add fresh bytes.
	 */
	set imageRelId(value: string) {
		setAttr(this.#getOrAddBlip(), 'r:embed', value)
		this.markDirty()
	}

	/** Absolute partname of the embedded image, resolved via the owning part's relationships, or `null`. */
	get imagePartName(): string | null {
		const relId = this.imageRelId
		return relId ? this.host.relationships.resolveTarget(relId) : null
	}

	/**
	 * Relationship id of the embedded **vector (SVG)** image, read from the
	 * Microsoft SVG blip extension (`a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed`),
	 * or `null` when the picture has no SVG. PowerPoint usually pairs this with a
	 * raster fallback in `a:blip/@r:embed` ({@link imageRelId}), but some exporters
	 * emit an SVG-only blip where `imageRelId` is absent and only this resolves —
	 * so a reader that wants the real drawn art must consult both.
	 */
	get svgRelId(): string | null {
		const svg = this.#svgBlip()
		return svg ? attr(svg, 'r:embed') : null
	}

	/** Absolute partname of the embedded SVG image, resolved via the owning part's relationships, or `null`. */
	get svgPartName(): string | null {
		const relId = this.svgRelId
		return relId ? this.host.relationships.resolveTarget(relId) : null
	}

	/**
	 * Which drawable media this picture carries:
	 * - `'raster'` — only a raster blip (`a:blip/@r:embed`);
	 * - `'svg'` — only a vector blip (`asvg:svgBlip/@r:embed`, no raster). This
	 *   is what PowerPoint's *Insert → Icons* and a plain SVG insert produce;
	 * - `'both'` — a raster fallback *and* an SVG (PowerPoint's usual pairing);
	 * - `'none'` — a `p:pic` with no embedded blip at all (e.g. a linked image).
	 *
	 * Lets a caller distinguish an SVG-only picture — where {@link imagePartName}
	 * is legitimately `null` — from a genuinely empty one, without two null checks.
	 */
	get mediaKind(): 'raster' | 'svg' | 'both' | 'none' {
		const hasRaster = this.imageRelId != null
		const hasSvg = this.svgRelId != null
		if (hasRaster && hasSvg) return 'both'
		if (hasRaster) return 'raster'
		if (hasSvg) return 'svg'
		return 'none'
	}

	/**
	 * Absolute partname of whichever part actually carries this picture's drawn
	 * data — the raster part when present, otherwise the SVG part — or `null`
	 * when the picture embeds neither. Use this when you just want "the bytes
	 * this picture shows"; prefer {@link imagePartName} / {@link svgPartName}
	 * (and {@link mediaKind}) when you need to know *which* kind it is. An
	 * SVG-only picture returns its SVG part here even though `imagePartName` is
	 * `null`.
	 */
	get mediaPartName(): string | null {
		return this.imagePartName ?? this.svgPartName
	}

	/**
	 * The picture's crop as fractions of the *source image*, read from
	 * `p:blipFill/a:srcRect` — `{ left, top, right, bottom }`, each the amount
	 * trimmed off that edge (so `0.1` = 10 % cropped away, an uncropped edge is
	 * `0`). `null` when there is no `a:srcRect` at all; an explicit
	 * `{0,0,0,0}` crop still reports zeros, since its presence is meaningful.
	 * The raw attributes are thousandths of a percent; this divides by 100000 to
	 * match the fraction convention used elsewhere in the read API (see
	 * {@link recolor}).
	 */
	get crop(): { left: number; top: number; right: number; bottom: number } | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const srcRect = blipFill && firstChild(blipFill, 'a:srcRect')
		if (!srcRect) return null
		const edge = (name: string): number => {
			const v = intValue(attr(srcRect, name))
			return v === null ? 0 : v / PERCENT_SCALE
		}
		return { left: edge('l'), top: edge('t'), right: edge('r'), bottom: edge('b') }
	}

	/**
	 * The picture's blip recolour effect (`p:blipFill/a:blip` recolour child), or
	 * `null` when the blip carries none. Recognises the effects a faithful reader
	 * needs to reproduce a recoloured image: `a:duotone` (the two-stop icon-tint
	 * trick), `a:clrChange`, `a:grayscl`, `a:biLevel`, and `a:alphaModFix`; the
	 * first such effect in document order wins. Colours mirror the
	 * {@link GradientStop} split (`color`/`schemeColor`/`presetColor`) so theme
	 * tokens resolve through {@link Slide.themeContext}. `threshold`/`amount` are
	 * 0–1 fractions. {@link hidden} (the duotone fallback-layer trick) reports the
	 * *visibility* of a recolour source; this reports the *tint* itself.
	 */
	get recolor(): Recolor | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		if (!blip) return null
		for (const child of childElements(blip)) {
			if (child.namespaceURI !== OOXML_NS.a) continue
			switch (child.localName) {
				case 'duotone':
					return {
						kind: 'duotone',
						stops: childElements(child)
							.map(recolorColorOf)
							.filter((c): c is RecolorColor => c !== null),
					}
				case 'clrChange': {
					const from = firstChild(child, 'a:clrFrom')
					const to = firstChild(child, 'a:clrTo')
					return {
						kind: 'clrChange',
						from: from ? recolorColorOf(firstChildElement(from)) : null,
						to: to ? recolorColorOf(firstChildElement(to)) : null,
					}
				}
				case 'grayscl':
					return { kind: 'grayscale' }
				case 'biLevel': {
					const thresh = intValue(attr(child, 'thresh'))
					return { kind: 'biLevel', threshold: thresh === null ? null : thresh / PERCENT_SCALE }
				}
				case 'alphaModFix': {
					const amt = intValue(attr(child, 'amt'))
					return { kind: 'alphaModFix', amount: amt === null ? 1 : amt / PERCENT_SCALE }
				}
			}
		}
		return null
	}

	/** The `<asvg:svgBlip>` element inside the blip's extLst, or `null` when the picture carries no SVG. */
	#svgBlip(): Element | null {
		const blipFill = firstChild(this.element, 'p:blipFill')
		const blip = blipFill && firstChild(blipFill, 'a:blip')
		const extLst = blip && firstChild(blip, 'a:extLst')
		if (!extLst) return null
		for (const ext of getElements(extLst, 'a:ext')) {
			for (let node = ext.firstChild; node; node = node.nextSibling) {
				if (node.nodeType !== ELEMENT_NODE) continue
				const el = node as Element
				if (el.localName === 'svgBlip' && el.namespaceURI === ASVG_NS) return el
			}
		}
		return null
	}

	/**
	 * Replace this picture's image with new bytes. Mints a fresh media part under
	 * `/ppt/media/`, registers its content type, wires an `image` relationship
	 * from the owning part, and repoints the blip's `@r:embed` at it.
	 *
	 * Copy-on-write: the previous media part is never mutated or removed, so any
	 * other picture sharing it (common after `importSlide`/dedup) is unaffected;
	 * an orphaned old part is left in place for a later GC pass to prune.
	 *
	 * `contentType` is required (e.g. `image/png`); the bytes are not sniffed.
	 * `extension` defaults from the content type.
	 *
	 * `fit` controls the picture's `a:srcRect` crop against its current frame
	 * extent (`a:xfrm/a:ext`):
	 * - omitted (default): geometry and crop are left untouched — the caller owns
	 *   sizing. Note an inherited `a:srcRect` was tuned to the *previous* image's
	 *   aspect ratio, so swapping in an image of a different ratio reuses a crop
	 *   that no longer fits and the result looks stretched; pass `fit` to refit.
	 * - `'cover'`: fill the frame, cropping the overflowing axis (no distortion).
	 * - `'contain'`: fit the whole image inside the frame, letterboxing the short
	 *   axis (no distortion).
	 * - `'stretch'`: drop any crop so the full image is stretched to the frame.
	 *
	 * `'cover'`/`'contain'` measure the new bytes' natural size; if unmeasurable
	 * (e.g. an unknown format) the crop is left as-is and a warning is emitted.
	 */
	setImage(
		bytes: Uint8Array,
		options: { contentType: string; extension?: string; fit?: 'cover' | 'contain' | 'stretch' }
	): void {
		const { contentType } = options
		if (!contentType)
			throw new InvalidOptionError('image/missing-content-type', 'setImage requires a contentType (e.g. "image/png")')
		const extension = (options.extension ?? extFromContentType(contentType)).toLowerCase().replace(/^\./, '')

		const opc = this.host.opc
		const mediaPartName = opc.reserveMediaPartName(extension)
		opc.addPart(mediaPartName, contentType, bytes)
		const relId = this.host.relationships.add(IMAGE_REL, relativePartName(this.host.partName, mediaPartName)).id

		setAttr(this.#getOrAddBlip(), 'r:embed', relId)
		if (options.fit) this.#applyFit(options.fit, bytes)
		this.markDirty()
	}

	/**
	 * Refit the blip crop after a {@link setImage} swap. `stretch` removes any
	 * `a:srcRect`; `cover`/`contain` recompute it from the new image's natural
	 * size against the frame extent so the swap is aspect-correct.
	 */
	#applyFit(fit: 'cover' | 'contain' | 'stretch', bytes: Uint8Array): void {
		const blipFill = getOrAddChild(this.element, 'p:blipFill', PIC_AFTER_BLIPFILL)
		if (fit === 'stretch') {
			removeChildrenByQName(blipFill, ['a:srcRect'])
			return
		}
		const natural = getImageSizeFromBytes(bytes)
		if (!natural) {
			warn(
				'image/unmeasurable-natural-size',
				`setImage fit '${fit}': could not measure the new image's natural size; leaving the crop unchanged (it may look stretched). Provide a raster (PNG/JPEG/GIF/BMP/WebP) or an SVG with width/height or a viewBox.`
			)
			return
		}
		const cx = this.width
		const cy = this.height
		if (cx == null || cy == null) {
			throw new InvalidOptionError(
				'image/fit-needs-extent',
				`setImage fit '${fit}' needs a frame extent (a:xfrm/a:ext); this picture has no transform`
			)
		}
		const { l, r, t, b } = fitSrcRectPercents(fit, { w: natural.w, h: natural.h }, { w: cx, h: cy })
		const srcRect = getOrAddChild(blipFill, 'a:srcRect', ['a:tile', 'a:stretch'])
		setAttr(srcRect, 'l', String(l))
		setAttr(srcRect, 'r', String(r))
		setAttr(srcRect, 't', String(t))
		setAttr(srcRect, 'b', String(b))
	}

	/** Get-or-add `p:blipFill/a:blip`, keeping both in document order. */
	#getOrAddBlip(): Element {
		const blipFill = getOrAddChild(this.element, 'p:blipFill', PIC_AFTER_BLIPFILL)
		return getOrAddChild(blipFill, 'a:blip', BLIPFILL_AFTER_BLIP)
	}
}
