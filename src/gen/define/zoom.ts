/**
 * ts-pptx: Zoom Definition (Slide / Section / Summary Zoom — Insert ▸ Zoom).
 *
 * Resolves a zoom's target(s) to the ids the emitter needs (`sldId` / section GUID), registers
 * the preview-image media rel and the `.../slide` fallback rel(s), lays out the Summary Zoom grid,
 * and pushes a `SlideObject{ _type: zoom, zoom }` for `gen/slide/objects/zoom.ts` to emit as a
 * `<p:graphicFrame>`. See {@link ../../types/zoom} for the preview-image behavior.
 */
import { SlideObjectType } from '../../core-enums.js'
import { warn } from '../../diagnostics.js'
import type { SectionZoomProps, SlideZoomProps, SummaryZoomProps } from '../../types/zoom.js'
import type {
	PresSlideInternal,
	SectionInternalProps,
	SlideObject,
	ZoomInternal,
	ZoomTileInternal,
} from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, getUuid, validateObjectName } from '../../gen-utils.js'
import { getSmartParseNumber } from '../../units-internal.js'
import { nextObjectNameIdx } from './object-name.js'
import { registerPreviewImage } from './preview-image.js'

const ZOOM_LABEL = { slide: 'Slide Zoom', section: 'Section Zoom', summary: 'Summary Zoom' } as const

/** A fresh, braced, upper-case v4 GUID for a `zmPr@id`. */
function zoomGuid(): string {
	return `{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}`
}

/** Register a `.../slide` rel (used by the fallback picture's `hlinkClick`) to a 1-based slide number. */
function registerSlideRel(target: PresSlideInternal, slideNum: number): number {
	const rId = getNewRelId(target)
	target._rels.push({ type: SlideObjectType.hyperlink, data: 'slide', rId, Target: String(slideNum) })
	return rId
}

/** Shared object-name + object-scaffold for all three variants. */
function pushZoomObject(
	target: PresSlideInternal,
	variant: ZoomInternal['variant'],
	opts: SlideZoomProps | SectionZoomProps | SummaryZoomProps,
	zoom: Omit<ZoomInternal, 'variant'>
): void {
	const nameIdx = nextObjectNameIdx(target, SlideObjectType.zoom)
	const objectName = opts.objectName
		? encodeXmlAttrValue(validateObjectName(opts.objectName, 'zoom'))
		: `${ZOOM_LABEL[variant]} ${nameIdx + 1}`
	const newObject: SlideObject = {
		_type: SlideObjectType.zoom,
		options: { x: opts.x ?? 0, y: opts.y ?? 0, w: opts.w ?? 0, h: opts.h ?? 0, objectName },
		zoom: { variant, ...zoom },
	}
	target._slideObjects.push(newObject)
}

/** Slide Zoom — one tile linking to a single target slide. */
export function addSlideZoomDefinition(target: PresSlideInternal, opts: SlideZoomProps): void {
	if (!opts?.target) {
		warn(
			'zoom/missing-target',
			'addSlideZoom requires a `target` slide (a Slide object or its 1-based number); ignoring.'
		)
		return
	}
	const targetSlide = opts.target as PresSlideInternal
	const sldId = typeof opts.target === 'number' ? 256 + (opts.target - 1) : targetSlide._slideId
	const slideNum = typeof opts.target === 'number' ? opts.target : targetSlide._slideNum
	if (sldId == null || slideNum == null) {
		warn('zoom/unresolved-target', 'addSlideZoom: could not resolve the target slide; ignoring.')
		return
	}

	const previewRid = registerPreviewImage(target, opts.coverImage)
	const fallbackSlideRid = registerSlideRel(target, slideNum)
	const tile: ZoomTileInternal = { sldId, previewRid, fallbackSlideRid, zmPrId: zoomGuid() }
	pushZoomObject(target, 'slide', opts, {
		tiles: [tile],
		returnToParent: !!opts.returnToParent,
		transitionDur: opts.transitionDur ?? 1000,
	})
}

/** Section Zoom — one tile linking to the start of a named section. */
export function addSectionZoomDefinition(
	target: PresSlideInternal,
	opts: SectionZoomProps,
	sections: SectionInternalProps[]
): void {
	if (!opts?.sectionTitle) {
		warn('zoom/missing-section-title', 'addSectionZoom requires a `sectionTitle`; ignoring.')
		return
	}
	const section = sections.find((s) => s.title === opts.sectionTitle)
	if (!section) {
		warn('zoom/section-not-found', `addSectionZoom: no section titled "${opts.sectionTitle}"; ignoring.`)
		return
	}
	const firstSlide = section._slides[0]
	if (!firstSlide) {
		warn('zoom/section-empty', `addSectionZoom: section "${opts.sectionTitle}" has no slides; ignoring.`)
		return
	}

	const previewRid = registerPreviewImage(target, opts.coverImage)
	const fallbackSlideRid = registerSlideRel(target, firstSlide._slideNum)
	const tile: ZoomTileInternal = { sectionId: section._id, previewRid, fallbackSlideRid, zmPrId: zoomGuid() }
	pushZoomObject(target, 'section', opts, {
		tiles: [tile],
		returnToParent: !!opts.returnToParent,
		transitionDur: opts.transitionDur ?? 1000,
	})
}

/**
 * Summary Zoom — a grid of tiles, one per section (excluding the host slide's own section).
 * All tiles share the single placeholder/cover image; each links to its section's first slide.
 */
export function addSummaryZoomDefinition(
	target: PresSlideInternal,
	opts: SummaryZoomProps,
	sections: SectionInternalProps[]
): void {
	// Exclude the section that contains this (host) slide — a summary does not link to itself.
	const hostSection = sections.find((s) => s._slides.some((sl) => sl._slideNum === target._slideNum))
	const targets = sections.filter((s) => s !== hostSection && s._slides.length > 0)
	if (targets.length === 0) {
		warn(
			'zoom/no-sections-to-summarize',
			'addSummaryZoom: no sections to summarize (need at least one section besides this slide’s own); ignoring.'
		)
		return
	}

	// Grid geometry (EMU). Tiles preserve the slide aspect ratio; the grid is centered in the frame,
	// last row left-aligned (matching PowerPoint). See the Summary oracle in plan `foamy-imagining-narwhal`.
	const frameCx = getSmartParseNumber(opts.w ?? 0, 'X', target._presLayout)
	const frameCy = getSmartParseNumber(opts.h ?? 0, 'Y', target._presLayout)
	const ar = target._presLayout.width / target._presLayout.height
	const n = targets.length
	const cols = Math.ceil(Math.sqrt(n))
	const rows = Math.ceil(n / cols)
	const gap = Math.round(frameCx * 0.0124)
	let tileW = (frameCx - (cols - 1) * gap) / cols
	let tileH = tileW / ar
	const tileHFit = (frameCy - (rows - 1) * gap) / rows
	if (tileH > tileHFit) {
		tileH = tileHFit
		tileW = tileH * ar
	}
	tileW = Math.round(tileW)
	tileH = Math.round(tileH)
	const gridW = cols * tileW + (cols - 1) * gap
	const gridH = rows * tileH + (rows - 1) * gap
	const originX = Math.round((frameCx - gridW) / 2)
	const originY = Math.round((frameCy - gridH) / 2)

	// One shared preview image across all tiles (identical placeholder/cover).
	const previewRid = registerPreviewImage(target, opts.coverImage)
	const tiles: ZoomTileInternal[] = []
	targets.forEach((section, i) => {
		const firstSlide = section._slides[0]
		if (!firstSlide) return // guarded above (length > 0), but keeps the emit total-function
		const r = Math.floor(i / cols)
		const c = i % cols
		tiles.push({
			sectionId: section._id,
			previewRid,
			fallbackSlideRid: registerSlideRel(target, firstSlide._slideNum),
			zmPrId: zoomGuid(),
			grid: { x: originX + c * (tileW + gap), y: originY + r * (tileH + gap), cx: tileW, cy: tileH },
		})
	})

	pushZoomObject(target, 'summary', opts, {
		tiles,
		returnToParent: !!opts.returnToParent,
		transitionDur: opts.transitionDur ?? 1000,
	})
}
