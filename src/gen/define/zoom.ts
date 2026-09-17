/**
 * ts-pptx: Zoom Definition (Slide / Section / Summary Zoom — Insert ▸ Zoom).
 *
 * Resolves a zoom's target(s) to the ids the emitter needs (`sldId` / section GUID), registers
 * the preview-image media rel and the `.../slide` fallback rel(s), lays out the Summary Zoom grid,
 * and pushes a `SlideObject{ _type: zoom, zoom }` for `gen/slide/objects/zoom.ts` to emit as a
 * `<p:graphicFrame>`. See {@link ../../types/zoom} for the preview-image behavior.
 */
import { SlideObjectType } from '../../enums.js'
import { warn } from '../../diagnostics.js'
import type { SectionZoomProps, SlideZoomProps, SummaryZoomProps } from '../../types/zoom.js'
import type {
	ObjectOptionsInternal,
	PresSlideInternal,
	SectionInternalProps,
	SlideObject,
	ZoomInternal,
	ZoomTileInternal,
} from '../../types/internal.js'
import { bracedGuid, getNewRelId } from '../utils.js'
import { clampRangedInput, getSmartParseNumber } from '../../units-internal.js'
import { framedObjectOptions } from './object-options.js'
import { registerPreviewImage } from './preview-image.js'
import { MIN_SLIDE_ID } from '../../ooxml/ids.js'

const ZOOM_LABEL = { slide: 'Slide Zoom', section: 'Section Zoom', summary: 'Summary Zoom' } as const
const ZOOM_API = { slide: 'addSlideZoom', section: 'addSectionZoom', summary: 'addSummaryZoom' } as const

/** The longest `zmPr@transitionDur` written, in milliseconds: the largest signed 32-bit integer. */
const MAX_TRANSITION_DUR_MS = 2147483647

/** Register a `.../slide` rel (used by the fallback picture's `hlinkClick`) to a 1-based slide number. */
function registerSlideRel(target: PresSlideInternal, slideNum: number): number {
	const rId = getNewRelId(target)
	target._rels.push({ type: SlideObjectType.hyperlink, data: 'slide', rId, Target: String(slideNum) })
	return rId
}

/** The two `zmPr` settings every variant takes from its options the same way. */
type ZoomSettings = Pick<ZoomInternal, 'returnToParent' | 'transitionDur'>

/**
 * Resolve the `zmPr` settings a zoom's options state, before the definer registers anything.
 *
 * All three definers copied these, and `transitionDur` went into the attribute unchecked: `NaN`
 * and `Infinity` were written as that text, and `-1` and `1.5` as given. It is clamped here the
 * one way an out-of-range number is: `NaN` throws naming the option, a number outside
 * 0-2147483647 moves to the bound with a warning, and the result is a whole number of milliseconds.
 * @param opts - the zoom's options
 * @param variant - which zoom, for the option name in a message
 */
function resolveZoomSettings(
	opts: SlideZoomProps | SectionZoomProps | SummaryZoomProps,
	variant: ZoomInternal['variant']
): ZoomSettings {
	const transitionDur =
		opts.transitionDur == null
			? 1000
			: Math.round(
					clampRangedInput(
						opts.transitionDur,
						0,
						MAX_TRANSITION_DUR_MS,
						'zoom/transition-duration-out-of-range',
						`${ZOOM_API[variant]} transitionDur`,
						'zoom/invalid-transition-duration'
					)
				)
	return { returnToParent: !!opts.returnToParent, transitionDur }
}

/**
 * A zoom's resolved options: its Selection Pane name and its frame.
 *
 * Split from {@link pushZoomObject} so the Summary grid can lay itself out inside the *resolved*
 * frame. It used to re-read `opts.w ?? 0` and `opts.h ?? 0` for the grid while the emitted object
 * took its frame from here, so the same input was parsed twice by two expressions that nothing
 * held together.
 *
 * A zoom has no default size, unlike every other framed object: `{ w: 0, h: 0 }` is what the
 * caller gets when they state no frame, and `resolveAuthoredFrame` says so rather than quoting the
 * zero back at them.
 */
function zoomObjectOptions(
	target: PresSlideInternal,
	variant: ZoomInternal['variant'],
	opts: SlideZoomProps | SectionZoomProps | SummaryZoomProps
): ObjectOptionsInternal {
	// A zoom's options declare no `altText`, and its emitter writes none.
	return framedObjectOptions(target, SlideObjectType.zoom, opts, {
		label: ZOOM_LABEL[variant],
		kind: 'zoom',
		api: ZOOM_API[variant],
		defaults: { x: 0, y: 0, w: 0, h: 0 },
	})
}

/** Shared object-scaffold for all three variants. */
function pushZoomObject(
	target: PresSlideInternal,
	variant: ZoomInternal['variant'],
	options: ObjectOptionsInternal,
	zoom: Omit<ZoomInternal, 'variant'>
): void {
	const newObject: SlideObject = {
		_type: SlideObjectType.zoom,
		options,
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
	// A number has to be a slide number: `-2` and `1.5` were written as `slide-2.xml` and
	// `sldId="256.5"`. Whether that slide exists is checked when the deck is written, once it is complete.
	if (typeof opts.target === 'number' && !(Number.isInteger(opts.target) && opts.target >= 1)) {
		warn('zoom/unresolved-target', `addSlideZoom: \`target\` ${opts.target} is not a 1-based slide number; ignoring.`)
		return
	}
	const targetSlide = opts.target as PresSlideInternal
	const sldId = typeof opts.target === 'number' ? MIN_SLIDE_ID + (opts.target - 1) : targetSlide._slideId
	const slideNum = typeof opts.target === 'number' ? opts.target : targetSlide._slideNum
	if (sldId == null || slideNum == null) {
		warn('zoom/unresolved-target', 'addSlideZoom: could not resolve the target slide; ignoring.')
		return
	}
	const settings = resolveZoomSettings(opts, 'slide')

	const previewRid = registerPreviewImage(target, opts.coverImage, 'a zoom `coverImage`')
	const fallbackSlideRid = registerSlideRel(target, slideNum)
	const tile: ZoomTileInternal = { sldId, previewRid, fallbackSlideRid, zmPrId: bracedGuid() }
	pushZoomObject(target, 'slide', zoomObjectOptions(target, 'slide', opts), {
		tiles: [tile],
		...settings,
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
	const settings = resolveZoomSettings(opts, 'section')

	const previewRid = registerPreviewImage(target, opts.coverImage, 'a zoom `coverImage`')
	const fallbackSlideRid = registerSlideRel(target, firstSlide._slideNum)
	const tile: ZoomTileInternal = { sectionId: section._id, previewRid, fallbackSlideRid, zmPrId: bracedGuid() }
	pushZoomObject(target, 'section', zoomObjectOptions(target, 'section', opts), {
		tiles: [tile],
		...settings,
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
	const settings = resolveZoomSettings(opts, 'summary')

	const options = zoomObjectOptions(target, 'summary', opts)

	// Grid geometry (EMU). Tiles preserve the slide aspect ratio; the grid is centered in the frame,
	// last row left-aligned (matching PowerPoint). See the Summary oracle in plan `foamy-imagining-narwhal`.
	// The frame is the one `options` carries, not a second reading of `opts`.
	const frameCx = getSmartParseNumber(options.w ?? 0, 'X', target._presLayout)
	const frameCy = getSmartParseNumber(options.h ?? 0, 'Y', target._presLayout)
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
	const previewRid = registerPreviewImage(target, opts.coverImage, 'a zoom `coverImage`')
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
			zmPrId: bracedGuid(),
			grid: { x: originX + c * (tileW + gap), y: originY + r * (tileH + gap), cx: tileW, cy: tileH },
		})
	})

	pushZoomObject(target, 'summary', options, {
		tiles,
		...settings,
	})
}
