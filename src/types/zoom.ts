/**
 * ts-pptx: Zoom link types (Slide / Section / Summary Zoom — PowerPoint's Insert ▸ Zoom).
 *
 * A zoom is a clickable tile on a slide that navigates to a target slide (Slide Zoom), the
 * start of a section (Section Zoom), or — laid out as a grid — the start of every section
 * (Summary Zoom). Emitted as a `<p:graphicFrame>` in the 2016 zoom namespaces, wrapped in
 * `<mc:AlternateContent>` with a hyperlinked-picture fallback for pre-2016 consumers.
 *
 * PREVIEW IMAGE: the tile shows a thumbnail of its target. ts-pptx is Node-first and cannot
 * rasterize a slide, so by default it emits a neutral gray **placeholder** — PowerPoint does NOT
 * refresh it on open, but it regenerates to the live thumbnail once the target slide is next
 * edited. Supply `coverImage` for a picture that ships as-authored instead.
 */
import type { PositionProps } from './core.js'
import type { ObjectLockProps } from './object.js'
import type { Slide } from './slide.js'

/** Options shared by all three zoom kinds. */
export interface ZoomBaseProps extends PositionProps {
	/** Object name (PowerPoint selection-pane label). Auto-named (e.g. `Slide Zoom 3`) when omitted. */
	objectName?: string
	/**
	 * Ship a caller-supplied thumbnail instead of the gray placeholder — a raster image `path`
	 * (Node/local) or base64 `data:` URI. Shown as-authored (not auto-refreshed by PowerPoint).
	 */
	coverImage?: { path?: string; data?: string }
	/**
	 * After the zoom animates to its target, return to this (parent) slide when the target's
	 * content finishes. Maps to `zmPr@returnToParent`. Default `false`.
	 */
	returnToParent?: boolean
	/** Zoom transition duration in milliseconds (`zmPr@transitionDur`). Default `1000`. */
	transitionDur?: number
	/**
	 * Object lock flags ({@link ObjectLockProps}), as `objectLock` on every other object type.
	 *
	 * A zoom is a `p:graphicFrame`, so these land on its `a:graphicFrameLocks` and the flags that
	 * apply are the graphic-frame ones (`noGrp`, `noDrilldown`, `noSelect`, `noChangeAspect`,
	 * `noMove`, `noResize`); a picture-only flag warns and is dropped, as everywhere else. The
	 * `mc:Fallback` picture keeps its own fixed lock set — it is what pre-2016 consumers draw, and
	 * folding a graphic-frame flag set onto a `a:picLocks` element would warn about every flag the
	 * two element types do not share.
	 *
	 * `noChangeAspect` is on by default (a zoom tile mirrors its target slide's aspect); pass it
	 * `false` to lift it.
	 * @example { noMove: true, noResize: true } // pin the tile in place
	 */
	objectLock?: ObjectLockProps
}

/** Slide Zoom: a tile that zooms to a single target slide. */
export interface SlideZoomProps extends ZoomBaseProps {
	/** Target slide — a `Slide` returned by `addSlide()`, or its 1-based slide number. */
	target: Slide | number
}

/** Section Zoom: a tile that zooms to the start of a named section. */
export interface SectionZoomProps extends ZoomBaseProps {
	/** Title of the target section (as passed to `addSection({ title })`). */
	sectionTitle: string
}

/**
 * Summary Zoom: a grid of tiles, one per section, that navigates to each section's start.
 * The section containing the host slide is excluded (a summary does not link to itself).
 */
export type SummaryZoomProps = ZoomBaseProps
