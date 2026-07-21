/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * Generator-internal wire shapes. NOT part of the public authoring contract: these are the
 * normalized structures the emitters pass around, and their `_`-prefixed members are
 * implementation detail that may change without notice.
 *
 * Kept in one module so the public surface can be drawn around them.
 */
import type { CHART_NAME, ChartType, SHAPE_NAME, SlideObjectType } from '../core-enums.js'
import type { EmbeddedFont } from '../embedded-fonts.js'
import type { AnimationProps } from './animation.js'
import type { ChartMulti, ChartOpts, OptsChartData } from './chart.js'
import type { BackgroundProps, Margin } from './core.js'
import type { MasterTextStyleProps, SlideNumberProps } from './master.js'
import type { MediaProps, MediaType } from './media.js'
import type { PresLayout, PresentationProps, SectionProps } from './pres.js'
import type { ObjectOptions, PresSlide, SlideLayout } from './slide.js'
import type { HyperlinkProps, ShadowProps } from './style.js'
import type { TableCell, TableStyleProps } from './table.js'
import type { SlideComment, TextProps } from './text.js'

/**
 * Internal, wire-normalized shadow shape produced by `correctShadowOptions` — not part of the
 * public `ShadowProps` input. `opacity` (0.0 fully transparent – 1.0 fully opaque) is the alpha
 * derived from the public `transparency` option (or a color's embedded alpha byte); it is what
 * every emit site reads, so downstream code stays unit-agnostic about the public percent scale.
 */
export interface ShadowPropsInternal extends ShadowProps {
	opacity?: number
}
// Used internally, probably shouldn't be used by end users
export interface OptsChartDataInternal extends OptsChartData {
	labels?: string[][]
	/** Series index; always assigned by addChartDefinition() before this internal shape is built. */
	_dataIndex: number
}
export interface ChartOptsInternal extends ChartOpts {
	_type?: ChartType | ChartMulti[] // internal, normalized from `CHART_NAME`
}
export interface SlideRelChart extends OptsChartData {
	type: CHART_NAME | ChartMulti[]
	opts: ChartOptsInternal
	data: OptsChartDataInternal[]
	// internal below
	rId: number
	Target: string
	globalId: number
	fileName: string
	/**
	 * True when `type` is a chartEx (cx:) chart (e.g. waterfall): the part is emitted as
	 * `chartEx{globalId}.xml` in the chart-extension namespace, gets the `chartex+xml` content
	 * type and the MS `.../2014/relationships/chartEx` rel type, and is referenced from the slide
	 * via `<mc:AlternateContent>`. Set at define time from {@link isChartExType}.
	 */
	isChartEx?: boolean
}
/**
 * Internal record pairing a registered custom table style with its generated GUID.
 */
export interface TableStyleInternal {
	/** Braced, upper-case GUID emitted as both `styleId` and `<a:tableStyleId>`. */
	guid: string
	def: TableStyleProps
}
// PRIVATE vvv
export interface SlideRel {
	type: SlideObjectType
	/**
	 * Relationship target, stored **unescaped**. Every emitter escapes it on the way out
	 * (`gen/slide/object.ts`, `gen/slide/notes.ts`, and `read/opc/relationships.ts` for the
	 * append path). Escaping at definition time instead would double-escape on append,
	 * where the serializer escapes again — so `&` in a hyperlink or online-video URL must
	 * arrive here verbatim.
	 */
	Target: string
	fileName?: string
	data: any[] | string
	opts?: ChartOpts
	path?: string
	extn?: string
	globalId?: number
	rId: number
}
export interface SlideRelMedia {
	type: string
	opts?: MediaProps
	path?: string
	extn?: string
	data?: string | ArrayBuffer
	/** used to indicate that a media file has already been read/enocded (PERF) */
	isDuplicate?: boolean
	isSvgPng?: boolean
	svgSize?: { w: number; h: number }
	rId: number
	/** Unescaped — see {@link SlideRel.Target}. Doubles as the zip entry name for embedded media. */
	Target: string
}
export interface SlideObject {
	_type: SlideObjectType
	options?: ObjectOptions
	// text
	text?: TextProps[]
	// table
	arrTabRows?: TableCell[][]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: HyperlinkProps
	// media
	media?: string
	mtype?: MediaType
	mediaRid?: number
	loop?: boolean
	loopCount?: number
	shape?: SHAPE_NAME
	// group (flat group): child render-objects emitted inside this object's `<p:grpSp>`
	_groupObjects?: SlideObject[]
	// zoom (slide/section/summary): resolved tiles + attrs for the `<p:graphicFrame>` emitter
	zoom?: ZoomInternal
}
/** One tile inside a zoom object (a Slide/Section Zoom has one; a Summary Zoom has N). */
export interface ZoomTileInternal {
	/** Slide Zoom: target slide id (= `PresSlideInternal._slideId`, emitted as `sldZmObj@sldId`). */
	sldId?: number
	/** Section/Summary Zoom: target section GUID (`{...}`, emitted as `sectionZmObj@sectionId`). */
	sectionId?: string
	/** rId of the preview/cover image blip (`zmPr>blipFill>a:blip@r:embed`). Shared across tiles when identical. */
	previewRid: number
	/** rId of the `.../slide` rel used by the `mc:Fallback` picture's `hlinkClick` (target/section-start slide). */
	fallbackSlideRid: number
	/** Per-tile `zmPr@id` GUID (`{...}`). */
	zmPrId: string
	/** Summary Zoom only: this tile's grid cell within the graphic-frame coordinate space (EMU). */
	grid?: { x: number; y: number; cx: number; cy: number }
}
/** Resolved zoom payload carried on a {@link SlideObject} until `gen/slide/object.ts` emits it. */
export interface ZoomInternal {
	variant: 'slide' | 'section' | 'summary'
	tiles: ZoomTileInternal[]
	/** `zmPr@returnToParent`; emitted as `"0"`/`"1"` for Slide Zoom, omitted for Section/Summary. */
	returnToParent: boolean
	/** `zmPr@transitionDur` (ms). */
	transitionDur: number
}
export interface SlideBaseProps {
	_bkgdImgRid?: number
	_margin?: Margin
	_name?: string
	_presLayout: PresLayout
	_rels: SlideRel[]
	_relsChart: SlideRelChart[] // needed as we use args:"PresSlide|SlideLayout" often
	_relsMedia: SlideRelMedia[] // needed as we use args:"PresSlide|SlideLayout" often
	_relsNotes?: SlideRel[] // hyperlink rels emitted in the notes-slide part (notesSlideN.xml.rels)
	_comments?: SlideComment[] // review comments emitted in the per-slide comments part (commentN.xml)
	_txStyles?: MasterTextStyleProps // per-level master text styles emitted in slideMaster1.xml <p:txStyles> (deck-wide; set via defineSlideMaster textStyles)
	_slideNum: number
	_slideNumberProps?: SlideNumberProps | null
	_slideObjects: SlideObject[]
	/**
	 * Per-kind counters backing default Selection Pane names (`Shape 0`, `Image 1`, `Group 1`, …),
	 * keyed by `SlideObjectType`. Monotonic for the life of the slide: an object consumes its index
	 * when it is added, whether it stays top-level or is moved into a group's `_groupObjects`.
	 * Lazily created by `nextObjectNameIdx` (`gen/define/object-name.ts`).
	 */
	_objectNameCounts?: Partial<Record<SlideObjectType, number>>

	background?: BackgroundProps
}
export interface SlideLayoutInternal extends SlideBaseProps, SlideLayout {
	_slide?: {
		_bkgdImgRid?: number
		back: string
		color: string
		hidden?: boolean
	} | null
}
export interface PresSlideInternal extends SlideBaseProps, PresSlide {
	_rId: number
	_slideLayout: SlideLayoutInternal | null
	_slideId: number
	/** Preset build animations on this slide, in play order (see {@link PresSlide.addAnimation}). */
	_animations: AnimationProps[]
}
export interface SectionInternalProps extends SectionProps {
	_type?: 'user' | 'default'
	_slides: PresSlideInternal[]
	/** Stable section GUID (`{XXXXXXXX-...}`), assigned at section creation so Section/Summary Zoom can address it. */
	_id: string
}
// PRIVATE interface
export interface PresentationPropsInternal extends PresentationProps {
	masterSlide: PresSlideInternal
	sections: SectionInternalProps[]
	slideLayouts: SlideLayoutInternal[]
	slides: PresSlideInternal[]
	/** Author-side embedded fonts (see {@link PptxGenJS.embedFont}); empty when none. */
	embeddedFonts: EmbeddedFont[]
}
