/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * The slide model: the PRIVATE generator wire shapes (`SlideRel`, `SlideRelMedia`, `SlideObject`),
 * groups, the merged `ObjectOptions` bag, and the `SlideLayout`/`PresSlide` authoring surfaces.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { CHART_NAME, PLACEHOLDER_TYPE, SHAPE_NAME, SlideObjectType, TableStyle } from '../core-enums.js'
import type { AnimationProps, TransitionProps } from './animation.js'
import type { ChartMulti, ChartOpts, OptsChartData, SlideRelChart } from './chart.js'
import type { BackgroundProps, Coord, HexColor, Margin, PositionProps } from './core.js'
import type { MasterTextStyleProps, SlideNumberProps } from './master.js'
import type { ImageBaseProps, ImageProps, MediaProps, MediaType } from './media.js'
import type { ObjectNameProps } from './object.js'
import type { PresLayout } from './pres.js'
import type { ShapeProps } from './shape.js'
import type { ConnectorProps, HyperlinkProps } from './style.js'
import type { TableCell, TableCellProps, TableProps, TableRow } from './table.js'
import type { CommentProps, NotesProps, SlideComment, TextProps, TextPropsOptions } from './text.js'

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
}
// PRIVATE ^^^
/**
 * A child object that can be placed inside a group via `slide.addGroup()`.
 *
 * Uses the same key-tagged descriptor shape as `SlideMasterObject`, but limited to the
 * object types `addGroup` supports. A `group` child nests another group (an identity child
 * coordinate space is kept at every depth, so descendants keep their slide-absolute
 * coordinates). Charts, media, tables, and placeholders are intentionally excluded (see
 * `addGroup`); passing one logs a warning and skips it.
 */
export type GroupChildProps =
	| { image: ImageProps }
	| { line: ShapeProps }
	| { rect: ShapeProps }
	| { roundRect: ShapeProps }
	| { shape: { type: SHAPE_NAME; options?: ShapeProps } }
	| { text: { text: string | number | TextProps[]; options?: TextPropsOptions } }
	| { group: { children: GroupChildProps[]; options?: GroupProps } }
/**
 * Options for `slide.addGroup()`.
 *
 * The group keeps an identity child coordinate space (`chOff/chExt == off/ext`) at every depth:
 * children — including nested groups — keep their slide-absolute `x/y/w/h`. A consequence worth
 * knowing: a group's own frame never moves or scales its children, it only places the selection
 * handle and the rotate pivot.
 *
 * The frame is all-or-nothing. Pass all four of `x/y/w/h` to set it explicitly, or none to have it
 * auto-computed as the bounding box of the children. A partial frame is ambiguous, so it warns and
 * falls back to auto-bounds on every axis.
 */
export interface GroupProps extends PositionProps, ObjectNameProps {
	/** Rotation in degrees (applied to the whole group) */
	rotate?: number
	/** Flip the group horizontally */
	flipH?: boolean
	/** Flip the group vertically */
	flipV?: boolean
}
export interface ObjectOptions extends ImageBaseProps, PositionProps, ShapeProps, TableCellProps, TextPropsOptions {
	_placeholderIdx?: number
	_placeholderType?: PLACEHOLDER_TYPE
	/** Connector adjust-guide values (OOXML 1000ths-of-a-percent), one per bend; emitted as `<a:gd name="adjN">` */
	_connectorAdj?: number[]
	/** Connector start-point binding: target shape `objectName` + connection-site index; resolved to `<a:stCxn>` at serialize time */
	_startCxn?: { name: string; idx: number }
	/** Connector end-point binding: target shape `objectName` + connection-site index; resolved to `<a:endCxn>` at serialize time */
	_endCxn?: { name: string; idx: number }
	/**
	 * Image: which dimensions were omitted by the user and should be derived from the image's
	 * natural pixel size at serialize time. Path-based images can't be measured synchronously in
	 * `addImage()` (bytes are loaded async during export), so the missing extent is backfilled
	 * once `_relsMedia[].data` is populated. `{ w, h }` true means "derive this side from the
	 * natural ratio". Base64 `data` images are measured eagerly in `addImage()` and never set this.
	 */
	_szAuto?: { w: boolean; h: boolean }

	cx?: Coord
	cy?: Coord
	margin?: Margin
	colW?: number | number[] // table
	rowH?: number | number[] // table
	hasHeader?: boolean // table
	hasFooter?: boolean // table
	hasBandedRows?: boolean // table
	hasBandedColumns?: boolean // table
	hasFirstColumn?: boolean // table
	hasLastColumn?: boolean // table
	rtl?: boolean // table
	tableStyle?: TableStyle | string // table
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
export interface SlideLayout {
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
export interface PresSlide {
	addChart(data: OptsChartData[], options: ChartOpts & { type: CHART_NAME }): PresSlide
	addChart(charts: ChartMulti[], options?: ChartOpts): PresSlide
	addConnector: (options: ConnectorProps) => PresSlide
	addImage: (options: ImageProps) => PresSlide
	addMedia: (options: MediaProps) => PresSlide
	addComment: (options: CommentProps) => PresSlide
	addNotes: (notes: string | NotesProps | NotesProps[]) => PresSlide
	addShape: (shapeName: SHAPE_NAME, options?: ShapeProps) => PresSlide
	addTable: (tableRows: TableRow[], options?: TableProps) => PresSlide
	addText: (text: string | number | TextProps[], options?: TextPropsOptions) => PresSlide
	addAnimation: (options: AnimationProps) => PresSlide

	readonly newAutoPagedSlides?: PresSlide[]

	/**
	 * Slide-show transition played when advancing to this slide (`p:transition`).
	 * @example slide.transition = { type: 'fade', durationMs: 1500 }
	 */
	transition?: TransitionProps

	/**
	 * Slide width in inches, resolved from the active presentation layout.
	 * Use for coordinate math instead of hard-coding layout dimensions.
	 * @example slide.addText('Centered', { x: 0, w: slide.width, align: 'center' })
	 */
	readonly width?: number
	/**
	 * Slide height in inches, resolved from the active presentation layout.
	 */
	readonly height?: number

	/**
	 * Background color or image (`color` | `path` | `data`)
	 * @example { color: 'FF3399' } - hex color
	 * @example { color: 'FF3399', transparency:50 } - hex color with 50% transparency
	 * @example { path: 'https://onedrives.com/myimg.png` } - retrieve image via URL
	 * @example { path: '/home/user/images/myimg.png` } - retrieve image via local path
	 * @example { data: 'image/png;base64,iVtDaDrF[...]=' } - base64 string
	 */
	background?: BackgroundProps
	/**
	 * Default text color (hex format)
	 * @example 'FF3399'
	 * @default '000000' (DEF_FONT_COLOR)
	 */
	color?: HexColor
	/**
	 * Whether slide is hidden
	 * @default false
	 */
	hidden?: boolean
	/**
	 * Slide number options
	 */
	slideNumber?: SlideNumberProps
}
export interface PresSlideInternal extends SlideBaseProps, PresSlide {
	_rId: number
	_slideLayout: SlideLayoutInternal | null
	_slideId: number
	/** Preset build animations on this slide, in play order (see {@link PresSlide.addAnimation}). */
	_animations: AnimationProps[]
}
export interface AddSlideProps {
	/** Title of the slide master to use for the new slide (the `title` passed to {@link SlideMasterProps} via `defineSlideMaster`). */
	masterTitle?: string
	sectionTitle?: string
}
