/**
 * The slide model: the PRIVATE generator wire shapes (`SlideRel`, `SlideRelMedia`, `SlideObject`),
 * groups, the merged `ObjectOptions` bag, and the `SlideLayout`/`Slide` authoring surfaces.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { CHART_NAME, PLACEHOLDER_TYPE, SHAPE_NAME, TableStyle } from '../enums.js'
import type { AnimationProps, TransitionProps } from './animation.js'
import type { ChartMulti, ChartOpts, OptsChartData } from './chart.js'
import type { BackgroundProps, Coord, HexColor, Margin, PositionProps } from './core.js'
import type { SlideNumberProps } from './master.js'
import type { ImageBaseProps, ImageProps, MediaProps, OleObjectProps } from './media.js'
import type { ObjectNameProps } from './object.js'
import type { ShapeProps } from './shape.js'
import type { ConnectorProps } from './style.js'
import type { TableCellProps, TableProps, TableRow } from './table.js'
import type { CommentProps, NotesProps, TextProps, TextPropsOptions } from './text.js'

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
	/** Connector start-point binding: target shape `objectName` (raw, as the caller spelled it) + connection-site index; resolved to `<a:stCxn>` at serialize time */
	_startCxn?: { name: string; idx: number }
	/** Connector end-point binding: target shape `objectName` (raw, as the caller spelled it) + connection-site index; resolved to `<a:endCxn>` at serialize time */
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
	/**
	 * Table perimeter border. `addTableDefinition` normalizes it to a 4-tuple (TRBL, sides
	 * left unset as `undefined`) before it reaches the emitter, so the serializer never has
	 * to re-handle the single-`BorderProps` form.
	 */
	outerBorder?: TableProps['outerBorder'] // table
	/**
	 * Table: the resolved answer to "may a cell border be left to the table style?" — the
	 * caller's `styleDrivenCells` *and* a `tableStyle` that `defineTableStyle()` registered.
	 * Resolved in `addTableDefinition` because only the presentation can tell a registered
	 * style from a built-in one; the emitter needs it to know that a cell with no borders of
	 * its own means "the style draws them" rather than "this cell was never through the
	 * definition step".
	 */
	_styleDrivenCells?: boolean // table
	/** Table background (`a:tblPr` fill), as distinct from `fill`, which is stamped onto each cell. */
	tableFill?: TableProps['tableFill'] // table
}
export interface SlideLayout {
	background?: BackgroundProps
}
export interface Slide {
	addChart(data: OptsChartData[], options: ChartOpts & { type: CHART_NAME }): Slide
	addChart(charts: ChartMulti[], options?: ChartOpts): Slide
	addConnector: (options: ConnectorProps) => Slide
	addImage: (options: ImageProps) => Slide
	addMedia: (options: MediaProps) => Slide
	/** Embed an OLE object (Insert ▸ Object) whose bytes travel inside the `.pptx`. */
	addOleObject: (options: OleObjectProps) => Slide
	addComment: (options: CommentProps) => Slide
	addNotes: (notes: string | NotesProps | NotesProps[]) => Slide
	addShape: (shapeName: SHAPE_NAME, options?: ShapeProps) => Slide
	addTable: (tableRows: TableRow[], options?: TableProps) => Slide
	addText: (text: string | number | TextProps[], options?: TextPropsOptions) => Slide
	addAnimation: (options: AnimationProps) => Slide
	/** Group child object descriptors into a single PowerPoint group (`<p:grpSp>`). */
	addGroup: (children: GroupChildProps[], options?: GroupProps) => Slide
	/** Group objects already on this slide, addressed by their `objectName`, into a single group. */
	groupObjects: (objectNames: string[], options?: GroupProps) => Slide

	readonly newAutoPagedSlides?: Slide[]

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
export interface AddSlideProps {
	/** Title of the slide master to use for the new slide (the `title` passed to {@link SlideMasterProps} via `defineSlideMaster`). */
	masterTitle?: string
	sectionTitle?: string
}
