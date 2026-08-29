/**
 * The slide model: the PRIVATE generator wire shapes (`SlideRel`, `SlideRelMedia`, `SlideObject`),
 * groups, the merged `ObjectOptions` bag, and the `SlideLayout`/`Slide` authoring surfaces.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { CHART_NAME, PLACEHOLDER_TYPE, SHAPE_NAME, SlideObjectType, TableStyle } from '../enums.js'
import type { AnimationProps, TransitionProps } from './animation.js'
import type { ChartMulti, ChartOpts, OptsChartData } from './chart.js'
import type { BackgroundProps, Coord, HexColor, Margin, PositionProps } from './core.js'
import type { SlideNumberProps } from './master.js'
import type { ImageBaseProps, ImageProps, MediaProps, OleObjectProps } from './media.js'
import type { Model3dProps } from './model3d.js'
import type { CommonObjectDescriptor, ObjectNameProps } from './object.js'
import type { ShapeProps } from './shape.js'
import type { ConnectorProps } from './style.js'
import type { TableCellProps, TableProps, TableRow } from './table.js'
import type { CommentProps, NotesProps, TextProps, TextPropsOptions } from './text.js'

/**
 * A child object that can be placed inside a group via `slide.addGroup()`.
 *
 * The six descriptors it shares with `SlideMasterObject` are {@link CommonObjectDescriptor};
 * `group` is the group's own, and the master's `chart` and `placeholder` are deliberately not
 * here. A `group` child nests another group (an identity child
 * coordinate space is kept at every depth, so descendants keep their slide-absolute
 * coordinates). Charts, media, tables, and placeholders are intentionally excluded (see
 * `addGroup`); passing one logs a warning and skips it.
 */
export type GroupChildProps = CommonObjectDescriptor | { group: { children: GroupChildProps[]; options?: GroupProps } }
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
	tableStyle?: TableStyle // table
	/**
	 * Table perimeter border. `addTableDefinition` normalizes it to a 4-tuple (TRBL, sides
	 * left unset as `undefined`) before it reaches the emitter, so the serializer never has
	 * to re-handle the single-`BorderProps` form.
	 */
	outerBorder?: TableProps['outerBorder'] // table
	/** Table background (`a:tblPr` fill), as distinct from `fill`, which is stamped onto each cell. */
	tableFill?: TableProps['tableFill'] // table
}
export interface SlideLayout {
	background?: BackgroundProps
}

/**
 * One authored object on a slide, as {@link Slide.objects} reports it: what it is and how it is
 * addressed, not how it is drawn. It is a snapshot taken when the getter ran, not a live handle —
 * nothing on it writes back, and a later `addShape`/`groupObjects` leaves it describing the slide
 * as it was.
 *
 * The point of it is composition. A slide assembled by independent renderers arrives with objects
 * whose descriptors nobody kept, and the only durable handle on them is the `objectName` each one
 * was authored with; without a way to enumerate those, a consumer that wants to act on what is
 * already there has to either make every renderer surrender its internals or keep a parallel
 * ledger of what it added — and a ledger is wrong the moment a renderer adds an object it did not
 * announce.
 */
export interface SlideObjectInfo {
	/** What kind of object this is — the same value the object was authored as. */
	readonly type: SlideObjectType
	/**
	 * Selection Pane name, in the spelling the caller passed to `objectName`, so it can be handed
	 * straight back to `groupObjects()` or to an animation or connector reference.
	 *
	 * Always a string: an object authored without a name still gets the generated `Shape 3` /
	 * `Text 1` / `Group 2` identity PowerPoint shows in the Selection Pane, and that name addresses
	 * it just as well. The two are indistinguishable here on purpose — nothing downstream of
	 * authoring records which is which, and a consumer that cares can tell them apart by the
	 * naming convention it chose for its own objects.
	 */
	readonly objectName: string
	/** True when the object occupies a layout placeholder, which grouping refuses on top of kind. */
	readonly isPlaceholder: boolean
	/**
	 * Whether `groupObjects()` would accept this object on kind alone. It cannot speak for the
	 * *selection* — an unresolved, duplicated or ambiguous name still throws — so it answers "is
	 * this object groupable", never "will this call succeed".
	 */
	readonly canGroup: boolean
	/**
	 * Children, when `type` is `group`; empty otherwise. Ordered bottom-to-top like
	 * {@link Slide.objects} itself, and nested to whatever depth the groups nest.
	 */
	readonly children: readonly SlideObjectInfo[]
}
export interface Slide {
	addChart(data: OptsChartData[], options: ChartOpts & { type: CHART_NAME }): Slide
	addChart(charts: ChartMulti[], options?: ChartOpts): Slide
	addConnector: (options: ConnectorProps) => Slide
	addImage: (options: ImageProps) => Slide
	addMedia: (options: MediaProps) => Slide
	/** Embed a 3D model (Insert ▸ 3D Models) — a `.glb` PowerPoint 2019+ renders live. */
	addModel3d: (options: Model3dProps) => Slide
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

	/**
	 * The objects authored on this slide so far, bottom-to-top in z-order — the read-back half of
	 * `groupObjects()`, which addresses those same objects by name but until now gave no way to
	 * learn the names.
	 *
	 * A fresh snapshot on every access, and inert: it describes the slide, and writing to it does
	 * nothing. To act on what it reports, call the authoring API with the names it hands you.
	 * @example
	 * const cards = slide.objects.filter((o) => o.canGroup && o.objectName?.startsWith('card:'))
	 * slide.groupObjects(cards.map((o) => o.objectName), { objectName: 'Cards' })
	 */
	readonly objects: readonly SlideObjectInfo[]

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
