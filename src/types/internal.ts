/* oxlint-disable typescript/no-explicit-any */
/**
 * Generator-internal wire shapes. NOT part of the public authoring contract: these are the
 * normalized structures the emitters pass around, and their `_`-prefixed members are
 * implementation detail that may change without notice.
 *
 * Kept in one module so the public surface can be drawn around them.
 */
import type {
	AlignH,
	CHART_NAME,
	ChartType,
	PLACEHOLDER_TYPE,
	SHAPE_NAME,
	SlideObjectType,
	TextAnchor,
} from '../enums.js'
import type { EmbeddedFont } from '../embedded-fonts.js'
import type { AnimationProps, TransitionProps } from './animation.js'
import type { ChartMulti, ChartOpts, OptsChartData } from './chart.js'
import type { BackgroundOption, Margin, TextVertType } from './core.js'
import type { MasterTextStyleProps, SlideNumberProps } from './master.js'
import type { MediaProps, MediaType } from './media.js'
import type { TextShapeType } from '../ooxml/st-enums.js'
import type { PresLayout, PresentationProps, SectionProps } from './pres.js'
import type { ObjectOptions, Slide, SlideLayout } from './slide.js'
import type { HyperlinkProps, ShadowProps, ShapeFillProps } from './style.js'
import type { TableCell, TableCellProps, TableProps, TableRow, TableToSlidesProps } from './table.js'
import type { SlideComment, TextBulletProps, TextProps, TextPropsOptions } from './text.js'

/** The keys of `T` that are declared optional. */
// oxlint-disable-next-line typescript/no-empty-object-type -- `{} extends Pick<T, K>` is the test for "K is optional"; `object` fails it on a weak type.
export type OptionalKeysOf<T> = { [K in keyof T]-?: {} extends Pick<T, K> ? K : never }[keyof T]

/**
 * `T`, but every *optional* property may also be spelled as a present `undefined`.
 *
 * Under `exactOptionalPropertyTypes` a `foo?: T` says the key is either absent or holds a `T`.
 * That distinction is real for a bag that is **stored, spread or enumerated** — `{ ...defaults,
 * ...options }` inherits from an absent key and is overridden by a present `undefined` — and
 * those bags keep the strict declaration, with `delete` as their one spelling of absent (the
 * write-side twin of `compact()` in `script/from-read/values.ts`).
 *
 * It is noise for a bag that is only ever **read**: typically a function parameter its call sites
 * assemble inline out of values that may be unset (`{ color: opts.titleColor, … }`), where the
 * reader cannot tell the two states apart in the first place. Those parameters take this instead
 * of forcing every call site to build the object key by key.
 *
 * Required properties are left alone, so this never quietly makes a mandatory field optional.
 */
export type MaybeUndefined<T> = Omit<T, OptionalKeysOf<T>> & { [K in OptionalKeysOf<T>]?: T[K] | undefined }

/**
 * Internal, wire-normalized shadow shape produced by `normalizeShadowOptions` — not part of the
 * public `ShadowProps` input. `_alpha` (0.0 fully transparent – 1.0 fully opaque) is the alpha
 * derived from the public `transparency` option (or a color's embedded alpha byte); it is what
 * every emit site reads, so downstream code stays unit-agnostic about the public percent scale.
 * The `_` prefix keeps it clear of the removed public `opacity` input: a stray `opacity` from an
 * untyped caller lands on a field nothing reads, so it is inert rather than silently honored.
 */
/**
 * A transition carrying the relationship id its embedded sound part was registered under.
 *
 * Stamped by `registerTransitionSounds` at export time and read by the `p:sndAc` emitter. It
 * was declared on the public {@link TransitionProps}, where it was a documented option key
 * nobody intended: a caller setting it by hand would collide with the registration pass.
 */
export interface TransitionPropsInternal extends TransitionProps {
	_sndRId?: number
}

/**
 * A hyperlink carrying the relationship id `registerHyperlinkRel` minted for it.
 *
 * Every carrier of a hyperlink stays public (`hyperlink?: HyperlinkProps` on a shape, a run, a
 * table cell, a media object): a `HyperlinkProps` is assignable to this, because the one added
 * member is optional, so only the emitters that read or stamp the id name this type. The field
 * was on the public interface, where a caller setting it by hand would collide with the
 * registration pass rather than being ignored.
 */
export interface HyperlinkPropsInternal extends HyperlinkProps {
	_rId?: number
}

/**
 * A fill carrying the media relationship id its image was registered under, assigned at
 * add-time and read by the `a:blipFill` emitter.
 *
 * The carriers stay public for the reason {@link HyperlinkPropsInternal} gives: a
 * `ShapeFillProps` is assignable to this, so only the two sites that stamp and read the id
 * name it. `ShapeLineProps` used to `Omit` this key by hand, which is what a stroke's own
 * "no image fill" rule is really about -- it now omits `image` alone and says so.
 */
export interface ShapeFillPropsInternal extends ShapeFillProps {
	_imgRid?: number
}

/**
 * A drawn bullet carrying the relationship ids `addText()` registered its image under.
 *
 * `_rId` is the `a:blip r:embed`; for an SVG bullet it is the PNG-preview rel and `_rIdSvg`
 * is the `asvg:svgBlip` one, the same dual-rel form `addImage()` emits.
 */
export interface TextBulletPropsInternal extends TextBulletProps {
	_rId?: number
	_rIdSvg?: number
}

/**
 * Text options carrying the two members the emitters write onto them.
 *
 * `_bodyProp` is the normalized `a:bodyPr` the definer builds out of the caller's `align`,
 * `valign`, `margin`, `columns`, `textDirection` and `textWarp` — a *derived* shape, not an
 * input, which is what made its presence on the public interface a de-facto API nobody
 * intended. `_lineIdx` is the paragraph index the run emitter stamps while walking an array
 * of runs.
 */
export interface TextPropsOptionsInternal extends TextPropsOptions {
	_bodyProp?: {
		// Note: Many of these duplicated as user options are transformed to _bodyProp options for XML processing
		align?: AlignH
		/**
		 * The resolved `a:bodyPr/@anchor` token. The enum's *values* rather than the enum, so
		 * that both spellings of the same three tokens are accepted: the definer writes
		 * `TextAnchor.ctr` for its own default, while `resolveTextAnchor` answers with the
		 * schema token from the shared `valign` table, which cannot name the enum (it lives in
		 * `ooxml/`, which holds no runtime imports).
		 */
		anchor?: `${TextAnchor}`
		lIns?: number
		rIns?: number
		tIns?: number
		bIns?: number
		numCol?: number
		spcCol?: number
		vert?: TextVertType
		wrap?: boolean
		prstTxWarp?: TextShapeType
	}
	_lineIdx?: number
}

/**
 * A slide object's options carrying the six members the definers and emitters write onto them.
 *
 * All six are *derived* rather than authored: a placeholder's resolved index and type, the
 * adjust guides and connection-site bindings a connector resolves at serialize time, and the
 * flags recording which image dimension has to be backfilled from the natural pixel size once
 * the bytes have loaded. They were on the public {@link ObjectOptions}, where the two `Cxn`
 * ones in particular read as a supported way to bind a connector -- the supported way is
 * `startShape`/`startShapeIdx`.
 */
export interface ObjectOptionsInternal extends ObjectOptions, TextPropsOptionsInternal {
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
}

/**
 * A table cell carrying the seven members the auto-pager and the table emitter write onto it.
 *
 * All seven are produced by the write path rather than authored: the span dummies
 * (`_hmerge`/`_vmerge`) and the `_spanOrigin` they inherit their paint from, the per-line
 * split (`_lines`, `_lineHeight`) the auto-pager computes, `_rowContinue`, and the object-kind
 * tag. A caller who set one by hand would be feeding the pager its own output.
 *
 * `_lines` and `_spanOrigin` refer to *this* type rather than to the public `TableCell`,
 * because a split line and a span origin are themselves pager output.
 */
export interface TableCellInternal extends TableCell {
	options?: TableCellPropsInternal
	/** Narrowed alongside `_lines`: a cell's runs are pager output too. */
	text?: string | TableCellInternal[]
	_type?: SlideObjectType.tablecell
	/** lines in this cell (autoPage) */
	_lines?: TableCellInternal[][]
	/** height in EMU */
	_lineHeight?: number
	_hmerge?: boolean
	_vmerge?: boolean
	_rowContinue?: number
	/** origin cell of a colspan/rowspan span, set on the dummy `_hmerge`/`_vmerge` cells so they can
	 * inherit the origin's border/fill and render the merged region's outer edges */
	_spanOrigin?: TableCellInternal
}

/**
 * The header rows the auto-pager repeats on each continuation slide.
 *
 * Stashed by `addTable` (and by `tableToSlides`, which builds them out of the HTML `thead`)
 * and read by the pager. Never authored: it was on both `TableProps` and `TableToSlidesProps`,
 * where it read as a way to supply header rows separately from the table's own.
 */
export interface TablePropsInternal extends TableProps {
	_arrObjTabHeadRows?: TableRow[]
}

/** A row of {@link TableCellInternal}, the shape the pager builds and the emitter reads. */
export type TableRowInternal = TableCellInternal[]

/** {@link TablePropsInternal}, for the auto-pager's own option bag. */
export interface TableToSlidesPropsInternal extends TableToSlidesProps {
	_arrObjTabHeadRows?: TableRow[]
}

/**
 * A cell's own text options, carrying the normalized `a:bodyPr` the definer derives.
 *
 * A cell renders through the same text-body builder a shape does, so it reaches `_bodyProp`
 * by the same route -- see {@link TextPropsOptionsInternal}. Only that one member is added:
 * a cell's options are `TableCellProps`, which is deliberately narrower than the full text
 * bag (`columns` means something else on a table).
 */
export interface TableCellPropsInternal extends TableCellProps {
	_bodyProp?: TextPropsOptionsInternal['_bodyProp']
}

export interface ShadowPropsInternal extends ShadowProps {
	_alpha?: number
}
// Used internally, probably shouldn't be used by end users
export interface OptsChartDataInternal extends OptsChartData {
	labels?: string[][]
	/** Series index; always assigned by addChartDefinition() before this internal shape is built. */
	_dataIndex: number
}
export interface ChartOptsInternal extends ChartOpts {
	_type?: ChartType | ChartMultiInternal[] // internal, normalized from `CHART_NAME`
}
/**
 * A combo subchart's option overrides after `addChartDefinition` vetted them — the one bag in the
 * generator where a key may be *present* and hold `undefined`.
 *
 * Everywhere else "unset" has exactly one spelling, an absent key: `compact()` enforces it on the
 * read side (`script/from-read/values.ts`) and the `delete` idiom does on the write side
 * (`gen/define/chart.ts`). This bag needs the third state because it is merged *over* the
 * chart-level options at emit time (`{ ...rel.opts, ...type.options }` in `gen/chart/chart-xml.ts`),
 * where a present `undefined` suppresses the chart-level value and an absent key inherits it.
 * Suppressing is the intent: the subchart supplied an override the schema rejects, and the
 * chart-level value is one it never asked for — inheriting it would be a different chart than
 * either party described. For a combo chart the chart-level pass cannot vet these itself, since
 * `_type` is an array there and every type-keyed correction matches no branch.
 */
export type ChartOptsOverrides = MaybeUndefined<ChartOptsInternal>
/**
 * One subchart of a combo chart, as `addChartDefinition` rebuilt it: series normalized onto
 * {@link OptsChartDataInternal} and options narrowed to the vetted {@link ChartOptsOverrides}.
 * `ChartOptsInternal._type` holds these, never the caller's own {@link ChartMulti} objects.
 */
export interface ChartMultiInternal extends Omit<ChartMulti, 'data' | 'options'> {
	data: OptsChartDataInternal[]
	options: ChartOptsOverrides
}
export interface SlideRelChart extends OptsChartData {
	type: CHART_NAME | ChartMultiInternal[]
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
	/**
	 * Embedded OLE object part (`addOleObject`): the `.rels` `Type` URI to emit — `.../package` for
	 * an embedded OPC package (xlsx/docx/pptx, themselves zips) or `.../oleObject` for a generic OLE
	 * server blob (ECMA-376 Part 1 §15.2.10). Present only on OLE payload rels, where it takes
	 * precedence over the image/audio/video branch in `slideObjectRelationsToXml` and `type` carries
	 * the part's content type verbatim for `[Content_Types].xml`. The object's preview picture is a
	 * separate, perfectly ordinary image rel.
	 */
	oleRelType?: string
	/**
	 * Embedded 3D model part (`addModel3d`): the `.rels` `Type` URI to emit — the Microsoft
	 * `…/office/2017/06/relationships/model3d` type, which no `includes('image'|'audio'|'video')`
	 * sniff would ever reach. Present only on the `.glb` payload rel, where it takes precedence
	 * over the image/audio/video branch in `slideObjectRelationsToXml` and `type` carries the
	 * part's content type verbatim for `[Content_Types].xml`. The model's preview picture is a
	 * separate, perfectly ordinary image rel. Same escape hatch as {@link oleRelType}.
	 */
	model3dRelType?: string
	rId: number
	/** Unescaped — see {@link SlideRel.Target}. Doubles as the zip entry name for embedded media. */
	Target: string
}
export interface SlideObject {
	_type: SlideObjectType
	options?: ObjectOptionsInternal
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
	// oleObject: resolved payload/preview rIds + `p:oleObj` attrs for the `<p:graphicFrame>` emitter
	ole?: OleInternal
	// model3d: resolved payload/preview rIds + camera/scale for the `<p:graphicFrame>` emitter
	model3d?: Model3dInternal
}
/**
 * Resolved 3D-model payload carried on a {@link SlideObject} until `gen/slide/objects/model3d.ts`
 * emits it. Camera values are already in `am3d`'s wire units so the emitter does no arithmetic:
 * positions in 1/36,000,000ths of a metre, `fov` in 60000ths of a degree, and
 * `meterPerModelUnitN` over a fixed denominator of 1,000,000.
 */
export interface Model3dInternal {
	/** rId of the embedded `.glb` part rel (see {@link SlideRelMedia.model3dRelType}). */
	modelRid: number
	/** rId of the preview picture's image rel — shared by `am3d:blip` and the `mc:Fallback` `p:pic`. */
	previewRid: number
	/** `am3d:camera/am3d:pos` (`@x`/`@y`/`@z`). */
	pos: { x: number; y: number; z: number }
	/** `am3d:camera/am3d:lookAt` (`@x`/`@y`/`@z`). */
	lookAt: { x: number; y: number; z: number }
	/** `am3d:camera/am3d:up` (`@dx`/`@dy`/`@dz`). */
	up: { dx: number; dy: number; dz: number }
	/** `am3d:perspective@fov`, in 60000ths of a degree. */
	fov: number
	/** `am3d:meterPerModelUnit@n`; `@d` is always 1,000,000. */
	meterPerModelUnitN: number
}
/** Resolved OLE payload carried on a {@link SlideObject} until `gen/slide/objects/ole.ts` emits it. */
export interface OleInternal {
	/** rId of the embedded object part rel (see {@link SlideRelMedia.oleRelType}). */
	objectRid: number
	/** rId of the preview picture's image rel (`p:pic > a:blip@r:embed`). */
	previewRid: number
	/** `p:oleObj@progId` — the OLE server PowerPoint launches on double-click. */
	progId: string
	/** `p:oleObj@name` — the object's kind as PowerPoint labels it (`Worksheet`, `Document`, …). */
	name: string
	/** `p:oleObj@showAsIcon`; omitted when false, matching PowerPoint. */
	showAsIcon: boolean
	/** `p:oleObj@imgW`/`@imgH` (EMU). Undefined means "use the frame's own extent". */
	imgW?: number
	imgH?: number
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
/** Resolved zoom payload carried on a {@link SlideObject} until `gen/slide/objects/zoom.ts` emits it. */
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
	_relsChart: SlideRelChart[] // needed as we use args:"Slide|SlideLayout" often
	_relsMedia: SlideRelMedia[] // needed as we use args:"Slide|SlideLayout" often
	_relsNotes?: SlideRel[] // hyperlink rels emitted in the notes-slide part (notesSlideN.xml.rels)
	_comments?: SlideComment[] // review comments emitted in the per-slide comments part (commentN.xml)
	_txStyles?: MasterTextStyleProps // per-level master text styles emitted in slideMaster1.xml <p:txStyles> (deck-wide; set via defineSlideMaster textStyles)
	_slideNum: number
	_slideNumberProps?: SlideNumberProps | null
	_slideObjects: SlideObject[]
	/**
	 * Per-kind counters backing default Selection Pane names (`Shape 1`, `Image 1`, `Group 1`, …),
	 * keyed by `SlideObjectType`. Monotonic for the life of the slide: an object consumes its index
	 * when it is added, whether it stays top-level or is moved into a group's `_groupObjects`.
	 * Lazily created by `nextObjectNameIdx` (`gen/define/object-name.ts`).
	 */
	_objectNameCounts?: Partial<Record<SlideObjectType, number>>

	/**
	 * `| undefined` for the same reason {@link Slide.background} carries it: `SlideBuilder`
	 * implements this as an accessor pair, so it is always *present* on a real slide or layout and
	 * answers `undefined` when none is set. The `_`-prefixed members above are plain fields and
	 * keep the strict form.
	 */
	background?: BackgroundOption | undefined
}
export interface SlideLayoutInternal extends SlideBaseProps, SlideLayout {
	_slide?: {
		_bkgdImgRid?: number
		back: string
		color: string
		hidden?: boolean
	} | null
}
export interface PresSlideInternal extends SlideBaseProps, Slide {
	/** Narrowed so the export-time `_sndRId` stamp has somewhere to live off the public type. */
	transition?: TransitionPropsInternal | undefined
	_rId: number
	_slideLayout: SlideLayoutInternal | null
	_slideId: number
	/** Preset build animations on this slide, in play order (see {@link Slide.addAnimation}). */
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
	/** Author-side embedded fonts (see {@link TsPptx.embedFont}); empty when none. */
	embeddedFonts: EmbeddedFont[]
}
