/**
 * PptxGenJS: Slide Class
 */

import { asChartType, type CHART_NAME, type SHAPE_NAME } from './core-enums.js'
import type {
	AddSlideProps,
	AnimationProps,
	BackgroundProps,
	CommentProps,
	ConnectorProps,
	GroupChildProps,
	GroupProps,
	HexColor,
	ChartMulti,
	ChartOpts,
	SlideComment,
	ImageProps,
	MediaProps,
	NotesProps,
	PresLayout,
	PresSlide,
	ShapeProps,
	SlideNumberProps,
	OptsChartData,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
	TransitionProps,
	SlideZoomProps,
	SectionZoomProps,
	SummaryZoomProps,
} from './core-interfaces.js'
import type {
	ChartOptsInternal,
	SlideObject,
	SlideRel,
	SlideRelChart,
	SlideRelMedia,
	PresSlideInternal,
	SlideLayoutInternal,
	SectionInternalProps,
} from './types/internal.js'
import { emuToInches } from './units.js'
import { addBackgroundDefinition } from './gen/define/background.js'
import { addChartDefinition } from './gen/define/chart.js'
import { addCommentDefinition } from './gen/define/comment.js'
import { addConnectorDefinition } from './gen/define/connector.js'
import { addGroupDefinition, groupObjectsDefinition } from './gen/define/group.js'
import { addImageDefinition } from './gen/define/image.js'
import { addMediaDefinition } from './gen/define/media.js'
import { addNotesDefinition } from './gen/define/notes.js'
import { addShapeDefinition } from './gen/define/shape.js'
import { addTableDefinition } from './gen/define/table.js'
import { addTextDefinition } from './gen/define/text.js'
import { addSectionZoomDefinition, addSlideZoomDefinition, addSummaryZoomDefinition } from './gen/define/zoom.js'

/** Distinguish a multi-type (combo) chart array (`ChartMulti[]`) from a single chart's data (`OptsChartData[]`). */
function isMultiChart(arg: OptsChartData[] | ChartMulti[]): arg is ChartMulti[] {
	const first = arg[0] as Partial<ChartMulti> | undefined
	return !!first && typeof first === 'object' && 'type' in first && 'data' in first
}

export default class Slide {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (options?: AddSlideProps) => PresSlideInternal
	public getSlide: (slideNum: number) => PresSlideInternal | undefined
	public getSections: () => SectionInternalProps[]
	public _name: string
	public _presLayout: PresLayout
	public _rels: SlideRel[]
	public _relsChart: SlideRelChart[]
	public _relsMedia: SlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayoutInternal | null
	public _slideNum: number
	public _slideNumberProps: SlideNumberProps | null
	public _slideObjects: SlideObject[]
	public _comments: SlideComment[] = []
	public _newAutoPagedSlides: PresSlideInternal[] = []
	public _animations: AnimationProps[] = []

	constructor(params: {
		addSlide: (options?: AddSlideProps) => PresSlideInternal
		getSlide: (slideNum: number) => PresSlideInternal | undefined
		getSections?: () => SectionInternalProps[]
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout?: SlideLayoutInternal
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this.getSections = params.getSections ?? (() => [])
		this._name = `Slide ${params.slideNumber}`
		this._presLayout = params.presLayout
		this._rId = params.slideRId
		this._rels = []
		this._relsChart = []
		this._relsMedia = []
		this._setSlideNum = params.setSlideNum
		this._slideId = params.slideId
		this._slideLayout = params.slideLayout || null
		this._slideNum = params.slideNumber
		this._slideObjects = []
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout?._slideNumberProps ? this._slideLayout._slideNumberProps : null
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 */
	private _background?: BackgroundProps
	public set background(props: BackgroundProps) {
		this._background = props
		// Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props) addBackgroundDefinition(props, this)
	}

	public get background(): BackgroundProps | undefined {
		return this._background
	}

	/**
	 * Default font color
	 * @type {HexColor}
	 */
	private _color?: HexColor
	public set color(value: HexColor) {
		this._color = value
	}

	public get color(): HexColor | undefined {
		return this._color
	}

	/**
	 * @type {boolean}
	 */
	private _hidden = false
	public set hidden(value: boolean) {
		this._hidden = value
	}

	public get hidden(): boolean {
		return this._hidden
	}

	/**
	 * Slide-show transition (`p:transition`) played when advancing to this slide.
	 * @type {TransitionProps}
	 */
	private _transition?: TransitionProps
	public set transition(value: TransitionProps | undefined) {
		this._transition = value
	}

	public get transition(): TransitionProps | undefined {
		return this._transition
	}

	/**
	 * @type {SlideNumberProps}
	 */
	public set slideNumber(value: SlideNumberProps) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumberProps = value
		this._setSlideNum(value)
	}

	public get slideNumber(): SlideNumberProps | undefined {
		return this._slideNumberProps ?? undefined
	}

	public get newAutoPagedSlides(): PresSlide[] {
		return this._newAutoPagedSlides
	}

	/** Slide width in inches (resolved from the active presentation layout). */
	public get width(): number {
		return emuToInches(this._presLayout.width)
	}

	/** Slide height in inches (resolved from the active presentation layout). */
	public get height(): number {
		return emuToInches(this._presLayout.height)
	}

	/**
	 * Add chart to Slide
	 * @param {OptsChartData[]} data - chart data
	 * @param {ChartOpts & { type: CHART_NAME }} options - chart options; `type` is required here
	 * @return {Slide} this Slide
	 */
	addChart(data: OptsChartData[], options: ChartOpts & { type: CHART_NAME }): Slide
	/**
	 * Add a multi-type (combo) chart to Slide
	 * @param {ChartMulti[]} charts - per-type chart definitions (each carries its own `type`/`data`)
	 * @param {ChartOpts} options - shared chart options
	 * @return {Slide} this Slide
	 */
	addChart(charts: ChartMulti[], options?: ChartOpts): Slide
	addChart(arg1: OptsChartData[] | ChartMulti[], arg2?: ChartOpts & { type?: CHART_NAME }): Slide {
		let type: CHART_NAME | ChartMulti[]
		let data: OptsChartData[]
		let options: ChartOpts

		if (Array.isArray(arg1) && isMultiChart(arg1)) {
			// Multi-type (combo) chart: addChart(ChartMulti[], options?)
			type = arg1
			data = []
			options = arg2 ?? {}
		} else {
			// Canonical single-type form: addChart(data, { type, ...options })
			data = arg1 ?? []
			options = arg2 ?? {}
			const optType = (options as ChartOpts & { type?: CHART_NAME }).type
			if (!optType) {
				throw new Error(
					'addChart: a chart `type` is required on the options object, e.g. addChart(data, { type: pptx.ChartType.bar }).'
				)
			}
			type = optType
		}

		// Set `_type` on ChartOptsInternal as it is what is used as the object is passed around
		;(options as ChartOptsInternal)._type = Array.isArray(type) ? type : asChartType(type)
		// addChartDefinition's multi-type branch reads the shared options from its `data` slot
		if (Array.isArray(type)) {
			addChartDefinition(this, type, options, undefined)
		} else {
			addChartDefinition(this, type, data, options)
		}
		return this
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {Slide} this Slide
	 */
	addImage(options: ImageProps): Slide {
		addImageDefinition(this, options)
		return this
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @param {string | NotesProps | NotesProps[]} notes - notes text, or rich runs with inline
	 * formatting / hyperlinks. A plain string is the single-run case; pass run objects to add
	 * hyperlinks (external `url` only) or per-run bold/italic/underline/color/fontSize/fontFace.
	 * @example slide.addNotes('Remember to smile')
	 * @example slide.addNotes([{ text: 'See ' }, { text: 'the docs', options: { hyperlink: { url: 'https://example.com/' } } }])
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string | NotesProps | NotesProps[]): Slide {
		addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add a review comment to the Slide (legacy PowerPoint comment).
	 * Comments by the same author (name + initials) are grouped under one author entry in the deck.
	 * @param {CommentProps} options - comment author, text, and optional marker position/date
	 * @return {Slide} this Slide
	 * @example slide.addComment({ author: 'Ada Lovelace', text: 'Tighten this headline', x: 1, y: 0.5 })
	 */
	addComment(options: CommentProps): Slide {
		addCommentDefinition(this, options)
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Slide} this Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide {
		// `shapeName` is a plain string preset name (e.g. `pptxgen.ShapeType.rect` === "rect").
		addShapeDefinition(this, shapeName, options || {})
		return this
	}

	/**
	 * Group slide objects into a single PowerPoint group (`<p:grpSp>`).
	 *
	 * Children keep their slide-absolute `x/y/w/h` (identity child coordinate space at every depth),
	 * and the objects become one selectable/movable group in PowerPoint. A `group` child nests
	 * another group. The group's frame is all-or-nothing: pass all four of `options.x/y/w/h`, or none
	 * to have the bounds be the bounding box of the children (recursing into nested groups). A
	 * partial frame warns and uses auto-bounds. Charts, media, tables, and placeholders are not
	 * supported as group children yet (each is skipped with a warning).
	 * @param {GroupChildProps[]} children - child object descriptors (`{ text }`, `{ image }`, `{ shape }`, `{ rect }`, `{ roundRect }`, `{ line }`, `{ group }`)
	 * @param {GroupProps} options - group position/size/name options
	 * @return {Slide} this Slide
	 * @example slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } }, { text: { text: 'Hi', options: { x: 1, y: 1, w: 2, h: 1 } } }])
	 * @example slide.addGroup([{ rect: { x: 1, y: 1, w: 4, h: 3 } }, { group: { children: [{ text: { text: 'Hi', options: { x: 1.5, y: 1.5, w: 2, h: 1 } } }] } }])
	 */
	addGroup(children: GroupChildProps[], options?: GroupProps): Slide {
		addGroupDefinition(this, children, options || {})
		return this
	}

	/**
	 * Group objects already added to this slide into a single PowerPoint group (`<p:grpSp>`),
	 * addressed by their `objectName`. The counterpart to {@link Slide.addGroup} for slides composed
	 * from independent renderers, where the objects exist already and replaying their descriptors just
	 * to group them is not practical.
	 *
	 * Grouping is visually a no-op: children keep their slide-absolute geometry and their relative
	 * z-order, and the group takes the topmost member's former slot in the stack. Name the objects in
	 * any order — z-order decides the children's order, not the array. Groups may be named, so groups
	 * can be nested into larger logical groups. Charts, media, tables, and placeholders cannot be
	 * grouped yet.
	 *
	 * Unlike `addGroup()`, every problem throws: a name that matches nothing, matches an object
	 * already inside another group, is ambiguous across two same-named objects, or names an
	 * ungroupable kind. Each would otherwise leave the object loose on the slide, silently.
	 * @param {string[]} objectNames - `objectName`s of the top-level objects to group
	 * @param {GroupProps} options - group position/size/name options (frame is all-or-nothing, as with `addGroup`)
	 * @return {Slide} this Slide
	 * @example slide.addText('Hi', { x: 1, y: 1, w: 2, h: 1, objectName: 'Caption' })
	 * @example slide.addImage({ path: 'logo.png', x: 1, y: 2, w: 2, h: 2, objectName: 'Logo' })
	 * @example slide.groupObjects(['Caption', 'Logo'], { objectName: 'Branding' })
	 */
	groupObjects(objectNames: string[], options?: GroupProps): Slide {
		groupObjectsDefinition(this, objectNames, options || {})
		return this
	}

	/**
	 * Add a connector (a line drawn between two points, emitted as a PowerPoint `<p:cxnSp>`).
	 * @param {ConnectorProps} options - connector endpoints (`x1,y1,x2,y2`) and line styling
	 * @return {Slide} this Slide
	 * @example slide.addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, endArrowType: 'triangle' })
	 */
	addConnector(options: ConnectorProps): Slide {
		addConnectorDefinition(this, options)
		return this
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		this._newAutoPagedSlides = addTableDefinition(
			this,
			tableRows,
			options || {},
			this._slideLayout,
			this._presLayout,
			this.addSlide,
			this.getSlide
		)
		return this
	}

	/**
	 * Add text to Slide
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {Slide} this Slide
	 */
	addText(text: string | number | TextProps[], options?: TextPropsOptions): Slide {
		const textParam = typeof text === 'string' || typeof text === 'number' ? [{ text, options }] : text
		addTextDefinition(this, textParam, options || {}, false)
		return this
	}

	/**
	 * Add a preset build animation (entrance/emphasis/exit) to a shape on this slide.
	 * Effects play in the order added and are grouped into click steps by `trigger`.
	 * Target the shape by its 0-based add order (`shapeIndex`) or by `objectName`.
	 * @param {AnimationProps} options - preset, target shape, trigger, and duration
	 * @return {Slide} this Slide
	 * @example slide.addAnimation({ preset: 'fadeIn', shapeIndex: 0 })
	 * @example slide.addAnimation({ preset: 'grow', objectName: 'logo', trigger: 'afterPrevious' })
	 */
	addAnimation(options: AnimationProps): Slide {
		this._animations.push(options)
		return this
	}

	/**
	 * Add a Slide Zoom — a clickable tile that zooms to a single target slide (Insert ▸ Zoom).
	 * The tile shows a neutral placeholder until PowerPoint regenerates the live thumbnail
	 * (once the target slide is next edited); pass `coverImage` to ship a fixed thumbnail.
	 * @param {SlideZoomProps} options - target slide (`Slide` or 1-based number), position, and options
	 * @return {Slide} this Slide
	 * @example slide.addSlideZoom({ target: intro, x: 1, y: 1, w: 3, h: 1.7 })
	 */
	addSlideZoom(options: SlideZoomProps): Slide {
		addSlideZoomDefinition(this, options)
		return this
	}

	/**
	 * Add a Section Zoom — a clickable tile that zooms to the start of a named section.
	 * @param {SectionZoomProps} options - target `sectionTitle`, position, and options
	 * @return {Slide} this Slide
	 * @example slide.addSectionZoom({ sectionTitle: 'Results', x: 1, y: 1, w: 3, h: 1.7 })
	 */
	addSectionZoom(options: SectionZoomProps): Slide {
		addSectionZoomDefinition(this, options, this.getSections())
		return this
	}

	/**
	 * Add a Summary Zoom — a grid of tiles, one per section (excluding this slide's own section),
	 * each zooming to that section's start.
	 * @param {SummaryZoomProps} options - grid position/size and options
	 * @return {Slide} this Slide
	 * @example slide.addSummaryZoom({ x: 0.5, y: 1.5, w: 11, h: 4.5 })
	 */
	addSummaryZoom(options: SummaryZoomProps): Slide {
		addSummaryZoomDefinition(this, options, this.getSections())
		return this
	}
}
