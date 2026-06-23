/**
 * PptxGenJS: Slide Class
 */

import { CHART_NAME, SHAPE_NAME } from './core-enums.js'
import {
	AddSlideProps,
	BackgroundProps,
	ConnectorProps,
	GroupChildProps,
	GroupProps,
	HexColor,
	IChartMulti,
	IChartOpts,
	IChartOptsLib,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ImageProps,
	MediaProps,
	NotesProps,
	PresLayout,
	PresSlide,
	PresSlideInternal,
	ShapeProps,
	SlideLayoutInternal,
	SlideNumberProps,
	OptsChartData,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
} from './core-interfaces.js'
import * as genObj from './gen-objects.js'
import { emuToInches } from './units.js'

export default class Slide {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (options?: AddSlideProps) => PresSlideInternal
	public getSlide: (slideNum: number) => PresSlideInternal
	public _name: string
	public _presLayout: PresLayout
	public _rels: ISlideRel[]
	public _relsChart: ISlideRelChart[]
	public _relsMedia: ISlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayoutInternal | null
	public _slideNum: number
	public _slideNumberProps: SlideNumberProps | null
	public _slideObjects: ISlideObject[]
	public _newAutoPagedSlides: PresSlideInternal[] = []

	constructor(params: {
		addSlide: (options?: AddSlideProps) => PresSlideInternal
		getSlide: (slideNum: number) => PresSlideInternal
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout?: SlideLayoutInternal
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
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
	 * Background color
	 * @type {string|BackgroundProps}
	 * @deprecated in v3.3.0 - use `background` instead
	 */
	private _bkgd?: string | BackgroundProps
	public set bkgd(value: string | BackgroundProps) {
		this._bkgd = value
		if (!this._background || !this._background.color) {
			if (!this._background) this._background = {}
			if (typeof value === 'string') this._background.color = value
		}
	}

	public get bkgd(): string | BackgroundProps | undefined {
		return this._bkgd
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 * @since v3.3.0
	 */
	private _background?: BackgroundProps
	public set background(props: BackgroundProps) {
		this._background = props
		// Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props) genObj.addBackgroundDefinition(props, this)
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
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {Slide} this Slide
	 */
	addChart(type: CHART_NAME, data: OptsChartData[], options?: IChartOpts): Slide
	addChart(type: IChartMulti[], options?: IChartOpts): Slide
	addChart(type: CHART_NAME | IChartMulti[], dataOrOptions: OptsChartData[] | IChartOpts = [], options?: IChartOpts): Slide {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		// Set `_type` on IChartOptsLib as its what is used as object is passed around
		const optionsWithType = (Array.isArray(type) && !Array.isArray(dataOrOptions) ? dataOrOptions : options) as IChartOptsLib | undefined
		if (optionsWithType) optionsWithType._type = type
		genObj.addChartDefinition(this, type, dataOrOptions, options)
		return this
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {Slide} this Slide
	 */
	addImage(options: ImageProps): Slide {
		genObj.addImageDefinition(this, options)
		return this
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		genObj.addMediaDefinition(this, options)
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string | NotesProps | NotesProps[]} notes - notes text, or rich runs with inline
	 * formatting / hyperlinks. A plain string is the single-run case; pass run objects to add
	 * hyperlinks (external `url` only) or per-run bold/italic/underline/color/fontSize/fontFace.
	 * @example slide.addNotes('Remember to smile')
	 * @example slide.addNotes([{ text: 'See ' }, { text: 'the docs', options: { hyperlink: { url: 'https://gitbrent.github.io/PptxGenJS/' } } }])
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string | NotesProps | NotesProps[]): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Slide} this Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		// let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, shapeName, options || {})
		return this
	}

	/**
	 * Group slide objects into a single PowerPoint group (`<p:grpSp>`).
	 *
	 * Children keep their slide-absolute `x/y/w/h` (identity child coordinate space at every depth),
	 * and the objects become one selectable/movable group in PowerPoint. A `group` child nests
	 * another group. When `options.x/y/w/h` are omitted the group's bounds are the bounding box of
	 * its children (recursing into nested groups). Charts, media, tables, and placeholders are not
	 * supported as group children yet (each is skipped with a warning).
	 * @param {GroupChildProps[]} children - child object descriptors (`{ text }`, `{ image }`, `{ shape }`, `{ rect }`, `{ roundRect }`, `{ line }`, `{ group }`)
	 * @param {GroupProps} options - group position/size/name options
	 * @return {Slide} this Slide
	 * @example slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } }, { text: { text: 'Hi', options: { x: 1, y: 1, w: 2, h: 1 } } }])
	 * @example slide.addGroup([{ rect: { x: 1, y: 1, w: 4, h: 3 } }, { group: { children: [{ text: { text: 'Hi', options: { x: 1.5, y: 1.5, w: 2, h: 1 } } }] } }])
	 */
	addGroup(children: GroupChildProps[], options?: GroupProps): Slide {
		genObj.addGroupDefinition(this, children, options || {})
		return this
	}

	/**
	 * Add a connector (a line drawn between two points, emitted as a PowerPoint `<p:cxnSp>`).
	 * @param {ConnectorProps} options - connector endpoints (`x1,y1,x2,y2`) and line styling
	 * @return {Slide} this Slide
	 * @example slide.addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, endArrowType: 'triangle' })
	 */
	addConnector(options: ConnectorProps): Slide {
		genObj.addConnectorDefinition(this, options)
		return this
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		this._newAutoPagedSlides = genObj.addTableDefinition(this, tableRows, options || {}, this._slideLayout, this._presLayout, this.addSlide, this.getSlide)
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
		genObj.addTextDefinition(this, textParam, options || {}, false)
		return this
	}
}
