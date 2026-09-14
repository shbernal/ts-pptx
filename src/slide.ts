/**
 * ts-pptx: SlideBuilder — write-side implementation of the public `Slide` interface
 *
 * What a slide is, and nothing about what can go on one. Every `add*` method comes from the
 * construct family that owns it, bound onto the instance from the list the presentation was
 * composed with (`families/shared.ts`); this class owns the slide's own state — its rels, its
 * objects, its number, its background — and the geometry accessors over it.
 *
 * The methods are bound rather than written here because a class method body is never shaken: one
 * that called `addChartDefinition` would put every plot module under `gen/chart/` in the graph of
 * any program that can make a slide, which is all of them.
 */

import type {
	AnimationProps,
	BackgroundOption,
	HexColor,
	SlideComment,
	PresLayout,
	SlideNumberProps,
	SlideObjectInfo,
	TransitionProps,
} from './types/index.js'
import type { Slide } from './types/slide.js'
import type {
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
import { isGroupableObject } from './gen/define/group.js'
import {
	familyMethodUnavailable,
	SLIDE_METHOD_FAMILIES,
	type ChildAuthors,
	type SlideAuthorMethod,
	type SlideAuthors,
} from './families/shared.js'

/**
 * Project one authored render-object onto the public {@link SlideObjectInfo}, recursing into a
 * group's children.
 *
 * The name is the stored one, which is the caller's own spelling: only `cNvPrOpen` escapes it. So
 * a name read here resolves when it is handed back to `groupObjects()`.
 *
 * Every `add*Definition` also *generates* a name when the caller passed none, so the fallback below
 * is unreachable through the public API. It is there because the field is typed non-optional
 * internally only by convention, and an empty name would otherwise read back as `undefined` from a
 * surface that promises a string.
 */
function toSlideObjectInfo(obj: SlideObject): SlideObjectInfo {
	const stored = obj.options?.objectName
	return {
		type: obj._type,
		objectName: typeof stored === 'string' && stored.length > 0 ? stored : '',
		isPlaceholder: !!obj.options?.placeholder,
		canGroup: isGroupableObject(obj),
		children: (obj._groupObjects ?? []).map(toSlideObjectInfo),
	}
}

export default class SlideBuilder {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (layout: SlideLayoutInternal | null) => PresSlideInternal
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

	/**
	 * The child-descriptor authors this presentation was composed with, for the group family: a
	 * group's children resolve through the same table a slide master's `objects` do, so a group can
	 * hold exactly the kinds this presentation can author.
	 */
	public readonly _childAuthors: ChildAuthors

	/**
	 * The authoring methods, one per construct family, bound in the constructor.
	 *
	 * Declared here and assigned there: the signature each one keeps is the public one on
	 * {@link Slide}, and a method the presentation was not composed with still answers — with the
	 * name of the family it needs, rather than `undefined is not a function`.
	 */
	declare public addAnimation: Slide['addAnimation']
	declare public addChart: Slide['addChart']
	declare public addComment: Slide['addComment']
	declare public addConnector: Slide['addConnector']
	declare public addGroup: Slide['addGroup']
	declare public addImage: Slide['addImage']
	declare public addMedia: Slide['addMedia']
	declare public addModel3d: Slide['addModel3d']
	declare public addNotes: Slide['addNotes']
	declare public addOleObject: Slide['addOleObject']
	declare public addSectionZoom: Slide['addSectionZoom']
	declare public addShape: Slide['addShape']
	declare public addSlideZoom: Slide['addSlideZoom']
	declare public addSummaryZoom: Slide['addSummaryZoom']
	declare public addTable: Slide['addTable']
	declare public addText: Slide['addText']
	declare public groupObjects: Slide['groupObjects']

	constructor(params: {
		addSlide: (layout: SlideLayoutInternal | null) => PresSlideInternal
		getSlide: (slideNum: number) => PresSlideInternal | undefined
		getSections?: () => SectionInternalProps[]
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout?: SlideLayoutInternal
		authors: SlideAuthors
		childAuthors: ChildAuthors
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
		this._childAuthors = params.childAuthors
		this.#bindFamilyMethods(params.authors)
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout?._slideNumberProps ? this._slideLayout._slideNumberProps : null
	}

	/**
	 * Put this presentation's family methods on the slide.
	 *
	 * Per instance rather than on the prototype: the family set belongs to the presentation, not to
	 * the class, and two presentations composed differently in one process each get their own. A
	 * slide count is small enough that per-instance function properties are not a memory question.
	 *
	 * Chaining is applied here, once, rather than by each author: every `add*` answers the slide, so
	 * an author has nothing to say about the return value and cannot get it wrong.
	 */
	#bindFamilyMethods(authors: SlideAuthors): void {
		const slide = this as unknown as Record<SlideAuthorMethod, unknown>
		for (const method of Object.keys(SLIDE_METHOD_FAMILIES) as SlideAuthorMethod[]) {
			const author = authors[method] as ((target: SlideBuilder, ...args: never[]) => void) | undefined
			slide[method] = author
				? (...args: never[]): SlideBuilder => {
						author(this, ...args)
						return this
					}
				: (): never => familyMethodUnavailable(`slide.${method}`, SLIDE_METHOD_FAMILIES[method])
		}
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 */
	private _background?: BackgroundOption | undefined
	public set background(props: BackgroundOption | undefined) {
		this._background = props
		// Every assignment reaches the definer, `undefined` included: it is what drops an image an
		// earlier assignment registered. Skipping it left `background = undefined` reading back as
		// no background while the slide still painted the image. (Image data/path must be captured
		// before the package is built.)
		addBackgroundDefinition(props, this)
	}

	public get background(): BackgroundOption | undefined {
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
	// `| undefined` because the setter takes one: `slide.transition = undefined` is how a
	// transition is removed, and the backing field has to be able to hold what the setter accepts.
	private _transition?: TransitionProps | undefined
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

	/**
	 * Every continuation slide this slide's tables have spilled onto, in call order, each once.
	 *
	 * Across calls, not per call: a slide can carry more than one `addTable`, and each auto-paging
	 * one appends whatever it spilled onto. The list was assigned rather than appended, so a second
	 * table erased the first table's report while its slides stayed in the deck. "Spilled onto"
	 * rather than "created" because a later table lands on the earlier one's continuations instead
	 * of making its own, and those are reported once.
	 */
	public get newAutoPagedSlides(): Slide[] {
		return this._newAutoPagedSlides
	}

	/**
	 * The objects authored on this slide so far, bottom-to-top in z-order.
	 *
	 * Built fresh on every access rather than cached: the underlying list is mutated in place by
	 * `addGroup`/`groupObjects` as well as appended to, so a cache would have to be invalidated from
	 * every authoring path and would go stale the first time one was added without remembering.
	 * Building it costs a walk of a list a slide-sized deck keeps in the dozens.
	 */
	public get objects(): readonly SlideObjectInfo[] {
		return this._slideObjects.map(toSlideObjectInfo)
	}

	/** Slide width in inches (resolved from the active presentation layout). */
	public get width(): number {
		return emuToInches(this._presLayout.width)
	}

	/** Slide height in inches (resolved from the active presentation layout). */
	public get height(): number {
		return emuToInches(this._presLayout.height)
	}
}
