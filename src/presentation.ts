/**
 * ts-pptx: the presentation core
 *
 * Home of {@link PresentationCore}, the class every entry point subclasses with a runtime
 * adapter (`index.ts`, `node.ts`, `browser.ts`). See the class doc below for the layout of
 * its body and where package assembly actually happens.
 *
 * Units: this library takes user coordinates in inches and converts to EMU on output — see
 * `units.ts` for the conversions and `STANDARD_LAYOUTS` for the built-in slide sizes.
 *
 * Derived from PptxGenJS (C) 2015-present Brent Ely (MIT), which in turn took some code from
 * the OfficeGen project (Copyright 2013 Ziv Barber). This project is likewise MIT; the full
 * license text is in `LICENSE` at the repository root.
 */

import { warn } from './diagnostics.js'
import { InternalError, InvalidOptionError } from './errors.js'
import SlideBuilder from './slide.js'
import { DEF_PRES_LAYOUT, DEF_PRES_LAYOUT_NAME, DEF_SLIDE_MARGIN_IN } from './constants-internal.js'
import type {
	AddSlideProps,
	CustomPropertyValue,
	MeasureTextOptions,
	OverflowBoxOptions,
	PresLayout,
	SectionProps,
	SlideLayout,
	PackagePart,
	PartsProps,
	SlideMasterProps,
	SlideNumberProps,
	TableLayoutResult,
	TableProps,
	TableRow,
	TextMeasurement,
	TextProps,
	ThemeProps,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './types/index.js'
import type {
	PresentationPropsInternal,
	PresSlideInternal,
	SectionInternalProps,
	SlideLayoutInternal,
} from './types/internal.js'
import type { Slide } from './types/slide.js'
import type { RuntimeAdapter } from './runtime/types.js'
import { FontMetricsRegistry, parseFontMetrics } from './measure/font-metrics.js'
import { isFontCollection } from './measure/font-collection.js'
import { type EmbeddedFont, type EmbeddedFontSlot, EMBEDDED_FONT_SLOTS } from './embedded-fonts.js'
import { measureText } from './measure/fit.js'
import { computeTableLayout } from './measure/table-fit.js'
import { resolveFontBytes } from './font-source.js'
import { inchesToEmu, STANDARD_LAYOUTS, type StandardLayout } from './units.js'
import { clampRangedInput } from './units-internal.js'
import type { ExtractedSlides } from './read/api/presentation-types.js'

import { addBackgroundDefinition } from './gen/define/background.js'
import { createSlideMaster } from './gen/define/master.js'
import { getUuid } from './gen/utils.js'
import { extractSlides as extractSlidesFrom } from './gen/extract-slides.js'
import { buildPackageParts, writePackage, type PackageSource } from './package/assemble.js'
import { pickDefined, setOrClear } from './options-internal.js'

const VERSION = '3.7.0'

function standardLayoutToPresLayout(layout: StandardLayout): PresLayout {
	return {
		name: layout.name,
		width: layout.widthEmu,
		height: layout.heightEmu,
	}
}

/**
 * One side of a `defineLayout` definition, in inches, clamped to what `ST_SlideSizeCoordinate`
 * allows.
 *
 * `p:sldSz/@cx` and `@cy` are bounded 914400 to 51206400 EMU -- 1 to 56 inches -- and nothing
 * checked. Every arm of `defineLayout`'s guard chain is a `warn`, so
 * `{ name: 'Badge', width: 0.5, height: 0.5 }` produced no diagnostic at all (both values are
 * truthy and finite) and emitted a `sldSz` PowerPoint offers to repair; `width: -5` is truthy
 * too, so it reached the file as `cx="-4572000"`.
 *
 * `Number()` first, because the guard chain above deliberately accepts a numeric string as
 * advice rather than as an error, and `clampRangedInput` takes a number. A value with nothing
 * to coerce from is `NaN`, which it throws on -- as `coord/non-finite`, the code the unit
 * conversion used to raise one line later.
 * @param value - the caller's `width` or `height`
 * @param side - which one, for the diagnostic
 */
function layoutSideInches(value: number | undefined, side: 'width' | 'height'): number {
	return clampRangedInput(
		Number(value),
		1,
		56,
		'layout/size-out-of-range',
		`defineLayout ${side} (inches)`,
		'coord/non-finite'
	)
}

/**
 * Main presentation class and package export flow — the public entry point consumers
 * instantiate (`new TsPptx()`). Owns presentation-level state and metadata, collects
 * slides, and drives `write`/`writeFile`/`stream`. The actual OOXML string building is
 * delegated to the `gen-*` modules; runtime file/stream output goes through the injected
 * `RuntimeAdapter`. Enums (`AlignH`, `ChartType`, `ShapeType`, …) are imported from the
 * package entry, not read off the instance.
 *
 * Rough layout of the class body:
 *   - Metadata accessors      layout / author / company / title / theme / sections getters+setters
 *   - Private slide helpers    addNewSlide, getSlide, setSlideNumber
 *   - Public export methods    stream / write / writeFile — delegate packaging to `writePackage`
 *   - Public authoring methods addSlide / defineLayout / defineSlideMaster / defineTableStyle
 *
 * Package assembly (`[Content_Types].xml`, the rels graph, per-part XML, the ZIP) lives in
 * `package/assemble.ts`; this class provides the deck state via {@link PresentationCore.packageSource}.
 * The live-DOM `tableToSlides()` is NOT here — it is added by the browser entry subclass
 * (`browser.ts`) so it stays out of the Node build and out of the core chunk.
 */
export default class PresentationCore {
	// Property getters/setters

	/**
	 * Presentation layout name
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10"    x 7.5")
	 * - 'LAYOUT_16x9'  (10"    x 5.625")
	 * - 'LAYOUT_16x10' (10"    x 6.25")
	 * - 'LAYOUT_WIDE'  (13.333" x 7.5")
	 * Custom layouts:
	 * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	private _layout: string = DEF_PRES_LAYOUT
	public set layout(value: string | StandardLayout) {
		// Accept either a layout key string or a STANDARD_LAYOUTS preset object directly.
		const layoutKey = typeof value === 'string' ? value : value?.layout
		const newLayout: PresLayout | undefined = layoutKey ? this.LAYOUTS[layoutKey] : undefined

		if (newLayout) {
			this._layout = layoutKey
			this._presLayout = newLayout
		} else {
			throw new InvalidOptionError(
				'layout/unknown',
				`Unknown layout ${JSON.stringify(layoutKey)}; pass a registered layout name (see LAYOUTS) or a STANDARD_LAYOUTS preset`
			)
		}
	}

	public get layout(): string {
		return this._layout
	}

	/**
	 * ts-pptx Library Version
	 */
	private readonly _version: string = VERSION
	public get version(): string {
		return this._version
	}

	/**
	 * @type {string}
	 */
	private _author: string
	public set author(value: string) {
		this._author = value
	}

	public get author(): string {
		return this._author
	}

	/**
	 * @type {string}
	 */
	private _company: string
	public set company(value: string) {
		this._company = value
	}

	public get company(): string {
		return this._company
	}

	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	private _revision: string
	public set revision(value: string) {
		this._revision = value
	}

	public get revision(): string {
		return this._revision
	}

	/**
	 * @type {string}
	 */
	private _subject: string
	public set subject(value: string) {
		this._subject = value
	}

	public get subject(): string {
		return this._subject
	}

	/**
	 * @type {ThemeProps}
	 */
	private _theme: ThemeProps = {}
	public set theme(value: ThemeProps) {
		this._theme = value
	}

	public get theme(): ThemeProps {
		return this._theme
	}

	/**
	 * @type {string}
	 */
	private _title: string
	public set title(value: string) {
		this._title = value
	}

	public get title(): string {
		return this._title
	}

	/** Slide number shown on the first slide (maps to firstSlideNum in presentation.xml) */
	private _firstSlideNum: number
	public set firstSlideNum(value: number) {
		this._firstSlideNum = value
	}

	public get firstSlideNum(): number {
		return this._firstSlideNum
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	private _rtlMode: boolean
	public set rtlMode(value: boolean) {
		this._rtlMode = value
	}

	public get rtlMode(): boolean {
		return this._rtlMode
	}

	/** master slide layout object */
	private readonly _masterSlide: PresSlideInternal
	public get masterSlide(): Slide {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private readonly _slides: PresSlideInternal[]
	public get slides(): Slide[] {
		return this._slides
	}

	/** this Presentation's sections */
	private readonly _sections: SectionInternalProps[]
	public get sections(): SectionProps[] {
		return this._sections
	}

	/** custom document properties stored in docProps/custom.xml */
	private _customProperties: Array<{ name: string; value: CustomPropertyValue }>

	/**
	 * slide layout definition objects, used for generating slide layout files.
	 * `protected` (not `private`) so the browser entry subclass can resolve
	 * `tableToSlides({ masterTitle })` against them — that method lives on the
	 * browser build only (it reads a live DOM), see `browser.ts`.
	 */
	protected readonly _slideLayouts: SlideLayoutInternal[]
	public get slideLayouts(): SlideLayout[] {
		return this._slideLayouts
	}

	private get internalPresentation(): PresentationPropsInternal {
		return {
			author: this.author,
			company: this.company,
			firstSlideNum: this.firstSlideNum,
			layout: this.layout,
			masterSlide: this._masterSlide,
			presLayout: this.presLayout,
			revision: this.revision,
			rtlMode: this.rtlMode,
			sections: this._sections,
			slideLayouts: this._slideLayouts,
			slides: this._slides,
			subject: this.subject,
			theme: this.theme,
			title: this.title,
			embeddedFonts: this._embeddedFonts,
		}
	}

	private LAYOUTS: { [key: string]: PresLayout }

	/**
	 * The built-in default layout, or a throw. `LAYOUTS` is a string-keyed map, so the lookup is
	 * typed as possibly-missing at every call site even though the constructor registers it — and
	 * the two callers that need it were spelling the same guard and the same message.
	 */
	#requireDefaultLayout(): PresLayout {
		const defLayout = this.LAYOUTS[DEF_PRES_LAYOUT]
		if (!defLayout)
			throw new InternalError(
				'layout/default-not-registered',
				`Default presentation layout "${DEF_PRES_LAYOUT}" is not registered`
			)
		return defLayout
	}

	private _presLayout: PresLayout
	public get presLayout(): PresLayout {
		return this._presLayout
	}

	private readonly _runtime: RuntimeAdapter

	/** Write-side font metrics for measured text fit (`fit:'shrink'`). @see registerFontMetrics */
	private readonly _fontMetrics = new FontMetricsRegistry()

	/** Author-side embedded font faces, accumulated by {@link embedFont} and emitted at write time. */
	private readonly _embeddedFonts: EmbeddedFont[] = []

	constructor(runtime: RuntimeAdapter) {
		this._runtime = runtime
		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: standardLayoutToPresLayout(STANDARD_LAYOUTS.LAYOUT_4x3),
			LAYOUT_16x9: standardLayoutToPresLayout(STANDARD_LAYOUTS.LAYOUT_16x9),
			LAYOUT_16x10: standardLayoutToPresLayout(STANDARD_LAYOUTS.LAYOUT_16x10),
			LAYOUT_WIDE: standardLayoutToPresLayout(STANDARD_LAYOUTS.LAYOUT_WIDE),
		}

		// Core
		this._author = 'ts-pptx'
		this._company = 'ts-pptx'
		this._revision = '1' // Note: Must be a whole number
		this._subject = 'ts-pptx Presentation'
		this._title = 'ts-pptx Presentation'
		// ts-pptx props
		const defLayout = this.#requireDefaultLayout()
		this._presLayout = {
			name: defLayout.name,
			width: defLayout.width,
			height: defLayout.height,
		}
		this._firstSlideNum = 1
		this._rtlMode = false
		//
		this._slideLayouts = [
			{
				_margin: DEF_SLIDE_MARGIN_IN,
				_name: DEF_PRES_LAYOUT_NAME,
				_presLayout: this._presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
				_slide: null,
				_slideNum: 1000,
				_slideNumberProps: null,
				_slideObjects: [],
			},
		]
		this._slides = []
		this._sections = []
		this._customProperties = []
		this._masterSlide = {
			addChart: null,
			addComment: null,
			addConnector: null,
			addImage: null,
			addMedia: null,
			addNotes: null,
			addShape: null,
			addTable: null,
			addText: null,
			addAnimation: null,
			//
			_name: null,
			_animations: [],
			_presLayout: this._presLayout,
			_rId: null,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideId: null,
			_slideLayout: null,
			_slideNum: null,
			_slideNumberProps: null,
			_slideObjects: [],
			// Deliberately-partial internal stub: the master slide carries only rels/objects,
			// so its authoring methods and ids are intentionally null (never invoked on the master).
		} as unknown as PresSlideInternal
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {AddSlideProps} options - slide masterTitle and/or sectionTitle
	 * @return {Slide} new Slide
	 */
	private readonly addNewSlide = (options?: AddSlideProps): PresSlideInternal => {
		const nextOptions = options || {}
		// Preserve the originating slide's section for all auto-paged continuation slides.
		// Search for the section that owns the current last slide rather than assuming it is
		// the last section — the originating slide may not be at the tail of the deck.
		const lastSlide = this._slides[this._slides.length - 1]
		const sourceSection = this._sections.find((sect) => sect._slides.some((s) => s._slideNum === lastSlide?._slideNum))
		setOrClear(nextOptions, 'sectionTitle', sourceSection?.title)

		return this.addSlide(nextOptions) as PresSlideInternal
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {Slide} Slide
	 */
	private readonly getSlide = (slideNum: number): PresSlideInternal | undefined =>
		this._slides.find((slide) => slide._slideNum === slideNum)

	/**
	 * Provides the live section list to the `Slide` class so Section/Summary Zoom can resolve
	 * a section title to its stable GUID and enumerate sections in order.
	 * @return {SectionInternalProps[]} the presentation's sections
	 */
	private readonly getSections = (): SectionInternalProps[] => this._sections

	/**
	 * Enables the `Slide` class to set ts-pptx [Presentation] master/layout slidenumbers
	 * @param {SlideNumberProps} slideNum - slide number config
	 */
	private readonly setSlideNumber = (slideNum: SlideNumberProps): void => {
		// 1: Add slideNumber to slideMaster1.xml
		this._masterSlide._slideNumberProps = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		const defLayout = this._slideLayouts.find((layout) => layout._name === DEF_PRES_LAYOUT_NAME)
		if (defLayout) defLayout._slideNumberProps = slideNum
	}

	/**
	 * Assemble the slice of internal state the packaging layer reads. `PresentationCore` owns the
	 * authored deck; `writePackage` (in `package/assemble.ts`) turns it into OOXML parts.
	 */
	private packageSource(): PackageSource {
		return {
			runtime: this._runtime,
			presentation: this.internalPresentation,
			customProperties: this._customProperties,
			fontMetrics: this._fontMetrics,
		}
	}

	/**
	 * Author + serialize this presentation's slides as injectable descriptors,
	 * WITHOUT producing a `.pptx` package. Runs the same media-encode, placeholder
	 * backfill, and measured-fit passes `write()` uses, then serializes each slide
	 * body and resolves its image media to decoded bytes — so a loaded deck can
	 * splice the slides in via `Presentation.appendSlides()` (see `ts-pptx/read`)
	 * while keeping its own masters/layouts/theme byte-identical.
	 *
	 * Returns the deck's slide size (EMU, for the destination size check) and one
	 * descriptor per authored slide. Charts, audio/video media, online media, and
	 * internal slide-to-slide hyperlinks are surfaced as descriptors alongside the
	 * serialized body — this method does not resolve them into a package; the append
	 * path consumes each descriptor to reserve the parts and rebuild the rel graph.
	 */
	extractSlides = async (opts: { onMediaError?: 'throw' | 'placeholder' } = {}): Promise<ExtractedSlides> =>
		extractSlidesFrom(
			{ runtime: this._runtime, presentation: this.internalPresentation, fontMetrics: this._fontMetrics },
			opts
		)

	// FONT METRICS (measured text fit)

	/**
	 * Register a font's metrics so `fit:'shrink'` text boxes are measured and a real
	 * `fontScale` is baked at export time (text renders pre-shrunk in headless
	 * renderers and on plain file-open, with no manual edit/resize).
	 *
	 * Without registered metrics, `fit:'shrink'` keeps its current behavior (a bare
	 * `<a:normAutofit/>` that only PowerPoint recomputes on edit). Register the same
	 * face once per weight/style you use; bold/italic advances differ.
	 *
	 * **Font collections** (`.ttc`/`.otc`, how MS Gothic, Yu Gothic, SimSun, Microsoft
	 * YaHei and Cambria ship on Windows) hold several fonts in one file, so one has to be
	 * chosen. With no `font` option, `face` is used as the selector, which is usually what
	 * you meant; a `face` that names no font in the collection throws rather than falling
	 * back to the first, since measuring the wrong member is invisible downstream. Pass
	 * `font` when the deck-side name differs from the name inside the file, and use
	 * `listFontFaces` from `ts-pptx/measure` to see what a file holds.
	 * @param {string} face - font family name as used in `fontFace` (e.g. 'Aptos')
	 * @param {string | Uint8Array | ArrayBuffer} source - font file path/URL (Node/web) or raw TTF/OTF/TTC
	 *   bytes. A string is always a path or URL here; base64 text is `embedFont`'s `data`, not this.
	 * @param {object} [opts] - variant flags (advances differ per weight/style) and collection selection
	 * @param {boolean} [opts.bold] - these are the bold advances
	 * @param {boolean} [opts.italic] - these are the italic advances
	 * @param {number | string} [opts.font] - which font inside a collection: 0-based index, or a name
	 * @example await pptx.registerFontMetrics('Aptos', '/usr/share/fonts/Aptos.ttf')
	 * @example await pptx.registerFontMetrics('Aptos', aptosBoldBytes, { bold: true })
	 * @example await pptx.registerFontMetrics('MS PGothic', 'C:/Windows/Fonts/msgothic.ttc')
	 * @example await pptx.registerFontMetrics('Cambria Math', 'C:/Windows/Fonts/cambria.ttc', { font: 1 })
	 */
	async registerFontMetrics(
		face: string,
		source: string | Uint8Array | ArrayBuffer,
		opts?: { bold?: boolean; italic?: boolean; font?: number | string }
	): Promise<void> {
		// A `string` here is a path or URL, never base64 -- see `resolveFontBytes` for the two
		// spellings and why `embedFont` reads its `data` string the other way.
		const bytes = await resolveFontBytes(source, this._runtime, undefined, 'registerFontMetrics: `source`')
		// Only a collection is name-selected. A plain TTF holds one font whose internal
		// name need not match the deck-side family (registering Silkscreen's bytes under
		// any `face` has always worked), so defaulting `font` there would break that.
		const font = opts?.font ?? (isFontCollection(bytes) ? face : undefined)
		const metrics = await parseFontMetrics(bytes, font === undefined ? undefined : { font })
		this._fontMetrics.set(face, metrics, pickDefined(opts ?? {}, ['bold', 'italic']))
	}

	/**
	 * Embed a font face so the deck renders with it even on machines that do not
	 * have it installed. The **whole** face is embedded (not glyph-subset) under
	 * `/ppt/fonts/` and wired into `presentation.xml` (`p:embeddedFontLst` +
	 * `embedTrueTypeFonts="1"`), exactly as PowerPoint's "Embed fonts in the file"
	 * does. Call once per face/weight you use.
	 *
	 * The declared `typeface` MUST match the family name used in your run/`fontFace`
	 * typefaces (e.g. `addText('hi', { fontFace: 'Silkscreen' })`) or PowerPoint
	 * will not bind the embedded bytes to the text. Repeated calls with the same
	 * `typeface` and different `style` accumulate into one `p:embeddedFont` entry.
	 *
	 * Font licensing is the caller's responsibility: the bytes are embedded as
	 * handed over; `OS/2.fsType` embedding-permission bits are not enforced.
	 * @param {object} opts - font source + identity
	 * @param {string} [opts.path] - font file path (Node) or URL (web) to a `.ttf`/`.otf`
	 * @param {ArrayBuffer | Uint8Array | string} [opts.data] - raw font bytes, or a base64 string (bare or a
	 *   whole `data:` URL), in lieu of `path`. Unlike `registerFontMetrics`, a string here is never a path.
	 * @param {string} opts.typeface - family name as referenced by run/`fontFace` typefaces
	 * @param {EmbeddedFontSlot} [opts.style] - face slot; defaults to `'regular'`
	 * @example await pptx.embedFont({ path: '/fonts/Silkscreen-Regular.ttf', typeface: 'Silkscreen' })
	 * @example await pptx.embedFont({ path: '/fonts/Silkscreen-Bold.ttf', typeface: 'Silkscreen', style: 'bold' })
	 */
	async embedFont(opts: {
		path?: string
		data?: ArrayBuffer | Uint8Array | string
		typeface: string
		style?: EmbeddedFontSlot
	}): Promise<void> {
		if (!opts || typeof opts.typeface !== 'string' || opts.typeface.trim() === '') {
			throw new InvalidOptionError(
				'font/missing-typeface',
				'embedFont: `typeface` is required (the family name your runs reference)'
			)
		}
		const slot: EmbeddedFontSlot = opts.style ?? 'regular'
		if (!EMBEDDED_FONT_SLOTS.includes(slot)) {
			throw new InvalidOptionError(
				'font/invalid-style-slot',
				`embedFont: invalid style "${slot}"; expected one of ${EMBEDDED_FONT_SLOTS.join(', ')}`
			)
		}

		if (opts.path === undefined && opts.data === undefined)
			throw new InvalidOptionError('font/missing-source', 'embedFont: provide either `path` or `data`')

		// `path` is a path/URL; a `data` string is base64. That difference is the whole reason
		// `resolveFontBytes` takes a flag rather than guessing from the value.
		const bytes =
			typeof opts.path === 'string'
				? await resolveFontBytes(opts.path, this._runtime, undefined, 'embedFont: `path`')
				: await resolveFontBytes(opts.data, this._runtime, { base64: true }, 'embedFont: `data` string')

		// Accumulate faces of one family under a single embeddedFont entry; a repeat
		// of the same typeface+style replaces the prior bytes (last call wins).
		let font = this._embeddedFonts.find((f) => f.typeface === opts.typeface)
		if (!font) {
			font = { typeface: opts.typeface, faces: [] }
			this._embeddedFonts.push(font)
		}
		const existing = font.faces.find((f) => f.slot === slot)
		if (existing) existing.bytes = bytes
		else font.faces.push({ slot, bytes })
	}

	/**
	 * Measure how tall text wraps at a given width, using the **same** calibrated wrap
	 * model the export-time autofit bake uses — so a layout-time prediction matches the
	 * value `fit:'shrink'`/`'resize'` would bake. Synchronous: register the face's metrics
	 * first with {@link TsPptx.registerFontMetrics} (lookup is sync).
	 *
	 * Lets a consumer size its own geometry before export — grow a card to fit its text,
	 * reflow a grid, or detect overflow at layout time. Heights err **tall** (conservative),
	 * so the returned height is ≥ what PowerPoint/LibreOffice render.
	 * @param {string | TextProps[]} text - a string or run array (per-run options override the defaults in `opts`)
	 * @param {MeasureTextOptions} opts - width (inches), font size/face (points), and spacing
	 * @returns {TextMeasurement} laid-out height + line count + fit helpers
	 * @example const m = pptx.measureText('Long heading…', { wIn: 3, fontSize: 18, fontFace: 'Aptos' })
	 * @example if (pptx.measureText(runs, { wIn, fontSize, fontFace }).heightIn > cardHeightIn) growCard()
	 */
	measureText(text: string | TextProps[], opts: MeasureTextOptions): TextMeasurement {
		return measureText(this._fontMetrics, text, opts)
	}

	/**
	 * Convenience overflow check: `true` if `text` does not fit a box of inner size
	 * `wIn`×`hIn` (inches) at full size. Because the model errs tall, this is a
	 * **conservative** (slightly over-reporting) check — appropriate for a build-time
	 * warning, not a hard gate. An unmeasurable (unnamed theme-default) face reports
	 * `false` (no overflow) so the linter does not false-positive on faces it cannot measure.
	 * @param {string | TextProps[]} text - a string or run array
	 * @param {OverflowBoxOptions} opts - {@link MeasureTextOptions} plus the box inner height `hIn`
	 * @returns {boolean} true if the text overflows the box
	 */
	overflowsBox(text: string | TextProps[], opts: OverflowBoxOptions): boolean {
		const m = measureText(this._fontMetrics, text, opts)
		return m.measurable && !m.fitsBox(opts.hIn)
	}

	/**
	 * Compute the per-cell geometry of a table laid out at `opts.x`/`y`/`w`, without
	 * adding it to a slide — so a consumer can place images or shapes precisely over
	 * cells. Column widths (cell `x`/`w`) are exact, derived from the same logic the
	 * writer uses. Row heights (`y`/`h`) are exact when pinned by `rowH` (array or
	 * scalar) or table `h`; an auto-height row is estimated with the same conservative
	 * (tall) text model as {@link TsPptx.measureText} and flagged `heightExact:false`
	 * (register the cell font via {@link TsPptx.registerFontMetrics} for an exact
	 * estimate). Geometry is for a single, un-paginated table — `autoPage` paging is
	 * not modeled.
	 * @param {TableRow[]} rows - the same `rows` passed to `slide.addTable`
	 * @param {TableProps} opts - the same table options (`x`, `y`, `w`, `colW`, `rowH`, `h`, …)
	 * @returns {TableLayoutResult} per-cell rectangles (inches) plus overall table bounds
	 * @example
	 * const g = pptx.tableLayout(rows, { x: 1, y: 1, w: 8, colW: [2, 3, 3] })
	 * const c = g.cells.find(c => c.row === 0 && c.col === 2)
	 * slide.addImage({ path: 'logo.png', x: c.xIn, y: c.yIn, w: c.wIn, h: c.hIn })
	 */
	tableLayout(rows: TableRow[], opts: TableProps): TableLayoutResult {
		return computeTableLayout(rows, opts, this._presLayout, this._fontMetrics)
	}

	// EXPORT METHODS

	/**
	 * Export the current presentation as portable ZIP bytes. Unlike the former
	 * `stream()` method, this does not imply incremental output or convert the
	 * archive to a Node `Buffer`: package assembly is synchronous and the returned
	 * `Uint8Array` works unchanged in Node, browsers, and workers.
	 * @param {WriteBaseProps} props - output properties
	 * @returns {Promise<Uint8Array>} complete `.pptx` archive bytes
	 */
	async toBytes(props?: WriteBaseProps): Promise<Uint8Array> {
		return (await writePackage(this.packageSource(), {
			...pickDefined(props ?? {}, ['compression', 'onMediaError']),
			outputType: 'uint8array',
		})) as Uint8Array
	}

	/**
	 * Export the current Presentation as ZIP content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	async write(props?: WriteProps): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		return await writePackage(
			this.packageSource(),
			pickDefined(props ?? {}, ['compression', 'outputType', 'onMediaError'])
		)
	}

	/**
	 * Assemble the current Presentation into its OOXML package parts WITHOUT zipping — the raw
	 * `path → bytes` contents `write()` would compress into the `.pptx`. Use this to stream parts
	 * into a custom container, inspect individual parts, or feed a pipeline that does its own
	 * archiving. To produce a real `.pptx` file, use {@link write} / {@link writeFile} instead.
	 *
	 * The returned parts are byte-identical, per part, to what `write()` emits for the same deck,
	 * and their order is the package's emission order. Both are a stability-guaranteed observable
	 * contract: adding a new part in a later release is backward-compatible as long as existing
	 * part paths and their relative order do not shift; renaming or reordering an existing part is
	 * a breaking change. Each call returns fresh `Uint8Array` views over per-call buffers (a new
	 * assembly runs each time), so callers may retain or transfer them without defensive copying.
	 * @param {PartsProps} props - assembly options (`onMediaError` only; compression/output shape
	 *   are zip concerns that do not apply to unzipped parts)
	 * @returns {Promise<PackagePart[]>} the package parts in emission order
	 */
	async toParts(props?: PartsProps): Promise<PackagePart[]> {
		const parts = await buildPackageParts(this.packageSource(), { onMediaError: props?.onMediaError })
		// Drop the internal `store` (DEFLATE) hint — an fflate-era zip optimization, not an OOXML fact.
		return parts.map(({ path, data }) => ({ path, data }))
	}

	/**
	 * Export the current Presentation.
	 * Write the generated presentation to disk (Node) or trigger a download (browser).
	 * @param {WriteFileProps} props - output file properties
	 * @returns {Promise<string>} the presentation name
	 */
	async writeFile(props?: WriteFileProps): Promise<string> {
		const rawName = props?.fileName ?? 'Presentation.pptx'
		const fileName = rawName.toLowerCase().endsWith('.pptx') ? rawName : `${rawName}.pptx`

		// `writeFileOutputType` is the archive type this runtime prefers where it has one (Node
		// asks for a Buffer); a runtime that states none leaves the key off, so `zipPackageParts`
		// applies its own default rather than being handed an `undefined` to interpret.
		const writeProps: WriteProps = pickDefined(props ?? {}, ['compression', 'onMediaError'])
		setOrClear(writeProps, 'outputType', this._runtime.writeFileOutputType ?? undefined)
		const data = await writePackage(this.packageSource(), writeProps)
		return await this._runtime.writeFile(fileName, data)
	}

	// PRESENTATION METHODS

	/**
	 * Set a custom document property stored in `docProps/custom.xml`.
	 * Calling with the same name replaces the existing value.
	 * @param name - property name
	 * @param value - string, integer/float number, boolean, or Date
	 */
	setCustomProperty(name: string, value: CustomPropertyValue): void {
		this._customProperties = this._customProperties.filter((p) => p.name !== name)
		this._customProperties.push({ name, value })
	}

	/**
	 * Add a new Section to Presentation
	 * @param {SectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: SectionProps): void {
		if (!section) {
			warn('section/missing-argument', 'addSection requires an argument')
			return
		} else if (!section.title) {
			warn('section/missing-title', 'addSection requires a title')
			return
		} else if (this._sections.some((sect) => sect.title === section.title)) {
			warn(
				'section/duplicate-title',
				`addSection: a section titled "${section.title}" already exists; ignoring duplicate`
			)
			return
		}

		const newSection: SectionInternalProps = {
			_type: 'user',
			_slides: [],
			_id: `{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}`,
			title: section.title,
		}

		if (section.order) this._sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} options - slide options
	 * @returns {Slide} the new Slide
	 */
	addSlide(options?: AddSlideProps): Slide {
		const masterTitle = options?.masterTitle ?? ''
		const defLayout = this.#requireDefaultLayout()
		let slideLayout: SlideLayoutInternal = {
			_name: defLayout.name,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideNum: this._slides.length + 1,
			_slideObjects: [],
		}

		if (masterTitle) {
			const tmpLayout = this._slideLayouts.find((layout) => layout._name === masterTitle)
			if (tmpLayout) slideLayout = tmpLayout
		}

		const newSlide: PresSlideInternal = new SlideBuilder({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
			getSections: this.getSections,
			presLayout: this.presLayout,
			setSlideNum: this.setSlideNumber,
			slideId: this._slides.length + 256,
			slideRId: this._slides.length + 2,
			slideNumber: this._slides.length + 1,
			slideLayout,
		})

		// A: Add slide to pres
		this._slides.push(newSlide)

		// B: Sections
		// B-1: Add slide to section (if any provided)
		// B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
		if (options?.sectionTitle) {
			const sect = this._sections.find((section) => section.title === options.sectionTitle)
			if (!sect)
				warn('slide/section-not-found', `addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else sect._slides.push(newSlide)
		} else if (this._sections && this._sections.length > 0 && !options?.sectionTitle) {
			const lastSect = this._sections[this._sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect?._type === 'default') lastSect._slides.push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else {
				this._sections.push({
					title: `Default-${this._sections.filter((sect) => sect._type === 'default').length + 1}`,
					_type: 'default',
					_id: `{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}`,
					_slides: [newSlide],
				})
			}
		}

		return newSlide
	}

	/**
	 * Create a custom Slide Layout in any size
	 * @param {PresLayout} layout - layout properties
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: PresLayout): void {
		// @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
		//
		// Every arm below is advice about a value the conversion can still recover -- a numeric
		// string, a missing name -- so they warn and the definition lands. The one thing that
		// cannot be recovered is a non-object, which used to warn and then throw a raw
		// `TypeError` on the next line's `layout.name`, breaking the contract that every failure
		// this library raises is a `TsPptxError`.
		if (!layout || typeof layout !== 'object')
			throw new InvalidOptionError(
				'layout/invalid-definition',
				'defineLayout requires an object `{ name, width, height }` with the dimensions in inches'
			)
		if (!layout.name) warn('layout/invalid-definition', 'defineLayout requires `name`')
		else if (!layout.width) warn('layout/invalid-definition', 'defineLayout requires `width`')
		else if (!layout.height) warn('layout/invalid-definition', 'defineLayout requires `height`')
		else if (typeof layout.height !== 'number')
			warn('layout/invalid-definition', 'defineLayout `height` should be a number (inches)')
		else if (typeof layout.width !== 'number')
			warn('layout/invalid-definition', 'defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = {
			name: layout.name,
			width: inchesToEmu(layoutSideInches(layout.width, 'width')),
			height: inchesToEmu(layoutSideInches(layout.height, 'height')),
		}
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props - layout properties
	 */
	defineSlideMaster(props: SlideMasterProps): void {
		// deep clone the props object to avoid mutating the original object.
		// structuredClone preserves the `SlideMasterProps` type (unlike JSON round-tripping, which widens to `any`).
		const propsClone = structuredClone(props)
		if (!propsClone.title)
			throw new InvalidOptionError(
				'master/missing-title',
				'defineSlideMaster() object argument requires a `title` value.'
			)

		const newLayout: SlideLayoutInternal = {
			_margin: propsClone.margin || DEF_SLIDE_MARGIN_IN,
			// Kept RAW (unescaped) here, unlike `objectName`'s single-escape-upstream design: `_name`
			// doubles as the lookup key `addSlide({masterTitle})` matches against the caller's raw
			// `title` string (see the `layout._name === masterTitle` comparisons in this file).
			// Escaping it here would break that match for any title containing `&`/`<`/`"`. It's
			// escaped once at emission instead -- see `slideObjectToXml` in gen/slide/object.ts.
			_name: propsClone.title,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slide: null,
			_slideNum: 1000 + this._slideLayouts.length + 1,
			_slideNumberProps: propsClone.slideNumber || null,
			_slideObjects: [],
			background: propsClone.background,
		}

		// STEP 1: Create the Slide Master/Layout
		createSlideMaster(propsClone, newLayout)

		// STEP 1b: Master text styles (<p:txStyles>) live on the single shared slide master, not per-layout.
		// Merge each provided group (title/body/other) onto the master, last-call-wins (deck-wide).
		if (propsClone.textStyles && typeof propsClone.textStyles === 'object') {
			this._masterSlide._txStyles = { ...this._masterSlide._txStyles, ...propsClone.textStyles }
		}

		// STEP 2: Add it to layout defs
		this._slideLayouts.push(newLayout)

		// STEP 3: Add background (image data/path must be captured before the package is built)
		if (propsClone.background) addBackgroundDefinition(propsClone.background, newLayout)

		// STEP 4: Add slideNumber to master slide (if any)
		if (newLayout._slideNumberProps && !this._masterSlide._slideNumberProps)
			this._masterSlide._slideNumberProps = newLayout._slideNumberProps
	}
}
