/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */

/**
 * Units of Measure used in PowerPoint documents
 *
 * PowerPoint units are in `DXA` (except for font sizing)
 * - 1 inch is 1440 DXA
 * - 1 inch is 72 points
 * -  1 DXA is 1/20th's of a point
 * - 20 DXA is 1 point
 *
 * Another form of measurement using is an `EMU`
 * - 914400 EMUs is 1 inch
 * -  12700 EMUs is 1 point
 *
 * @see https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 */

/**
 * Object Layouts
 *
 * - 16x9 (10" x 5.625")
 * - 16x10 (10" x 6.25")
 * - 4x3 (10" x 7.5")
 * - Wide (13.33" x 7.5")
 * - [custom] (any size)
 *
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)
 */

import { ZipWriter } from './zip.js'
import Slide from './slide.js'
import {
	AlignH,
	AlignV,
	CHART_TYPE,
	ChartType,
	DEF_PRES_LAYOUT,
	DEF_PRES_LAYOUT_NAME,
	DEF_SLIDE_MARGIN_IN,
	OutputType,
	SCHEME_COLOR_NAMES,
	SHAPE_TYPE,
	SchemeColor,
	ShapeType,
	WRITE_OUTPUT_TYPE,
} from './core-enums.js'
import {
	AddSlideProps,
	CustomPropertyValue,
	IPresentationProps,
	MeasureTextOptions,
	OverflowBoxOptions,
	PresLayout,
	PresSlide,
	PresSlideInternal,
	SectionProps,
	SectionInternalProps,
	SlideLayout,
	SlideLayoutInternal,
	SlideMasterProps,
	SlideNumberProps,
	TableStyleInternal,
	TableStyleProps,
	TableToSlidesProps,
	TextMeasurement,
	TextProps,
	ThemeProps,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './core-interfaces.js'
import * as genCharts from './gen-charts.js'
import * as genObj from './gen-objects.js'
import * as genMedia from './gen-media.js'
import * as genTable from './gen-tables.js'
import * as genXml from './gen-xml.js'
import type { RuntimeAdapter } from './runtime/types.js'
import { FontMetricsRegistry, parseFontMetrics } from './font-metrics.js'
import { applyMeasuredFit, measureText } from './measure-fit.js'
import { getUuid, decodeBase64ToBytes } from './gen-utils.js'
import { inchesToEmu, STANDARD_LAYOUTS, type StandardLayout } from './units.js'

export type { PresSlide as Slide } from './core-interfaces.js'
export type {
	AddSlideProps,
	CustomPropertyValue,
	BackgroundProps,
	BorderProps,
	Color,
	DataOrPathProps,
	DataOrPathRequiredProps,
	HAlign,
	HexColor,
	IChartAreaProps,
	IChartMulti,
	IChartOpts,
	IChartPropsAxisCat,
	IChartPropsAxisSer,
	IChartPropsAxisVal,
	IChartPropsBase,
	IChartPropsChartBar,
	IChartPropsChartDoughnut,
	IChartPropsChartLine,
	IChartPropsChartPie,
	IChartPropsChartRadar,
	IChartPropsDataLabel,
	IChartPropsDataTable,
	IChartPropsFillLine,
	IChartPropsLegend,
	IChartPropsTitle,
	ImageProps,
	Margin,
	MeasureTextOptions,
	MediaProps,
	MediaType,
	ObjectNameProps,
	OverflowBoxOptions,
	ObjectOptions,
	OptsChartData,
	OptsChartGridLine,
	PlaceholderProps,
	PositionProps,
	PresLayout,
	PresSlide,
	PresentationProps,
	SectionProps,
	ShadowProps,
	ShapeFillProps,
	ShapeLineProps,
	ShapeProps,
	SlideMasterChartProps,
	SlideMasterObject,
	SlideMasterProps,
	SlideNumberProps,
	TableCell,
	TableCellProps,
	TableProps,
	TableRow,
	TableRowSlide,
	TableToSlidesProps,
	TextBaseProps,
	TextGlowProps,
	TextMeasurement,
	TextProps,
	TextPropsOptions,
	ThemeColor,
	ThemeProps,
	VAlign,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './core-interfaces.js'
export type {
	CHART_NAME,
	CHART_TYPE,
	JSZIP_OUTPUT_TYPE,
	PLACEHOLDER_TYPE,
	PLACEHOLDER_TYPES,
	SCHEME_COLOR_NAMES,
	SCHEME_COLORS,
	SHAPE_NAME,
	SHAPE_TYPE,
	WRITE_OUTPUT_TYPE,
} from './core-enums.js'

const VERSION = '7.0.0'

/**
 * Media extensions whose bytes are already entropy-coded, so running the ZIP's
 * DEFLATE pass over them costs CPU for a negligible size gain. For these we set
 * the JSZip per-entry compression to STORE while leaving XML parts on DEFLATE.
 * In image/video-heavy decks media dominates the byte count, so this is the
 * dominant cost in `generateAsync` for large presentations (see #1006).
 * Formats that genuinely benefit from DEFLATE (bmp, wav, tiff, emf, wmf, svg)
 * are deliberately excluded so they keep inheriting the global compression.
 */
const ALREADY_COMPRESSED_MEDIA_EXTN = new Set([
	'jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif', 'avif',
	'mp4', 'm4v', 'mov', 'avi', 'mpg', 'mpeg', 'wmv', 'webm', 'mkv',
	'mp3', 'm4a', 'aac', 'ogg', 'oga',
])

function standardLayoutToPresLayout(layout: StandardLayout): PresLayout {
	return {
		name: layout.name,
		width: layout.widthEmu,
		height: layout.heightEmu,
	}
}

export default class PptxGenJS {
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
	private _layout: string
	public set layout(value: string | StandardLayout) {
		// Accept either a layout key string or a STANDARD_LAYOUTS preset object directly.
		const layoutKey = typeof value === 'string' ? value : value?.layout
		const newLayout: PresLayout | undefined = layoutKey ? this.LAYOUTS[layoutKey] : undefined

		if (newLayout) {
			this._layout = layoutKey
			this._presLayout = newLayout
		} else {
			throw new Error('UNKNOWN-LAYOUT')
		}
	}

	public get layout(): string {
		return this._layout
	}

	/**
	 * PptxGenJS Library Version
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
	private _theme: ThemeProps
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
	public get masterSlide(): PresSlide {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private readonly _slides: PresSlideInternal[]
	public get slides(): PresSlide[] {
		return this._slides
	}

	/** this Presentation's sections */
	private readonly _sections: SectionInternalProps[]
	public get sections(): SectionProps[] {
		return this._sections
	}

	/** custom document properties stored in docProps/custom.xml */
	private _customProperties: Array<{ name: string; value: CustomPropertyValue }>
	private readonly _tableStyles: TableStyleInternal[]

	/** slide layout definition objects, used for generating slide layout files */
	private readonly _slideLayouts: SlideLayoutInternal[]
	public get slideLayouts(): SlideLayout[] {
		return this._slideLayouts
	}

	private get internalPresentation(): IPresentationProps {
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
		}
	}

	private LAYOUTS: { [key: string]: PresLayout }

	// Exposed class props
	private readonly _alignH = AlignH
	public get AlignH(): typeof AlignH {
		return this._alignH
	}

	private readonly _alignV = AlignV
	public get AlignV(): typeof AlignV {
		return this._alignV
	}

	private readonly _chartType = ChartType
	public get ChartType(): typeof ChartType {
		return this._chartType
	}

	private readonly _outputType = OutputType
	public get OutputType(): typeof OutputType {
		return this._outputType
	}

	private _presLayout: PresLayout
	public get presLayout(): PresLayout {
		return this._presLayout
	}

	private readonly _schemeColor = SchemeColor
	public get SchemeColor(): typeof SchemeColor {
		return this._schemeColor
	}

	private readonly _shapeType = ShapeType
	public get ShapeType(): typeof ShapeType {
		return this._shapeType
	}

	/**
	 * @depricated use `ChartType`
	 */
	private readonly _charts = CHART_TYPE
	public get charts(): typeof CHART_TYPE {
		return this._charts
	}

	/**
	 * @depricated use `SchemeColor`
	 */
	private readonly _colors = SCHEME_COLOR_NAMES
	public get colors(): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}

	/**
	 * @depricated use `ShapeType`
	 */
	private readonly _shapes = SHAPE_TYPE
	public get shapes(): typeof SHAPE_TYPE {
		return this._shapes
	}

	private readonly _runtime: RuntimeAdapter

	/** Write-side font metrics for measured text fit (`fit:'shrink'`). @see registerFontMetrics */
	private readonly _fontMetrics = new FontMetricsRegistry()

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
		this._author = 'PptxGenJS'
		this._company = 'PptxGenJS'
		this._revision = '1' // Note: Must be a whole number
		this._subject = 'PptxGenJS Presentation'
		this._title = 'PptxGenJS Presentation'
		// PptxGenJS props
		this._presLayout = {
			name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			_sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
			width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
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
		this._tableStyles = []
		this._masterSlide = {
			addChart: null,
			addConnector: null,
			addImage: null,
			addMedia: null,
			addNotes: null,
			addShape: null,
			addTable: null,
			addText: null,
			//
			_name: null,
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
		}
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {AddSlideProps} options - slide masterName and/or sectionTitle
	 * @return {PresSlide} new Slide
	 */
	private readonly addNewSlide = (options?: AddSlideProps): PresSlideInternal => {
		const nextOptions = options || {}
		// Preserve the originating slide's section for all auto-paged continuation slides.
		// Search for the section that owns the current last slide rather than assuming it is
		// the last section — the originating slide may not be at the tail of the deck.
		const lastSlide = this._slides[this._slides.length - 1]
		const sourceSection = this._sections.find(sect => sect._slides.some(s => s._slideNum === lastSlide._slideNum))
		nextOptions.sectionTitle = sourceSection?.title ?? null

		return this.addSlide(nextOptions) as PresSlideInternal
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {PresSlide} Slide
	 * @since 3.0.0
	 */
	private readonly getSlide = (slideNum: number): PresSlideInternal => this._slides.find(slide => slide._slideNum === slideNum)

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {SlideNumberProps} slideNum - slide number config
	 */
	private readonly setSlideNumber = (slideNum: SlideNumberProps): void => {
		// 1: Add slideNumber to slideMaster1.xml
		this._masterSlide._slideNumberProps = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this._slideLayouts.find(layout => layout._name === DEF_PRES_LAYOUT_NAME)._slideNumberProps = slideNum
	}

	/**
	 * Create all chart and media rels for this Presentation
	 * @param {PresSlideInternal | SlideLayoutInternal} slide - slide with rels
	 * @param {ZipWriter} zip - zip writer
	 * @param {Promise<string>[]} chartPromises - promise array
	 */
	private readonly createChartMediaRels = (slide: PresSlideInternal | SlideLayoutInternal, zip: ZipWriter, chartPromises: Promise<string>[]): void => {
		slide._relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide._relsMedia.forEach(rel => {
			if (rel.type !== 'online' && rel.type !== 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (!data.includes(',') && !data.includes(';')) data = 'image/png;base64,' + data
				else if (!data.includes(',')) data = 'image/png;base64,' + data
				else if (!data.includes(';')) data = 'image/png;' + data

				// C: Add media. fflate needs decoded bytes (no base64 convenience), so
				// decode the payload here. Already-compressed formats (JPEG/PNG/video/…)
				// gain ~nothing from DEFLATE, so STORE them to avoid wasted compression
				// CPU on large decks (#1006); other parts inherit global compression.
				const bytes = decodeBase64ToBytes(data)
				if (!bytes) return
				const extn = (rel.extn || rel.Target.split('.').pop() || '').toLowerCase()
				zip.add(rel.Target.replace('..', 'ppt'), bytes, { store: ALREADY_COMPRESSED_MEDIA_EXTN.has(extn) })
			}
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	private readonly exportPresentation = async (props: WriteProps): Promise<string | ArrayBuffer | Blob | Uint8Array> => {
		const arrChartPromises: Promise<string>[] = []
		let arrMediaPromises: Promise<string>[] = []
		const zip = new ZipWriter()

		// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
		const onMediaError = props.onMediaError ?? 'throw'
		this._slides.forEach(slide => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(slide, this._runtime, onMediaError))
		})
		this._slideLayouts.forEach(layout => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(layout, this._runtime, onMediaError))
		})
		arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(this._masterSlide, this._runtime, onMediaError))

		// STEP 2: Wait for Promises (if any) then generate the PPTX file
		return await Promise.all(arrMediaPromises).then(async () => {
			// PERF: Collapse identical media to a single package part across the entire deck.
			// Each target (slide/layout/master) namespaces its media `Target` by slide, so the
			// same image used on multiple slides — or loaded from the same path — otherwise
			// embeds one copy per use. By now `encodeSlideMediaRels` has populated every
			// `rel.data`, so we can point later duplicates at the first occurrence's `Target`
			// (slide `.rels` reference media by rId, and sharing a part across slides is valid
			// OOXML). This subsumes the per-slide path/data de-dup for cross-slide reuse and
			// also covers background images (issue #1339).
			const canonicalMediaTargets = new Map<string, string>()
			for (const target of [...this._slides, ...this._slideLayouts, this._masterSlide]) {
				for (const rel of target._relsMedia || []) {
					if (rel.type === 'online' || rel.type === 'hyperlink' || typeof rel.data !== 'string' || !rel.data) continue
					// Key on extension + bytes so identical content with differing part
					// extensions is never merged into one mistyped file.
					const key = (rel.extn || '') + '\0' + rel.data
					const canonical = canonicalMediaTargets.get(key)
					if (canonical) rel.Target = canonical
					else canonicalMediaTargets.set(key, rel.Target)
				}
			}

			// A: Add empty placeholder objects to slides that don't already have them
			this._slides.forEach(slide => {
				if (slide._slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
			})

			// A.1: Measured text fit — bake a real fontScale onto `fit:'shrink'` text
			// boxes when font metrics are registered, before the sync XML pass reads it.
			applyMeasuredFit(this._slides, this._fontMetrics)

			// B: Add all required files. fflate keys on full slash-paths and emits no
			// directory entries, so there is no folder scaffolding to set up (and no
			// stray empty-directory entries to guard against on minimal decks).
			const hasCustomProps = this._customProperties.length > 0
			zip.add('[Content_Types].xml', genXml.makeXmlContTypes(this._slides, this._slideLayouts, this._masterSlide, hasCustomProps)) // TODO: pass only `this` like below! 20200206
			zip.add('_rels/.rels', genXml.makeXmlRootRels(hasCustomProps))
			zip.add('docProps/app.xml', genXml.makeXmlApp(this._slides, this.company)) // TODO: pass only `this` like below! 20200206
			zip.add('docProps/core.xml', genXml.makeXmlCore(this.title, this.subject, this.author, this.revision)) // TODO: pass only `this` like below! 20200206
			if (hasCustomProps) {
				zip.add('docProps/custom.xml', genXml.makeXmlCustomProperties(this._customProperties))
			}
			zip.add('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this._slides))
			zip.add('ppt/theme/theme1.xml', genXml.makeXmlTheme(this.internalPresentation))
			// emit a separate theme2.xml part so notesMaster1.xml.rels resolves
			zip.add('ppt/theme/theme2.xml', genXml.makeXmlTheme(this.internalPresentation))
			zip.add('ppt/presentation.xml', genXml.makeXmlPresentation(this.internalPresentation))
			zip.add('ppt/presProps.xml', genXml.makeXmlPresProps())
			zip.add('ppt/tableStyles.xml', genXml.makeXmlTableStyles(this._tableStyles))
			zip.add('ppt/viewProps.xml', genXml.makeXmlViewProps())

			// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
			this._slideLayouts.forEach((layout, idx) => {
				zip.add(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, genXml.makeXmlLayout(layout))
				zip.add(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, genXml.makeXmlSlideLayoutRel(idx + 1, this._slideLayouts))
			})
			this._slides.forEach((slide, idx) => {
				zip.add(`ppt/slides/slide${idx + 1}.xml`, genXml.makeXmlSlide(slide))
				zip.add(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, genXml.makeXmlSlideRel(this._slides, this._slideLayouts, idx + 1))
				// Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
				zip.add(`ppt/notesSlides/notesSlide${idx + 1}.xml`, genXml.makeXmlNotesSlide(slide))
				zip.add(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, genXml.makeXmlNotesSlideRel(slide, idx + 1))
			})
			zip.add('ppt/slideMasters/slideMaster1.xml', genXml.makeXmlMaster(this._masterSlide, this._slideLayouts))
			zip.add('ppt/slideMasters/_rels/slideMaster1.xml.rels', genXml.makeXmlMasterRel(this._masterSlide, this._slideLayouts))
			zip.add('ppt/notesMasters/notesMaster1.xml', genXml.makeXmlNotesMaster())
			zip.add('ppt/notesMasters/_rels/notesMaster1.xml.rels', genXml.makeXmlNotesMasterRel())

			// D: Create all Rels (images, media, chart data)
			this._slideLayouts.forEach(layout => {
				this.createChartMediaRels(layout, zip, arrChartPromises)
			})
			this._slides.forEach(slide => {
				this.createChartMediaRels(slide, zip, arrChartPromises)
			})
			this.createChartMediaRels(this._masterSlide, zip, arrChartPromises)

			// E: Wait for Promises (if any) then generate the PPTX file
			return await Promise.all(arrChartPromises).then(async () => {
				const compression = props.compression !== false
				if (props.outputType === 'STREAM') {
					// A: stream file
					return await zip.generate('nodebuffer', { compression })
				} else if (props.outputType) {
					// B: Node [fs]: Output type user option or default
					return await zip.generate(props.outputType, { compression })
				} else {
					// C: Browser: Output blob as app/ms-pptx
					return await zip.generate('blob', { compression })
				}
			})
		})
	}

	// FONT METRICS (measured text fit)

	/**
	 * Register a font's metrics so `fit:'shrink'` text boxes are measured and a real
	 * `fontScale` is baked at export time (text renders pre-shrunk in headless
	 * renderers and on plain file-open, with no manual edit/resize).
	 *
	 * Without registered metrics, `fit:'shrink'` keeps its current behavior (a bare
	 * `<a:normAutofit/>` that only PowerPoint recomputes on edit). Register the same
	 * face once per weight/style you use; bold/italic advances differ.
	 * @param {string} face - font family name as used in `fontFace` (e.g. 'Aptos')
	 * @param {string | Uint8Array | ArrayBuffer} source - font file path/URL (Node/web) or raw TTF/OTF bytes
	 * @param {object} [opts] - variant flags; advances differ per weight/style
	 * @example await pptx.registerFontMetrics('Aptos', '/usr/share/fonts/Aptos.ttf')
	 * @example await pptx.registerFontMetrics('Aptos', aptosBoldBytes, { bold: true })
	 * @since v6.1.0
	 */
	async registerFontMetrics(face: string, source: string | Uint8Array | ArrayBuffer, opts?: { bold?: boolean; italic?: boolean }): Promise<void> {
		let bytes: Uint8Array
		if (typeof source === 'string') bytes = await this._runtime.loadFontData(source)
		else if (source instanceof Uint8Array) bytes = source
		else bytes = new Uint8Array(source)
		const metrics = await parseFontMetrics(bytes)
		this._fontMetrics.set(face, metrics, opts)
	}

	/**
	 * Measure how tall text wraps at a given width, using the **same** calibrated wrap
	 * model the export-time autofit bake uses — so a layout-time prediction matches the
	 * value `fit:'shrink'`/`'resize'` would bake. Synchronous: register the face's metrics
	 * first with {@link PptxGenJS.registerFontMetrics} (lookup is sync).
	 *
	 * Lets a consumer size its own geometry before export — grow a card to fit its text,
	 * reflow a grid, or detect overflow at layout time. Heights err **tall** (conservative),
	 * so the returned height is ≥ what PowerPoint/LibreOffice render.
	 * @param {string | TextProps[]} text - a string or run array (per-run options override the defaults in `opts`)
	 * @param {MeasureTextOptions} opts - width (inches), font size/face (points), and spacing
	 * @returns {TextMeasurement} laid-out height + line count + fit helpers
	 * @example const m = pptx.measureText('Long heading…', { wIn: 3, fontSize: 18, fontFace: 'Aptos' })
	 * @example if (pptx.measureText(runs, { wIn, fontSize, fontFace }).heightIn > cardHeightIn) growCard()
	 * @since v6.1.0
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
	 * @since v6.1.0
	 */
	overflowsBox(text: string | TextProps[], opts: OverflowBoxOptions): boolean {
		const m = measureText(this._fontMetrics, text, opts)
		return m.measurable && !m.fitsBox(opts.hIn)
	}

	// EXPORT METHODS

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props - output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file stream
	 */
	async stream(props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		return await this.exportPresentation({
			compression: props?.compression,
			outputType: 'STREAM',
		})
	}

	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	async write(props?: WriteProps | WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		// DEPRECATED: @deprecated v3.5.0 - outputType - [[remove in v4.0.0]]
		const propsOutpType = typeof props === 'object' && props?.outputType ? props.outputType : props ? (props as WRITE_OUTPUT_TYPE) : null
		const propsCompress = typeof props === 'object' ? props?.compression : undefined
		const propsMediaError = typeof props === 'object' ? props?.onMediaError : undefined

		return await this.exportPresentation({
			compression: propsCompress,
			outputType: propsOutpType,
			onMediaError: propsMediaError,
		})
	}

	/**
	 * Export the current Presentation.
	 * Write the generated presentation to disk (Node) or trigger a download (browser).
	 * @param {WriteFileProps} props - output file properties
	 * @returns {Promise<string>} the presentation name
	 */
	async writeFile(props?: WriteFileProps | string): Promise<string> {
		if (typeof props === 'string') {
			// DEPRECATED: @deprecated v3.5.0 - fileName - [[remove in v4.0.0]]
			console.warn('[WARNING] writeFile(string) is deprecated - pass { fileName } instead.')
			props = { fileName: props }
		}
		const { fileName: rawName = 'Presentation.pptx', compression, onMediaError } = props as WriteFileProps
		const fileName = rawName.toLowerCase().endsWith('.pptx') ? rawName : `${rawName}.pptx`

		const data = await this.exportPresentation({ compression, outputType: this._runtime.writeFileOutputType, onMediaError })
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
		this._customProperties = this._customProperties.filter(p => p.name !== name)
		this._customProperties.push({ name, value })
	}

	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: SectionProps): void {
		if (!section) {
			console.warn('addSection requires an argument')
			return
		} else if (!section.title) {
			console.warn('addSection requires a title')
			return
		} else if (this._sections.some(sect => sect.title === section.title)) {
			console.warn(`addSection: a section titled "${section.title}" already exists; ignoring duplicate`)
			return
		}

		const newSection: SectionInternalProps = {
			_type: 'user',
			_slides: [],
			title: section.title,
		}

		if (section.order) this._sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} options - slide options
	 * @returns {PresSlide} the new Slide
	 */
	addSlide(options?: AddSlideProps): PresSlide {
		// TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
		const masterSlideName = typeof options === 'string' ? options : options?.masterName ? options.masterName : ''
		let slideLayout: SlideLayoutInternal = {
			_name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideNum: this._slides.length + 1,
			_slideObjects: [],
		}

		if (masterSlideName) {
			const tmpLayout = this._slideLayouts.find(layout => layout._name === masterSlideName)
			if (tmpLayout) slideLayout = tmpLayout
		}

		const newSlide: PresSlideInternal = new Slide({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
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
			const sect = this._sections.find(section => section.title === options.sectionTitle)
			if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else sect._slides.push(newSlide)
		} else if (this._sections && this._sections.length > 0 && (!options?.sectionTitle)) {
			const lastSect = this._sections[this._sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect._type === 'default') lastSect._slides.push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else {
				this._sections.push({
					title: `Default-${this._sections.filter(sect => sect._type === 'default').length + 1}`,
					_type: 'default',
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
		if (!layout) console.warn('defineLayout requires `{name, width, height}`')
		else if (!layout.name) console.warn('defineLayout requires `name`')
		else if (!layout.width) console.warn('defineLayout requires `width`')
		else if (!layout.height) console.warn('defineLayout requires `height`')
		else if (typeof layout.height !== 'number') console.warn('defineLayout `height` should be a number (inches)')
		else if (typeof layout.width !== 'number') console.warn('defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = {
			name: layout.name,
			_sizeW: inchesToEmu(Number(layout.width)),
			_sizeH: inchesToEmu(Number(layout.height)),
			width: inchesToEmu(Number(layout.width)),
			height: inchesToEmu(Number(layout.height)),
		}
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props - layout properties
	 */
	defineSlideMaster(props: SlideMasterProps): void {
		// (ISSUE#406;PULL#1176) deep clone the props object to avoid mutating the original object
		const propsClone = JSON.parse(JSON.stringify(props))
		if (!propsClone.title) throw new Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		const newLayout: SlideLayoutInternal = {
			_margin: propsClone.margin || DEF_SLIDE_MARGIN_IN,
			_name: propsClone.title,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slide: null,
			_slideNum: 1000 + this._slideLayouts.length + 1,
			_slideNumberProps: propsClone.slideNumber || null,
			_slideObjects: [],
			background: propsClone.background || null,
			bkgd: propsClone.bkgd || null,
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideMaster(propsClone, newLayout)

		// STEP 2: Add it to layout defs
		this._slideLayouts.push(newLayout)

		// STEP 3: Add background (image data/path must be captured before `exportPresentation()` is called)
		if (propsClone.background || propsClone.bkgd) genObj.addBackgroundDefinition(propsClone.background, newLayout)

		// STEP 4: Add slideNumber to master slide (if any)
		if (newLayout._slideNumberProps && !this._masterSlide._slideNumberProps) this._masterSlide._slideNumberProps = newLayout._slideNumberProps
	}

	/**
	 * Register a reusable custom table style and return its GUID.
	 * The style is written to `ppt/tableStyles.xml` and is editable in PowerPoint's
	 * Table Styles gallery. Pass the returned GUID as `TableProps.tableStyle`, and use
	 * the `has*` flags (`hasHeader`, `hasBandedRows`, …) to activate the matching regions.
	 * @param {TableStyleProps} props - custom table style definition (requires `name`)
	 * @returns {string} braced GUID to use as `tableStyle`
	 * @example
	 * const brand = pptx.defineTableStyle({
	 *   name: 'Brand Banded',
	 *   firstRow: { fill:'1A2B3C', color:'FFFFFF', bold:true },
	 *   band1H:   { fill:'EAF1F8' },
	 * })
	 * slide.addTable(rows, { tableStyle: brand, hasHeader:true, hasBandedRows:true })
	 */
	defineTableStyle(props: TableStyleProps): string {
		if (!props || typeof props !== 'object') throw new Error('defineTableStyle() requires a `{ name, ... }` object argument')
		if (!props.name || typeof props.name !== 'string') throw new Error('defineTableStyle() requires a non-empty `name`')

		const guid = `{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}`
		this._tableStyles.push({ guid, def: props })
		return guid
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		// @note `verbose` option is undocumented; used for verbose output of layout process
		genTable.genTableToSlides(
			this,
			eleId,
			options,
			options?.masterSlideName ? this._slideLayouts.find(layout => layout._name === options.masterSlideName) : null
		)
	}
}
