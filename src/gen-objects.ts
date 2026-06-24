/**
 * PptxGenJS: Slide Object Generators
 */

import {
	BARCHART_COLORS,
	CHART_NAME,
	CHART_TYPE,
	connectorPresetFor,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_IN,
	DEF_CHART_BORDER,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_LINE_COLOR,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	IMG_PLAYBTN,
	PIECHART_COLORS,
	SCHEME_COLOR_NAMES,
	SHAPE_NAME,
	SHAPE_TYPE,
	SLIDE_OBJECT_TYPES,
	TEXT_HALIGN,
	TEXT_VALIGN,
	VALID_SHAPE_PRESETS,
} from './core-enums.js'
import type { PLACEHOLDER_TYPE } from './core-enums.js'
import type {
	AddSlideProps,
	BackgroundProps,
	BorderProps,
	CommentProps,
	ConnectorProps,
	Coord,
	GroupChildProps,
	GroupProps,
	IChartMulti,
	IChartOpts,
	IChartOptsLib,
	IOptsChartData,
	ISlideObject,
	ImageProps,
	MediaProps,
	NotesProps,
	ObjectOptions,
	OptsChartData,
	OptsChartGridLine,
	PresLayout,
	PresSlideInternal,
	ShapeFillProps,
	ShapeLineProps,
	ShapeProps,
	SlideLayoutInternal,
	SlideMasterObject,
	SlideMasterProps,
	TableCell,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
} from './core-interfaces.js'
import { getSlidesForTableRows } from './gen-tables.js'
import { encodeXmlEntities, getNewRelId, getSmartParseNumber, inch2Emu, valToPts, correctShadowOptions, validateObjectName, svgMarkupToDataUri, getImageSizeFromBase64, imageContentType } from './gen-utils.js'

/** counter for included charts (used for index in their filenames) */
let _chartCounter = 0

/** DPI PowerPoint assumes when sizing an inserted raster image (natural pixels / 96 == inches) */
const IMAGE_NATURAL_DPI = 96

type BorderTuple = [BorderProps, BorderProps, BorderProps, BorderProps]
type HyperlinkTextObject = (TextProps | ISlideObject | TableCell) & {
	options?: TextPropsOptions | ObjectOptions
	text?: string | number | TextProps[] | TableCell[]
}

function normalizeBorderTuple(border: BorderProps | BorderTuple): BorderTuple {
	return Array.isArray(border) ? border : [border, border, border, border]
}

/**
 * Dispatch a key-tagged child-object descriptor (`{ text }`, `{ image }`, `{ shape }`, …) to the
 * matching `add*Definition`. Shared by `createSlideMaster` (slide master `objects`) and
 * `addGroupDefinition` (group children) so the descriptor mapping lives in one place.
 *
 * `placeholder` is intentionally not handled here — it is master-specific and needs the object's
 * index for `_placeholderIdx`, so `createSlideMaster` handles that case itself.
 * @param target - slide (or master) the object is appended to
 * @param object - the child descriptor
 * @returns `true` if the descriptor was recognized and added, else `false`
 */
function addChildDefinition(target: PresSlideInternal, object: SlideMasterObject | GroupChildProps): boolean {
	if ('chart' in object) addChartDefinition(target, object.chart.type, object.chart.data, object.chart.opts || object.chart.options || {})
	else if ('image' in object) addImageDefinition(target, object.image)
	else if ('line' in object) addShapeDefinition(target, SHAPE_TYPE.LINE, object.line)
	else if ('rect' in object) addShapeDefinition(target, SHAPE_TYPE.RECTANGLE, object.rect)
	else if ('roundRect' in object) addShapeDefinition(target, SHAPE_TYPE.ROUNDED_RECTANGLE, object.roundRect)
	else if ('shape' in object) addShapeDefinition(target, object.shape.type, object.shape.options || {})
	else if ('text' in object) addTextDefinition(target, Array.isArray(object.text.text) ? object.text.text : [{ text: object.text.text }], object.text.options || {}, false)
	else return false
	return true
}

/** Counter for default group names (`Group N`), incremented across nesting depth within a slide. */
let _groupNameCounter = 0

/**
 * Build a group (`<p:grpSp>`) render-object from its child descriptors, without appending the
 * group itself to the slide. Nested `group` children recurse, so a group can contain other groups.
 *
 * An identity child coordinate space is kept at every depth (emitted in `gen-xml` as
 * `chOff/chExt == off/ext`), so children — including descendants of nested groups — keep their
 * slide-absolute coordinates and grouping is visually a no-op while making the objects one
 * selectable PowerPoint group. Charts, media, tables, and placeholders are not supported as group
 * children yet; each is skipped with a warning. When `opts.x/y/w/h` are omitted the group's bounds
 * are auto-computed (in `gen-xml`) as the bounding box of its children.
 *
 * `target` stays the slide at every depth so leaf descendants register their image/chart rels and
 * unique ids slide-level, even when nested inside child groups.
 * @param target - slide the group's leaf children register rels against
 * @param children - the child-object descriptors
 * @param opts - group position/size/name options
 */
function buildGroupObject(target: PresSlideInternal, children: GroupChildProps[], opts: GroupProps): ISlideObject {
	const groupObjects: ISlideObject[] = []

	;(children || []).forEach(child => {
		// Nested group: recurse and embed the child group object directly (no slide splice — its own
		// leaf descendants still register against `target` inside the recursive call).
		if ('group' in child) {
			groupObjects.push(buildGroupObject(target, child.group.children, child.group.options || {}))
			return
		}
		// Reject object types grouping does not support yet (rels/ID/transform work pending).
		if ('chart' in child || 'placeholder' in child || 'table' in child || 'media' in child) {
			console.warn(`Warning: addGroup() does not support '${Object.keys(child)[0]}' children yet; skipping.`)
			return
		}
		// Reuse the existing add*Definition logic (which registers any image/chart rels on the slide,
		// correctly — grouped children still reference slide-level relationships), then move the
		// just-appended object(s) off the slide's top-level list into this group's child list.
		const before = target._slideObjects.length
		if (!addChildDefinition(target, child)) {
			console.warn(`Warning: addGroup() received an unrecognized child descriptor (${Object.keys(child).join(', ')}); skipping.`)
			return
		}
		groupObjects.push(...target._slideObjects.splice(before))
	})

	const objectName = opts.objectName
		? encodeXmlEntities(validateObjectName(opts.objectName, 'group'))
		: `Group ${++_groupNameCounter}`

	return {
		_type: SLIDE_OBJECT_TYPES.group,
		_groupObjects: groupObjects,
		options: {
			x: opts.x,
			y: opts.y,
			w: opts.w,
			h: opts.h,
			rotate: opts.rotate,
			flipH: opts.flipH,
			flipV: opts.flipV,
			objectName,
			altText: opts.altText,
			objectLock: opts.objectLock,
		},
	}
}

/**
 * Add a group (`<p:grpSp>`) of child objects to a slide. Children may include nested groups.
 * @param target - slide the group is added to
 * @param children - the child-object descriptors
 * @param opts - group position/size/name options
 */
export function addGroupDefinition(target: PresSlideInternal, children: GroupChildProps[], opts: GroupProps): void {
	target._slideObjects.push(buildGroupObject(target, children, opts))
}

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlideInternal|SlideLayoutInternal} target - empty slide object that should be updated by the passed definition
 */
export function createSlideMaster(props: SlideMasterProps, target: SlideLayoutInternal): void {
	// STEP 1: Add background if either the slide or layout has background props
	// if (props.background || target.background) addBackgroundDefinition(props.background, target)
	if (props.bkgd) target.bkgd = props.bkgd // DEPRECATED: (remove in v4.0.0)

	// STEP 2: Add all Slide Master objects in the order they were given
	if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
		props.objects.forEach((object, idx) => {
			const tgt = target as PresSlideInternal
			if (addChildDefinition(tgt, object)) {
				// handled by the shared chart/image/shape/text dispatch
			} else if ('placeholder' in object) {
				// TODO: 20180820: Check for existing `name`?
				const placeholder = object.placeholder
				const { name, type, ...rawPlaceholderOptions } = placeholder.options
				const placeholderOptions = rawPlaceholderOptions as TextPropsOptions & ObjectOptions
				placeholderOptions.placeholder = name
				placeholderOptions._placeholderType = type
				placeholderOptions._placeholderIdx = 100 + idx
				addTextDefinition(tgt, [{ text: placeholder.text }], placeholderOptions, true)
				// TODO: ISSUE#599 - only text is suported now (add more below)
				// else if (placeholder.image) addImageDefinition(tgt, placeholder.image)
				/* 20200120: So... image placeholders go into the "slideLayoutN.xml" file and addImage doesnt do this yet...
					<p:sp>
				  <p:nvSpPr>
					<p:cNvPr id="7" name="Picture Placeholder 6">
					  <a:extLst>
						<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
						  <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{CE1AE45D-8641-0F4F-BDB5-080E69CCB034}"/>
						</a:ext>
					  </a:extLst>
					</p:cNvPr>
					<p:cNvSpPr>
				*/
			}
		})
	}

	// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (props.slideNumber && typeof props.slideNumber === 'object') target._slideNumberProps = props.slideNumber
}

/**
 * Round and clamp an integer chart percentage/angle option into a schema-valid range.
 *
 * Several chart attributes are bounded integer types whose out-of-range values make
 * PowerPoint report the package as needing repair: `<c:overlap>` (ST_Overlap, -100..100),
 * `<c:gapWidth>`/`<c:gapDepth>` (ST_GapAmount, 0..500), `<c:holeSize>` (ST_HoleSize, 10..90)
 * and `<c:firstSliceAng>` (ST_FirstSliceAng, 0..360). Missing/non-numeric input returns
 * `undefined` so the caller can apply its own default; an out-of-range value is clamped
 * and a warning is emitted (per the library's warn-rather-than-degrade policy).
 * @param value - caller-supplied option value
 * @param min - inclusive lower bound
 * @param max - inclusive upper bound
 * @param name - option name, for the warning message
 */
function clampChartPct(value: number | undefined, min: number, max: number, name: string): number | undefined {
	if (typeof value !== 'number' || isNaN(value)) return undefined
	const clamped = Math.min(max, Math.max(min, Math.round(value)))
	if (clamped !== value) console.warn(`Warning: ${name} ${value} is outside the valid range ${min}-${max}; using ${clamped}.`)
	return clamped
}

/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_NAME | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOptsLib} `opt` chart options
 * @param {PresSlideInternal} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *    title: 'eSurvey chart',
 *    data: [
 *        {
 *            name: 'Income',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *        },
 *        {
 *            name: 'Expense',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [18.1, 22.8, 23.9, 25.1, 25]
 *        }
 *    ]
 * }
 */
export function addChartDefinition(target: PresSlideInternal, type: CHART_NAME | IChartMulti[], data: OptsChartData[] | IChartOpts, opt?: IChartOptsLib): object {
	function correctGridLineOptions(glOpts: OptsChartGridLine): void {
		if (!glOpts || glOpts.style === 'none') return
		if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
			console.warn('Warning: chart.gridLine.size must be greater than 0.')
			delete glOpts.size // delete prop to used defaults
		}
		if (glOpts.style && !['solid', 'dash', 'dot'].includes(glOpts.style)) {
			console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.')
			delete glOpts.style
		}
		if (glOpts.cap && !['flat', 'square', 'round'].includes(glOpts.cap)) {
			console.warn('Warning: chart.gridLine.cap options: `flat`, `square`, `round`.')
			delete glOpts.cap
		}
	}

	const chartId = ++_chartCounter
	const resultObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.chart,
	}
	// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
	// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt: IChartOpts | IChartOptsLib | undefined
	let tmpData: OptsChartData[] = []
	if (Array.isArray(type)) {
		// For multi-type charts there needs to be data for each type,
		// as well as a single data source for non-series operations.
		// The data is indexed below to keep the data in order when segmented
		// into types.
		type.forEach(obj => {
			tmpData = tmpData.concat(obj.data)
		})
		tmpOpt = !Array.isArray(data) && data && typeof data === 'object' ? data : opt
	} else {
		tmpData = Array.isArray(data) ? data : []
		tmpOpt = opt
	}
	tmpData.forEach((item, i) => {
		item._dataIndex = i

		// Converts the 'labels' array from string[] to string[][] (or the respective primitive type), if needed
		if (item.labels !== undefined && !Array.isArray(item.labels[0])) {
			item.labels = [item.labels as string[]]
		}
	})
	const options: IChartOptsLib = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {}

	// STEP 1: TODO: check for reqd fields, correct type, etc
	// `type` exists in CHART_TYPE
	// Array.isArray(data)
	/*
		if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
			&& rel.data[0].labels && Array.isArray(rel.data[0].labels)
			&& rel.data[0].values && Array.isArray(rel.data[0].values) ) {
			obj = rel.data[0];
		}
		else {
			console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
			return;
		}
		*/

	// STEP 2: Set default options/decode user options
	// A: Core
	options._type = type
	options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1
	options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'
	options.objectName = options.objectName
		? encodeXmlEntities(validateObjectName(options.objectName, 'chart'))
		: `Chart ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.chart).length}`

	// B: Options: misc
	if (!['bar', 'col'].includes(options.barDir || '')) options.barDir = 'col'

	// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
	// barGrouping must be handled before data label validation as it can affect valid label positioning
	if (options._type === CHART_TYPE.AREA) {
		if (!['stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (options._type === CHART_TYPE.BAR) {
		if (!['clustered', 'stacked', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'clustered'
	}
	if (options._type === CHART_TYPE.BAR3D) {
		if (!['clustered', 'stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (options.barGrouping?.includes('tacked')) {
		if (!options.barGapWidthPct) options.barGapWidthPct = 50
	}
	// Clean up and validate data label positions
	// REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f?redirectedfrom=MSDN
	if (options.dataLabelPosition) {
		const dataLabelPosition = options.dataLabelPosition
		if (options._type === CHART_TYPE.AREA || options._type === CHART_TYPE.BAR3D || options._type === CHART_TYPE.DOUGHNUT || options._type === CHART_TYPE.RADAR) { delete options.dataLabelPosition }
		if (options._type === CHART_TYPE.PIE) {
			if (!['bestFit', 'ctr', 'inEnd', 'outEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (options._type === CHART_TYPE.BUBBLE || options._type === CHART_TYPE.BUBBLE3D || options._type === CHART_TYPE.LINE || options._type === CHART_TYPE.SCATTER) {
			if (!['b', 'ctr', 'l', 'r', 't'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (options._type === CHART_TYPE.BAR) {
			if (!['stacked', 'percentStacked'].includes(options.barGrouping || '')) {
				if (!['ctr', 'inBase', 'inEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
			}
			if (!['clustered'].includes(options.barGrouping || '')) {
				if (!['ctr', 'inBase', 'inEnd', 'outEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
			}
		}
	}
	options.dataLabelBkgrdColors = options.dataLabelBkgrdColors || !options.dataLabelBkgrdColors ? options.dataLabelBkgrdColors : false
	if (!['b', 'l', 'r', 't', 'tr'].includes(options.legendPos || '')) options.legendPos = 'r'

	// 3D bar: ST_Shape
	if (!['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].includes(options.bar3DShape || '')) options.bar3DShape = 'box'
	// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
	// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
	if (!['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].includes(options.lineDataSymbol || '')) options.lineDataSymbol = 'circle'
	if (!['gap', 'span', 'zero'].includes(options.displayBlanksAs || '')) options.displayBlanksAs = 'gap'
	if (!['standard', 'marker', 'filled'].includes(options.radarStyle || '')) options.radarStyle = 'standard'
	// Marker size emits as `<c:size val>` (ST_MarkerSize): an integer in [2,72] points.
	// Out-of-range or non-integer values make PowerPoint report the file as needing
	// repair, so round and clamp into range and warn when the input is coerced.
	{
		const rawSymbolSize = options.lineDataSymbolSize
		const hasSymbolSize = rawSymbolSize != null && !isNaN(rawSymbolSize)
		const symbolSize = Math.min(72, Math.max(2, Math.round(hasSymbolSize ? rawSymbolSize : 6)))
		if (hasSymbolSize && symbolSize !== rawSymbolSize) {
			console.warn(`Warning: lineDataSymbolSize ${rawSymbolSize} is outside the valid marker size range (integer 2-72); using ${symbolSize}.`)
		}
		options.lineDataSymbolSize = symbolSize
	}
	options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? valToPts(options.lineDataSymbolLineSize) : valToPts(0.75)
	// `layout` allows the override of PPT defaults to maximize space
	const chartLayout = options.layout
	if (chartLayout) {
		;(['x', 'y', 'w', 'h'] as const).forEach(key => {
			const val = chartLayout[key]
			const numVal = Number(val)
			if (isNaN(numVal) || numVal < 0 || numVal > 1) {
				console.warn('Warning: chart.layout.' + key + ' can only be 0-1')
				delete chartLayout[key] // remove invalid value so that default will be used
			}
		})
	}

	// Set gridline defaults
	options.catGridLine = options.catGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	options.valGridLine = options.valGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : {})
	options.serGridLine = options.serGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	correctGridLineOptions(options.catGridLine)
	correctGridLineOptions(options.valGridLine)
	correctGridLineOptions(options.serGridLine)
	correctShadowOptions(options.shadow)

	// C: Options: plotArea
	options.showDataTable = options.showDataTable || !options.showDataTable ? options.showDataTable : false
	options.showDataTableHorzBorder = options.showDataTableHorzBorder || !options.showDataTableHorzBorder ? options.showDataTableHorzBorder : true
	options.showDataTableVertBorder = options.showDataTableVertBorder || !options.showDataTableVertBorder ? options.showDataTableVertBorder : true
	options.showDataTableOutline = options.showDataTableOutline || !options.showDataTableOutline ? options.showDataTableOutline : true
	options.showDataTableKeys = options.showDataTableKeys || !options.showDataTableKeys ? options.showDataTableKeys : true
	options.showLabel = options.showLabel || !options.showLabel ? options.showLabel : false
	options.showLegend = options.showLegend || !options.showLegend ? options.showLegend : false
	options.showPercent = options.showPercent || !options.showPercent ? options.showPercent : true
	options.showTitle = options.showTitle || !options.showTitle ? options.showTitle : false
	options.showValue = options.showValue || !options.showValue ? options.showValue : false
	options.showLeaderLines = options.showLeaderLines || !options.showLeaderLines ? options.showLeaderLines : false
	options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true
	options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true
	options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true

	options.v3DRotX = typeof options.v3DRotX === 'number' && !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30
	options.v3DRotY = typeof options.v3DRotY === 'number' && !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30
	options.v3DRAngAx = options.v3DRAngAx || !options.v3DRAngAx ? options.v3DRAngAx : true
	options.v3DPerspective = typeof options.v3DPerspective === 'number' && !isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30

	// D: Options: chart
	// `<c:gapWidth>`/`<c:gapDepth>` are ST_GapAmount (integer 0..500); `<c:overlap>` is
	// ST_Overlap (integer -100..100). Out-of-range values trigger PowerPoint repair.
	options.barGapWidthPct = clampChartPct(options.barGapWidthPct, 0, 500, 'barGapWidthPct') ?? 150
	options.barGapDepthPct = clampChartPct(options.barGapDepthPct, 0, 500, 'barGapDepthPct') ?? 150
	options.barOverlapPct = clampChartPct(options.barOverlapPct, -100, 100, 'barOverlapPct')
	// `<c:holeSize>` is ST_HoleSize (10..90); `<c:firstSliceAng>` is ST_FirstSliceAng (0..360).
	options.holeSize = clampChartPct(options.holeSize, 10, 90, 'holeSize')
	options.firstSliceAng = clampChartPct(options.firstSliceAng, 0, 360, 'firstSliceAng')

	options.chartColors = Array.isArray(options.chartColors)
		? options.chartColors
		: options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT
			? PIECHART_COLORS
			: BARCHART_COLORS
	options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : undefined
	// DEPRECATED: v3.11.0 - use `plotArea.border` vvv
	options.border = options.border && typeof options.border === 'object' ? options.border : undefined
	if (options.border && (!options.border.pt || isNaN(options.border.pt))) options.border.pt = DEF_CHART_BORDER.pt
	if (options.border && (!options.border.color || typeof options.border.color !== 'string')) options.border.color = DEF_CHART_BORDER.color
	// DEPRECATED: (remove above in v4.0) ^^^
	options.plotArea = options.plotArea || {}
	options.plotArea.border = options.plotArea.border && typeof options.plotArea.border === 'object' ? options.plotArea.border : undefined
	if (options.plotArea.border && (!options.plotArea.border.pt || isNaN(options.plotArea.border.pt))) options.plotArea.border.pt = DEF_CHART_BORDER.pt
	if (options.plotArea.border && (!options.plotArea.border.color || typeof options.plotArea.border.color !== 'string')) { options.plotArea.border.color = DEF_CHART_BORDER.color }
	if (options.border) options.plotArea.border = options.border // @deprecated [[remove in v4.0]]
	options.plotArea.fill = options.plotArea.fill || {}
	if (options.fill) options.plotArea.fill.color = options.fill // @deprecated [[remove in v4.0]]
	//
	options.chartArea = options.chartArea || {}
	options.chartArea.border = options.chartArea.border && typeof options.chartArea.border === 'object' ? options.chartArea.border : undefined
	if (options.chartArea.border) {
		options.chartArea.border = {
			color: options.chartArea.border.color || DEF_CHART_BORDER.color,
			pt: options.chartArea.border.pt || DEF_CHART_BORDER.pt,
		}
	}
	options.chartArea.roundedCorners = typeof options.chartArea.roundedCorners === 'boolean' ? options.chartArea.roundedCorners : true
	//
	options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : undefined
	if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt))) options.dataBorder.pt = 0.75
	if (options.dataBorder && options.dataBorder.color) {
		const isHexColor = typeof options.dataBorder.color === 'string' && options.dataBorder.color.length === 6 && /^[0-9A-Fa-f]{6}$/.test(options.dataBorder.color)
		const isSchemeColor = Object.values(SCHEME_COLOR_NAMES).includes(options.dataBorder.color as SCHEME_COLOR_NAMES)
		if (!isHexColor && !isSchemeColor) {
			options.dataBorder.color = 'F9F9F9' // Fallback if neither hex nor scheme color
		}
	}
	//
	if (!options.dataLabelFormatCode && options._type === CHART_TYPE.SCATTER) options.dataLabelFormatCode = 'General'
	if (!options.dataLabelFormatCode && (options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT)) { options.dataLabelFormatCode = options.showPercent ? '0%' : 'General' }
	options.dataLabelFormatCode = options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string' ? options.dataLabelFormatCode : '#,##0'
	//
	// Set default format for Scatter chart labels to custom string if not defined
	if (!options.dataLabelFormatScatter && options._type === CHART_TYPE.SCATTER) options.dataLabelFormatScatter = 'custom'
	//
	options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2
	options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : undefined

	if (options._type === CHART_TYPE.AREA || options._type === CHART_TYPE.BAR || options._type === CHART_TYPE.BAR3D || options._type === CHART_TYPE.LINE) {
		options.catAxisMultiLevelLabels = !!options.catAxisMultiLevelLabels
	} else {
		delete options.catAxisMultiLevelLabels
	}

	// STEP 4: Set props
	resultObject._type = SLIDE_OBJECT_TYPES.chart
	resultObject.options = options as unknown as ObjectOptions
	resultObject.chartRid = getNewRelId(target)

	// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	target._relsChart.push({
		rId: getNewRelId(target),
		data: tmpData as IOptsChartData[],
		opts: options,
		type: options._type,
		globalId: chartId,
		fileName: `chart${chartId}.xml`,
		Target: `/ppt/charts/chart${chartId}.xml`,
	})

	target._slideObjects.push(resultObject)
	return resultObject
}

/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlideInternal} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
/**
 * Register a raster image fill as a slide media relationship and stash the resolved
 * rId on the fill object so serialization can emit `<a:blipFill r:embed="rIdN">`.
 * Mirrors the non-SVG media-registration path used by `addImageDefinition()`,
 * including de-duplication of identical sources (issue #1339). SVG sources are not
 * supported as fills yet.
 * @param {PresSlideInternal} target - slide the owning object belongs to
 * @param {ShapeFillProps} fill - fill options carrying `image: { path | data }`
 */
function registerImageFillMedia(target: PresSlideInternal, fill: ShapeFillProps): void {
	const strImagePath = fill.image?.path || ''
	const strImageData = fill.image?.data || ''

	if (!strImagePath && !strImageData) {
		console.warn('Warning: image fill requires `image.path` or `image.data`; ignoring image fill.')
		fill.type = 'none'
		return
	}
	if (strImageData && !strImageData.toLowerCase().includes('base64,')) {
		console.warn('Warning: image fill `data` value lacks a base64 header (ex: \'image/png;base64,...\'); ignoring image fill.')
		fill.type = 'none'
		return
	}

	// Determine extension: path wins, else sniff the data: mime-type (mirror addImageDefinition())
	const imagePathFile = strImagePath.slice(strImagePath.lastIndexOf('/') + 1).split('?')[0] || ''
	let strImgExtn = ((imagePathFile.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()
	const imageMimeMatch = /image\/(\w+);/.exec(strImageData)
	if (strImageData && imageMimeMatch) strImgExtn = imageMimeMatch[1]
	else if (strImageData?.toLowerCase().includes('image/svg+xml')) strImgExtn = 'svg'

	if (strImgExtn === 'svg') {
		console.warn('Warning: SVG image fills are not supported; ignoring image fill. Use a raster format (PNG/JPEG/GIF/BMP/WebP).')
		fill.type = 'none'
		return
	}

	const imageRelId = getNewRelId(target)
	const mediaSlideKey = target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	const imgContentType = imageContentType(strImgExtn)
	const dupeItem = target._relsMedia.find(item => {
		if (item.isDuplicate || !item.Target || item.type !== imgContentType) return false
		return strImagePath ? item.path === strImagePath : !!strImageData && item.data === strImageData
	})

	target._relsMedia.push({
		path: strImagePath || 'preencoded.' + strImgExtn,
		type: imgContentType,
		extn: strImgExtn,
		data: strImageData || '',
		rId: imageRelId,
		isDuplicate: !!dupeItem?.Target,
		Target: dupeItem?.Target ? dupeItem.Target : `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
	})
	fill.type = 'image'
	fill._imgRid = imageRelId
}

export function addImageDefinition(target: PresSlideInternal, opt: ImageProps): void {
	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.image,
	}

	// Inherit geometry from a matching layout placeholder (issue #1258): an image targeting a
	// placeholder adopts that placeholder's position/size for any of x/y/w/h the caller omits.
	// Explicit `opt` values always win; this only fills the gaps so a picture placeholder no longer
	// collapses to the image's natural/1in fallback when no dimensions are supplied. Mirrors the
	// text-object placeholder inheritance in addTextDefinition() (issue #640).
	let phX: Coord | undefined
	let phY: Coord | undefined
	let phW: Coord | undefined
	let phH: Coord | undefined
	if (opt.placeholder && target._slideLayout?._slideObjects) {
		const placeHold = target._slideLayout._slideObjects.find(
			item => item._type === SLIDE_OBJECT_TYPES.placeholder && item.options?.placeholder === opt.placeholder
		)
		if (placeHold?.options) {
			phX = placeHold.options.x
			phY = placeHold.options.y
			phW = placeHold.options.w
			phH = placeHold.options.h
		}
	}

	// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	const intPosX = opt.x ?? phX ?? 0
	const intPosY = opt.y ?? phY ?? 0
	const intWidth = opt.w ?? phW ?? 0
	const intHeight = opt.h ?? phH ?? 0
	const sizing = opt.sizing
	const objHyperlink = opt.hyperlink || ''
	// Convenience: accept raw SVG markup via `svg` and encode it to a data URI.
	// `data`/`path` win when also supplied, matching the documented precedence.
	const strImageData = opt.data || (opt.svg && !opt.path ? svgMarkupToDataUri(opt.svg) : '')
	const strImagePath = opt.path || ''
	let imageRelId = getNewRelId(target)
	const objectName = opt.objectName ? encodeXmlEntities(validateObjectName(opt.objectName, 'image')) : `Image ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.image).length}`

	// REALITY-CHECK:
	if (!strImagePath && !strImageData) {
		console.error('ERROR: addImage() requires either \'data\' or \'path\' parameter!')
		return
	} else if (strImagePath && typeof strImagePath !== 'string') {
		console.error(`ERROR: addImage() 'path' should be a string, ex: {path:'/img/sample.png'} - you sent ${String(strImagePath)}`)
		return
	} else if (strImageData && typeof strImageData !== 'string') {
		console.error(`ERROR: addImage() 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent ${String(strImageData)}`)
		return
	} else if (strImageData && typeof strImageData === 'string' && !strImageData.toLowerCase().includes('base64,')) {
		console.error('ERROR: Image `data` value lacks a base64 header! Ex: \'image/png;base64,NMP[...]\')')
		return
	}

	// STEP 1: Set extension
	// NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
	const imagePathFile = strImagePath.slice(strImagePath.lastIndexOf('/') + 1).split('?')[0] || ''
	let strImgExtn = ((imagePathFile.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()

	// However, pre-encoded images can be whatever mime-type they want (and good for them!)
	const imageMimeMatch = /image\/(\w+);/.exec(strImageData)
	if (strImageData && imageMimeMatch) {
		strImgExtn = imageMimeMatch[1]
	} else if (strImageData?.toLowerCase().includes('image/svg+xml')) {
		strImgExtn = 'svg'
	}

	// STEP 2: Set type/path
	newObject._type = SLIDE_OBJECT_TYPES.image
	newObject.image = strImagePath || 'preencoded.png'

	// STEP 3: Default any missing dimension from the image's intrinsic (natural) size.
	// For base64 `data` images the bytes are already in hand, so we can read the
	// natural pixel size synchronously and avoid the legacy 1x1 fallback that
	// squished data-only images into a 1in square (issue #1351).
	// Path images can't be measured synchronously (bytes load async during export),
	// so the missing extent is flagged via `_szAuto` and backfilled at serialize time
	// once the media bytes are available (issue #1217).
	// PowerPoint inserts images at 96 DPI, so natural pixels / 96 == inches.
	let defWidth = intWidth
	let defHeight = intHeight
	let szAuto: { w: boolean, h: boolean } | undefined
	if ((!intWidth || !intHeight) && strImgExtn !== 'svg') {
		const natural = strImageData ? getImageSizeFromBase64(strImageData) : null
		if (natural) {
			if (!intWidth && !intHeight) {
				// Neither given: use the natural size (inches @ 96 DPI)
				defWidth = natural.w / IMAGE_NATURAL_DPI
				defHeight = natural.h / IMAGE_NATURAL_DPI
			} else if (typeof intWidth === 'number' && intWidth && !intHeight) {
				// Only width given: preserve aspect ratio for height (same unit as width)
				defHeight = intWidth * (natural.h / natural.w)
			} else if (typeof intHeight === 'number' && intHeight && !intWidth) {
				// Only height given: preserve aspect ratio for width (same unit as height)
				defWidth = intHeight * (natural.w / natural.h)
			}
		} else if (strImagePath) {
			// Path image: defer measurement to serialize time. Record which side(s) to derive
			// from the natural ratio; the 1in fallback below still applies if it stays unmeasurable.
			szAuto = { w: !intWidth, h: !intHeight }
		}
	}

	// STEP 4: Set image properties & options
	const objectOptions: ObjectOptions = {
		x: intPosX || 0,
		y: intPosY || 0,
		w: defWidth || 1,
		h: defHeight || 1,
		altText: opt.altText || '',
		rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
		shape: opt.shape,
		points: opt.points,
		rectRadius: opt.rectRadius,
		shapeAdjust: opt.shapeAdjust,
		sizing,
		crop: opt.crop,
		placeholder: opt.placeholder,
		rotate: opt.rotate || 0,
		flipV: opt.flipV || false,
		flipH: opt.flipH || false,
		transparency: opt.transparency || 0,
		duotone: opt.duotone,
		objectName,
		objectLock: opt.objectLock,
		shadow: correctShadowOptions(opt.shadow),
		...(szAuto ? { _szAuto: szAuto } : {}),
	}
	newObject.options = objectOptions

	// STEP 5: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	// Use a namespaced key for media targets so slide master (sm) and slide layouts (sl-N, _slideNum >= 1000)
	// never collide with regular slide media names in large decks (issue #1416).
	const mediaSlideKey = target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	if (strImgExtn === 'svg') {
		// SVG files consume *TWO* rId's: (a png version and the svg image)
		// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
		// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
		target._relsMedia.push({
			path: strImagePath || strImageData + 'png',
			type: 'image/png',
			extn: 'png',
			data: strImageData || '',
			rId: imageRelId,
			Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.png`,
			isSvgPng: true,
			svgSize: { w: getSmartParseNumber(objectOptions.w, 'X', target._presLayout), h: getSmartParseNumber(objectOptions.h, 'Y', target._presLayout) },
		})
		newObject.imageRid = imageRelId
		target._relsMedia.push({
			path: strImagePath || strImageData,
			type: 'image/svg+xml',
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId + 1,
			Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		newObject.imageRid = imageRelId + 1
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy.
		// File-path images are matched by `path`; base64/`data` images have no real path
		// (all share the `preencoded.<extn>` placeholder), so they are matched by their data
		// payload instead so identical inline images are embedded once (issue #1339).
		const imgContentType = imageContentType(strImgExtn)
		const dupeItem = target._relsMedia.find(item => {
			if (item.isDuplicate || !item.Target || item.type !== imgContentType) return false
			return strImagePath ? item.path === strImagePath : !!strImageData && item.data === strImageData
		})

		target._relsMedia.push({
			path: strImagePath || 'preencoded.' + strImgExtn,
			type: imgContentType,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId,
			isDuplicate: !!(dupeItem?.Target),
			Target: dupeItem?.Target ? dupeItem.Target : `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		newObject.imageRid = imageRelId
	}

	// STEP 6: Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide) throw new Error('ERROR: `hyperlink` option requires either: `url` or `slide`')
		else {
			imageRelId++

			target._rels.push({
				type: SLIDE_OBJECT_TYPES.hyperlink,
				data: objHyperlink.slide ? 'slide' : 'dummy',
				rId: imageRelId,
				Target: objHyperlink.url ? encodeXmlEntities(objHyperlink.url) : String(objHyperlink.slide),
			})

			objHyperlink._rId = imageRelId
			newObject.hyperlink = objHyperlink
		}
	}

	// STEP 7: Add object to slide
	target._slideObjects.push(newObject)
}

/**
 * Adds a media object to a slide definition.
 * @param {PresSlideInternal} `target` - slide object that the media will be added to
 * @param {MediaProps} `opt` - media options
 */
export function addMediaDefinition(target: PresSlideInternal, opt: MediaProps): void {
	const intPosX = opt.x || 0
	const intPosY = opt.y || 0
	const intSizeX = opt.w || 2
	const intSizeY = opt.h || 2
	const strData = opt.data || ''
	const strLink = opt.link || ''
	const strPath = opt.path || ''
	const strType = opt.type || 'audio'
	let strExtn = ''
	const strCover = opt.cover || IMG_PLAYBTN
	const objectName = opt.objectName ? encodeXmlEntities(validateObjectName(opt.objectName, 'media')) : `Media ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.media).length}`
	const slideData: ISlideObject = { _type: SLIDE_OBJECT_TYPES.media }

	// STEP 1: REALITY-CHECK
	if (!strPath && !strData && strType !== 'online') {
		throw new Error('addMedia() error: either `data` or `path` are required!')
	} else if (strData && !strData.toLowerCase().includes('base64,')) {
		throw new Error('addMedia() error: `data` value lacks a base64 header! Ex: \'video/mpeg;base64,NMP[...]\')')
	} else if (strCover && !strCover.toLowerCase().includes('base64,')) {
		throw new Error('addMedia() error: `cover` value lacks a base64 header! Ex: \'data:image/png;base64,iV[...]\')')
	}
	// Online Video: requires `link`
	if (strType === 'online' && !strLink) {
		throw new Error('addMedia() error: online videos require `link` value')
	}

	// FIXME: 20190707
	// strType = strData ? strData.split(';')[0].split('/')[0] : strType
	strExtn = opt.extn || (strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop()) || 'mp3'

	// STEP 2: Set type, media
	slideData.mtype = strType
	slideData.media = strPath || 'preencoded.mov'
	slideData.options = {}

	// Playback looping (embedded audio/video only; online embeds have no timing tree)
	if (strType !== 'online') {
		if (opt.loop) slideData.loop = true
		else if (typeof opt.loopCount === 'number' && opt.loopCount > 0) slideData.loopCount = opt.loopCount
	}

	// STEP 3: Set media properties & options
	slideData.options.x = intPosX
	slideData.options.y = intPosY
	slideData.options.w = intSizeX
	slideData.options.h = intSizeY
	slideData.options.objectName = objectName
	if (opt.altText) slideData.options.altText = opt.altText
	if (opt.objectLock) slideData.options.objectLock = opt.objectLock

	// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	/**
	 * NOTE:
	 * - rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	 *
	 * NOTE:
	 * - Audio/Video files consume *TWO* rId's:
	 * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
	 * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
	 */
	if (strType === 'online') {
		const relId1 = getNewRelId(target)
		// A: ECMA video rel (external link) — referenced by <a:videoFile r:link>.
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: relId1,
			Target: strLink,
		})
		slideData.mediaRid = relId1

		// B: MS-2007 media rel — PowerPoint authors a second external rel sharing the
		// same link Target; the body points at it via <p14:media r:link>. (Mirrors the
		// embedded A/V pair, but External and with no media binary part.)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: getNewRelId(target),
			Target: strLink,
		})

		// C: Add cover (preview/overlay) image
		target._relsMedia.push({
			path: 'preencoded.png',
			data: strCover,
			type: 'image/png',
			extn: 'png',
			rId: getNewRelId(target),
			Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`,
		})
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy.
		// Path-based media match by `path`; base64/`data` media (which share the `preencoded`
		// placeholder path) match by their data payload so identical inline media embed once (issue #1339).
		const dupeItem = target._relsMedia.find(item => {
			if (item.isDuplicate || !item.Target || item.type !== strType + '/' + strExtn) return false
			return strPath ? item.path === strPath : !!strData && item.data === strData
		})

		// A: "relationships/video"
		const relId1 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: relId1,
			isDuplicate: !!(dupeItem?.Target),
			Target: dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 1}.${strExtn}`,
		})
		slideData.mediaRid = relId1

		// B: "relationships/media"
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: getNewRelId(target),
			isDuplicate: !!(dupeItem?.Target),
			Target: dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 0}.${strExtn}`,
		})

		// C: Add cover (preview/overlay) image
		target._relsMedia.push({
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			data: strCover,
			rId: getNewRelId(target),
			Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`,
		})
	}

	// LAST
	target._slideObjects.push(slideData)
}

/**
 * Adds Notes to a slide.
 * @param {PresSlideInternal} `target` slide object
 * @param {string | NotesProps | NotesProps[]} `notes` plain text, or rich runs (inline formatting / hyperlinks)
 * @since 2.3.0
 */
export function addNotesDefinition(target: PresSlideInternal, notes: string | NotesProps | NotesProps[]): void {
	// Normalize all input forms to a TextProps[] run list so the notes-slide serializer
	// (which reuses the standard text-run generator) can handle plain and rich notes uniformly.
	const runs: TextProps[] =
		typeof notes === 'string'
			? [{ text: notes }]
			: (Array.isArray(notes) ? notes : [notes]).map(run => ({ text: run.text, options: run.options }))

	target._slideObjects.push({
		_type: SLIDE_OBJECT_TYPES.notes,
		text: runs,
	})
}

/**
 * Derive 1-2 letter initials from an author display name (e.g. "Ada Lovelace" -> "AL").
 * Falls back to the first character when the name is a single word.
 */
function deriveAuthorInitials(author: string): string {
	const words = author.trim().split(/\s+/).filter(Boolean)
	if (words.length === 0) return '?'
	if (words.length === 1) return words[0].charAt(0).toUpperCase()
	return (words[0].charAt(0) + words[words.length - 1].charAt(0)).toUpperCase()
}

/**
 * Adds a review comment to a slide (legacy ISO/IEC 29500 §13 comment).
 * @param {PresSlideInternal} target slide object the comment is attached to
 * @param {CommentProps} opts comment author/text/position options
 * @since v4.1.0
 */
export function addCommentDefinition(target: PresSlideInternal, opts: CommentProps): void {
	const author = typeof opts?.author === 'string' ? opts.author.trim() : ''
	const text = typeof opts?.text === 'string' ? opts.text : ''
	// Don't silently coerce: a comment with no author or no body is meaningless, so warn + skip
	// rather than emit a degenerate <p:cm> (API policy: warn over silent coercion).
	if (!author) {
		console.warn('Warning: addComment() requires a non-empty `author`; comment ignored.')
		return
	}
	if (!text) {
		console.warn('Warning: addComment() requires non-empty `text`; comment ignored.')
		return
	}

	const initials = typeof opts.initials === 'string' && opts.initials.trim() ? opts.initials.trim() : deriveAuthorInitials(author)
	const x = typeof opts.x === 'number' && Number.isFinite(opts.x) ? opts.x : 0.5
	const y = typeof opts.y === 'number' && Number.isFinite(opts.y) ? opts.y : 0.5
	let date: string | undefined
	if (opts.date instanceof Date) date = opts.date.toISOString()
	else if (typeof opts.date === 'string' && opts.date) date = opts.date

	if (!target._comments) target._comments = []
	target._comments.push({ author, initials, text, x, y, date })
}

/**
 * Map of common friendly shape names users pass as bare strings to their
 * valid OOXML preset values. PowerPoint can't parse the friendly spellings
 * and removes the shape during repair .
 */
const SHAPE_NAME_ALIASES: { [key: string]: SHAPE_NAME } = {
	oval: 'ellipse',
	rectangle: 'rect',
	roundedRectangle: 'roundRect',
	roundedrectangle: 'roundRect',
}

/**
 * Adds a shape object to a slide definition.
 * @param {PresSlideInternal} target slide object that the shape should be added to
 * @param {SHAPE_NAME} shapeName shape name
 * @param {ShapeProps} opts shape options
 */
export function addShapeDefinition(target: PresSlideInternal, shapeName: SHAPE_NAME, opts: ShapeProps): void {
	const options = typeof opts === 'object' ? opts : {}
	options.line = options.line || { type: 'none' }
	options.shadow = correctShadowOptions(options.shadow)
	// Normalize friendly shape names (e.g. "oval" -> "ellipse") to their valid
	// OOXML preset spellings before storing on the slide object.
	const resolvedShapeName: SHAPE_NAME = (typeof shapeName === 'string' && SHAPE_NAME_ALIASES[shapeName])
		? SHAPE_NAME_ALIASES[shapeName]
		: shapeName
	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.text,
		shape: resolvedShapeName || SHAPE_TYPE.RECTANGLE,
		options,
	}

	// Reality check
	if (!shapeName) throw new Error('Missing/Invalid shape parameter! Example: `addShape(pptxgen.shapes.LINE, {x:1, y:1, w:1, h:1});`')

	// Reject presets PowerPoint can't parse. An invalid `prst` value (a typo or an
	// unmapped friendly name) corrupts the package and triggers the repair dialog,
	// so fail loudly here rather than emit degenerate OOXML. Use `pptxgen.shapes.*`
	// for the canonical names.
	if (!VALID_SHAPE_PRESETS.has(resolvedShapeName)) {
		throw new Error(`Invalid shape "${String(shapeName)}"! Use a value from \`pptxgen.shapes.*\` (e.g. \`pptxgen.shapes.RECTANGLE\`). PowerPoint can't render unknown preset geometries and will drop the shape during repair.`)
	}

	// 1: ShapeLineProps defaults
	// A stroke can carry a non-solid paint (a `gradient`) just like a fill, so infer the
	// stroke `type` from the gradient when the caller omits it (`line: { gradient }`) and
	// preserve the gradient through normalization. Only a solid stroke gets the default
	// line color; a gradient stroke takes its colors from its stops.
	const lineType = options.line.type || (options.line.gradient ? 'gradient' : 'solid')
	const newLineOpts: ShapeLineProps = {
		type: lineType,
		color: lineType === 'solid' ? options.line.color || DEF_SHAPE_LINE_COLOR : options.line.color,
		transparency: options.line.transparency || 0,
		width: options.line.width || 1,
		dashType: options.line.dashType || 'solid',
		beginArrowType: options.line.beginArrowType,
		endArrowType: options.line.endArrowType,
		gradient: options.line.gradient,
	}
	if (typeof options.line === 'object' && options.line.type !== 'none') options.line = newLineOpts

	// 2: Set options defaults
	options.x = options.x || (options.x === 0 ? 0 : 1)
	options.y = options.y || (options.y === 0 ? 0 : 1)
	options.w = options.w || (options.w === 0 ? 0 : 1)
	options.h = options.h || (options.h === 0 ? 0 : 1)
	options.objectName = options.objectName
		? encodeXmlEntities(validateObjectName(options.objectName, 'shape'))
		: `Shape ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.text).length}`

	// 3: Handle line (lots of deprecated opts)
	if (typeof options.line === 'string') {
		const tmpOpts = newLineOpts
		tmpOpts.color = String(options.line) // @deprecated `options.line` string (was line color)
		options.line = tmpOpts
	}
	if (typeof options.lineSize === 'number') options.line.width = options.lineSize // @deprecated (part of `ShapeLineProps` now)
	if (typeof options.lineDash === 'string') options.line.dashType = options.lineDash // @deprecated (part of `ShapeLineProps` now)
	if (typeof options.lineHead === 'string') options.line.beginArrowType = options.lineHead // @deprecated (part of `ShapeLineProps` now)
	if (typeof options.lineTail === 'string') options.line.endArrowType = options.lineTail // @deprecated (part of `ShapeLineProps` now)

	// 4: Create hyperlink rels
	createHyperlinkRels(target, newObject)

	// 5: Register an image fill (if any) as a media relationship for serialize-time blipFill
	if (typeof options.fill === 'object' && (options.fill.type === 'image' || options.fill.image)) {
		registerImageFillMedia(target, options.fill)
	}

	// LAST: Add object to slide
	target._slideObjects.push(newObject)
}

/**
 * Adds a connector object to a slide definition.
 * A connector is a line between two points emitted as a PowerPoint connector (`<p:cxnSp>`).
 * Endpoints are converted to a bounding box (`x/y/w/h`) plus `flipH`/`flipV` so the box can be
 * oriented from any corner; the connector preset geometry is derived from `type`.
 * @param {PresSlideInternal} target - slide the connector is added to
 * @param {ConnectorProps} opts - connector options (endpoints + line styling)
 */
export function addConnectorDefinition(target: PresSlideInternal, opts: ConnectorProps): void {
	if (!opts || [opts.x1, opts.y1, opts.x2, opts.y2].some(v => typeof v === 'undefined')) {
		throw new Error('addConnector requires { x1, y1, x2, y2 }. Example: `slide.addConnector({ x1:1, y1:1, x2:4, y2:3 })`')
	}

	const type = opts.type || 'straight'
	if (type !== 'straight' && type !== 'elbow' && type !== 'curved') {
		throw new Error(`Invalid connector type "${String(type)}". Use 'straight', 'elbow', or 'curved'.`)
	}

	// Resolve the preset variant + adjust guides. `bentConnector{3,4,5}` / `curvedConnector{3,4,5}`
	// each expose `bends` adjustable jogs as `<a:gd name="adjN" fmla="val …"/>` (1000ths-of-a-percent,
	// so 50% → 50000; values verified against PowerPoint-authored decks). `straightConnector1` has none.
	const adjInput = opts.adj === undefined ? [] : Array.isArray(opts.adj) ? opts.adj : [opts.adj]
	const bends = opts.bends ?? (adjInput.length || 1)
	let connectorAdj: number[] = []
	if (type === 'straight') {
		if (opts.bends !== undefined || opts.adj !== undefined) {
			console.warn('Warning: addConnector `bends`/`adj` are ignored for type "straight" (a straight connector has no bends).')
		}
	} else {
		if (bends !== 1 && bends !== 2 && bends !== 3) {
			throw new Error(`addConnector \`bends\` must be 1, 2, or 3 (got ${String(bends)}).`)
		}
		if (opts.adj !== undefined && adjInput.length !== bends) {
			throw new Error(`addConnector \`adj\` must supply ${bends} value(s) to match \`bends\`=${bends} (got ${adjInput.length}).`)
		}
		// Convert each percent to OOXML 1000ths-of-a-percent. Fail loud on non-finite input
		// (silent coercion would emit a degenerate guide PowerPoint repairs); warn but allow
		// out-of-range, which legitimately places a jog beyond the endpoint box.
		connectorAdj = adjInput.map((pct, i) => {
			if (typeof pct !== 'number' || !Number.isFinite(pct)) {
				throw new Error(`addConnector \`adj\` value #${i + 1} must be a finite number (percent 0–100); got ${String(pct)}.`)
			}
			if (pct < 0 || pct > 100) {
				console.warn(`Warning: addConnector \`adj\` value ${pct} is outside 0–100; the bend will sit beyond the endpoint box.`)
			}
			return Math.round(pct * 1000)
		})
	}
	const preset = connectorPresetFor(type, bends)

	// Optional shape binding (<a:stCxn>/<a:endCxn>). The target id is resolved at serialize time
	// (it equals the shape's slide-object index + 2); here we just capture the name + site index.
	// The site index must be a non-negative integer — a bad idx makes PowerPoint repair the connector.
	const resolveCxn = (shapeName: string | undefined, idx: number | undefined, end: 'startShape' | 'endShape'): { name: string, idx: number } | undefined => {
		if (shapeName === undefined) return undefined
		if (typeof shapeName !== 'string' || shapeName.trim().length === 0) {
			throw new Error(`addConnector \`${end}\` must be a non-empty shape objectName.`)
		}
		const site = idx ?? 0
		if (!Number.isInteger(site) || site < 0) {
			throw new Error(`addConnector \`${end}Idx\` must be a non-negative integer (got ${String(site)}).`)
		}
		// Match the shape's stored objectName, which is XML-entity-encoded at add time.
		return { name: encodeXmlEntities(shapeName), idx: site }
	}
	const startCxn = resolveCxn(opts.startShape, opts.startShapeIdx, 'startShape')
	const endCxn = resolveCxn(opts.endShape, opts.endShapeIdx, 'endShape')

	// Resolve all four endpoints to inches up front (handles every `Coord` form: number,
	// '50%', '2in', etc.). The connector box uses the min corner as its origin and flips
	// horizontally/vertically when the end point is left of / above the start point.
	const x1 = getSmartParseNumber(opts.x1, 'X', target._presLayout) / EMU
	const y1 = getSmartParseNumber(opts.y1, 'Y', target._presLayout) / EMU
	const x2 = getSmartParseNumber(opts.x2, 'X', target._presLayout) / EMU
	const y2 = getSmartParseNumber(opts.y2, 'Y', target._presLayout) / EMU

	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.connector,
		// store the connector preset on `shape`; the serializer emits it as the prstGeom `prst`
		shape: preset,
		options: {
			x: Math.min(x1, x2),
			y: Math.min(y1, y2),
			w: Math.abs(x2 - x1),
			h: Math.abs(y2 - y1),
			flipH: x2 < x1,
			flipV: y2 < y1,
			_connectorAdj: connectorAdj.length ? connectorAdj : undefined,
			_startCxn: startCxn,
			_endCxn: endCxn,
			line: {
				type: 'solid',
				color: opts.color || DEF_SHAPE_LINE_COLOR,
				width: typeof opts.width === 'number' ? opts.width : 1,
				dashType: opts.dashType || 'solid',
				beginArrowType: opts.beginArrowType,
				endArrowType: opts.endArrowType,
			},
			altText: opts.altText,
			objectName: opts.objectName
				? encodeXmlEntities(validateObjectName(opts.objectName, 'connector'))
				: `Connector ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.connector).length}`,
		},
	}

	target._slideObjects.push(newObject)
}

/**
 * Adds a table object to a slide definition.
 * @param {PresSlideInternal} target - slide object that the table should be added to
 * @param {TableRow[]} tableRows - table data
 * @param {TableProps} options - table options
 * @param {SlideLayoutInternal} slideLayout - Slide layout
 * @param {PresLayout} presLayout - Presentation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
export function addTableDefinition(
	target: PresSlideInternal,
	tableRows: TableRow[],
	options: TableProps,
	slideLayout: SlideLayoutInternal | null,
	presLayout: PresLayout,
	addSlide: (options?: AddSlideProps) => PresSlideInternal,
	getSlide: (slideNumber: number) => PresSlideInternal
): PresSlideInternal[] {
	const slides: PresSlideInternal[] = [target] // Create array of Slides as more may be added by auto-paging
	const opt: TableProps = options && typeof options === 'object' ? options : {}
	opt.objectName = opt.objectName ? encodeXmlEntities(validateObjectName(opt.objectName, 'table')) : `Table ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.table).length}`

	// STEP 0: PLACEHOLDER — a table targeting a layout placeholder inherits that placeholder's
	// position/size for any of x/y/w/h the caller omits (#1151), mirroring the image (#1258) and
	// text (#640) placeholder inheritance. Explicit values always win; this only fills the gaps so
	// the table fills the placeholder geometry rather than the default 1in/full-width fallback.
	if (opt.placeholder && slideLayout?._slideObjects) {
		const placeHold = slideLayout._slideObjects.find(
			item => item._type === SLIDE_OBJECT_TYPES.placeholder && item.options?.placeholder === opt.placeholder
		)
		if (placeHold?.options) {
			if (opt.x === undefined) opt.x = placeHold.options.x
			if (opt.y === undefined) opt.y = placeHold.options.y
			if (opt.w === undefined) opt.w = placeHold.options.w
			if (opt.h === undefined) opt.h = placeHold.options.h
		}
	}

	// STEP 1: REALITY-CHECK
	{
		// A: check for empty
		if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
			throw new Error('addTable: Array expected! EX: \'slide.addTable( [rows], {options} );\' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)')
		}

		// B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
		if (!tableRows[0] || !Array.isArray(tableRows[0])) {
			throw new Error(
				'addTable: \'rows\' should be an array of cells! EX: \'slide.addTable( [ [\'A\'], [\'B\'], {text:\'C\',options:{align:\'center\'}} ] );\' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)'
			)
		}

		// TODO: FUTURE: This is wacky and wont function right (shows .w value when there is none from demo.js?!) 20191219
		/*
		if (opt.w && opt.colW) {
			console.warn('addTable: please use either `colW` or `w` - not both (table will use `colW` and ignore `w`)')
			console.log(`${opt.w} ${opt.colW}`)
		}
		*/
	}

	// STEP 2: Transform `tableRows` into well-formatted TableCell's
	// tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
	const arrRows: TableCell[][] = []
	tableRows.forEach(row => {
		const newRow: TableCell[] = []

		if (Array.isArray(row)) {
			row.forEach((cell: number | string | TableCell) => {
				// A:
				const newCellOptions = typeof cell === 'object' && cell.options ? cell.options : {}
				const newCell: TableCell = {
					_type: SLIDE_OBJECT_TYPES.tablecell,
					text: '',
					options: newCellOptions,
				}

				// B:
				if (typeof cell === 'string' || typeof cell === 'number') newCell.text = cell.toString()
				else if (cell.text) {
					// Cell can contain complex text type, or string, or number
					if (typeof cell.text === 'string' || typeof cell.text === 'number') newCell.text = cell.text.toString()
					else if (cell.text) newCell.text = cell.text
					// Capture options
					if (cell.options && typeof cell.options === 'object') newCell.options = cell.options
				}

				// C: Set cell borders
				newCellOptions.border = newCellOptions.border || opt.border || [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }]
				let cellBorder = newCellOptions.border

				// CASE 1: border interface is: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
				if (cellBorder && typeof cellBorder === 'object') {
					cellBorder = normalizeBorderTuple(cellBorder)
					newCellOptions.border = cellBorder
				}
				// Handle: [null, null, {type:'solid'}, null]
				const cellBorderTuple = newCellOptions.border as BorderTuple
				if (!cellBorderTuple[0]) cellBorderTuple[0] = { type: 'none' }
				if (!cellBorderTuple[1]) cellBorderTuple[1] = { type: 'none' }
				if (!cellBorderTuple[2]) cellBorderTuple[2] = { type: 'none' }
				if (!cellBorderTuple[3]) cellBorderTuple[3] = { type: 'none' }

				// set complete BorderOptions for all sides
				const arrSides = [0, 1, 2, 3] as const
				arrSides.forEach(idx => {
					cellBorderTuple[idx] = {
						type: cellBorderTuple[idx].type || DEF_CELL_BORDER.type,
						color: cellBorderTuple[idx].color || DEF_CELL_BORDER.color,
						pt: typeof cellBorderTuple[idx].pt === 'number' ? cellBorderTuple[idx].pt : DEF_CELL_BORDER.pt,
					}
				})
				newCellOptions.border = cellBorderTuple

				// LAST:
				newRow.push(newCell)
			})
		} else {
			console.log('addTable: tableRows has a bad row. A row should be an array of cells. You provided:')
			console.log(row)
		}

		arrRows.push(newRow)
	})

	// STEP 3: Set options
	// Keep x/y/w/h as raw user `Coord` (inches/percent/unit-string). They are resolved to EMU
	// exactly once at emission (gen-xml) and by the auto-pager (getSlidesForTableRows); no
	// pre-conversion here, so a value is never parsed twice. Default position is 0.5in.
	if (opt.x === undefined || opt.x === null) opt.x = 0.5
	if (opt.y === undefined || opt.y === null) opt.y = 0.5
	// NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_IN
	if (typeof opt.margin === 'number') opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	// defensive fallback - if `opt.margin` is not a 4-element array of finite numbers, use defaults so non-numeric table-level margins don't leak NaN into <a:tcPr>
	if (!Array.isArray(opt.margin) || opt.margin.length !== 4 || opt.margin.some((v: unknown) => typeof v !== 'number' || !Number.isFinite(v))) {
		opt.margin = DEF_CELL_MARGIN_IN
	}
	// NOTE: dont add default color on tables with hyperlinks! (it causes any textObj's with hyperlinks to have subsequent words to be black)
	if (!JSON.stringify({ arrRows }).includes('hyperlink')) {
		if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR // Set default color if needed (table option > inherit from Slide > default to black)
	}
	if (typeof opt.border === 'string') {
		console.warn('addTable `border` option must be an object. Ex: `{border: {type:\'none\'}}`')
		opt.border = undefined
	} else if (Array.isArray(opt.border)) {
		const border = opt.border
		;([0, 1, 2, 3] as const).forEach(idx => {
			border[idx] = border[idx]
				? { type: border[idx].type || DEF_CELL_BORDER.type, color: border[idx].color || DEF_CELL_BORDER.color, pt: border[idx].pt || DEF_CELL_BORDER.pt }
				: { type: 'none' }
		})
	}

	opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false
	opt.autoPagePlaceholder = typeof opt.autoPagePlaceholder === 'boolean' ? opt.autoPagePlaceholder : false
	opt.autoPageRepeatHeader = typeof opt.autoPageRepeatHeader === 'boolean' ? opt.autoPageRepeatHeader : false
	opt.autoPageHeaderRows = typeof opt.autoPageHeaderRows !== 'undefined' && !isNaN(Number(opt.autoPageHeaderRows)) ? Number(opt.autoPageHeaderRows) : 1
	opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0
	if (opt.autoPageLineWeight) {
		if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1
		else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1
	}
	// autoPage ^^^

	// Set/Calc table width
	// Get slide margins - start with default values, then adjust if master or slide margins exist
	let arrTableMargin = DEF_SLIDE_MARGIN_IN
	// Case 1: Master margins
	if (slideLayout && typeof slideLayout._margin !== 'undefined') {
		if (Array.isArray(slideLayout._margin)) arrTableMargin = slideLayout._margin
		else if (!isNaN(Number(slideLayout._margin))) { arrTableMargin = [Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin)] }
	}
	// Case 2: Table margins
	/* FIXME: add `_margin` option to slide options
		else if ( addNewSlide._margin ) {
			if ( Array.isArray(addNewSlide._margin) ) arrTableMargin = addNewSlide._margin;
			else if ( !isNaN(Number(addNewSlide._margin)) ) arrTableMargin = [Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin)];
		}
	*/

	/**
	 * Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
	 * The API does not require a `w` value, but XML generation does, hence, code to calc a width below using colW value(s)
	 */
	if (opt.colW) {
		const firstRowColCnt = arrRows[0].reduce((totalLen, c) => {
			if (c?.options?.colspan && typeof c.options.colspan === 'number') {
				totalLen += c.options.colspan
			} else {
				totalLen += 1
			}
			return totalLen
		}, 0)

		if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
			// Ex: `colW = 3` or `colW = '3'`
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length === 1 && firstRowColCnt > 1) {
			// Ex: `colW=[3]` but with >1 cols (same as above, user is saying "use this width for all")
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== firstRowColCnt) {
			// Err: Mismatched colW and cols count
			console.warn('addTable: mismatch: (colW.length != data.length) Therefore, defaulting to evenly distributed col widths.')
			opt.colW = undefined
		}
	} else if (opt.w) {
		// Keep raw user `Coord` — resolved to EMU once at emission. (No pre-conversion.)
	} else {
		opt.w = Math.floor((presLayout._sizeW || presLayout.width) / EMU - arrTableMargin[1] - arrTableMargin[3])
	}

	// STEP 5: Loop over cells: transform each to ITableCell; check to see whether to unset `autoPage` while here
	arrRows.forEach(row => {
		row.forEach((cell, idy) => {
			// A: Transform cell data if needed
			/* Table rows can be an object or plain text - transform into object when needed
				// EX:
				const arrTabRows1 = [
					[ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
					,[ 'B2', 'C2', 'D2', 'E2' ]
				]
			*/
			if (typeof cell === 'number' || typeof cell === 'string') {
				// Grab table formatting `opts` to use here so text style/format inherits as it should
				row[idy] = { _type: SLIDE_OBJECT_TYPES.tablecell, text: String(row[idy]), options: opt }
			} else if (typeof cell === 'object') {
				// ARG0: `text`
				if (typeof cell.text === 'number') row[idy].text = cell.text.toString()
				else if (typeof cell.text === 'undefined' || cell.text === null) row[idy].text = ''

				// ARG1: `options`: ensure options exists
				row[idy].options = cell.options || {}

				// Set type to tabelcell
				row[idy]._type = SLIDE_OBJECT_TYPES.tablecell
			}

			// B: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
			// if (cell.text && Array.isArray(cell.text)) opt.autoPage = false
			// TODO: FIXME: WIP: 20210807: We cant do this anymore
		})
	})

	// If autoPage = true, we need to return references to newly created slides if any
	const newAutoPagedSlides: PresSlideInternal[] = []

	// STEP 6: Auto-Paging: (via {options} and used internally)
	// (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && !opt.autoPage) {
		// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
		createHyperlinkRels(target, arrRows)

		// Add slideObjects (NOTE: Use `extend` to avoid mutation)
		target._slideObjects.push({
			_type: SLIDE_OBJECT_TYPES.table,
			arrTabRows: arrRows,
			options: { ...opt },
		})
	} else {
		if (opt.autoPageRepeatHeader) opt._arrObjTabHeadRows = arrRows.filter((_row, idx) => idx < (opt.autoPageHeaderRows || 1))

		// #1136: snapshot populated placeholders on the source slide (e.g. a title added via
		// `addText(text, { placeholder })`) so they can be re-rendered on each overflow slide.
		// Overflow slides otherwise inherit only the layout's empty placeholders. Captured before
		// the loop so the table object added per-slide below is never included.
		const sourcePlaceholders =
			opt.autoPagePlaceholder && Array.isArray(target._slideObjects)
				? target._slideObjects.filter(obj => obj._type !== SLIDE_OBJECT_TYPES.table && obj.options?.placeholder)
				: []

		// Loop over rows and create 1-N tables as needed (ISSUE#21)
		getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((slide, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			if (!getSlide(target._slideNum + idx)) slides.push(addSlide({ masterName: slideLayout?._name || undefined }))

			// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			// Keep raw inches — resolved to EMU once at emission. (No pre-conversion.)
			if (idx > 0) opt.y = opt.autoPageSlideStartY || opt.newSlideStartY || arrTableMargin[0]

			// C: Add this table to new Slide
			{
				const newSlide: PresSlideInternal = getSlide(target._slideNum + idx)

				opt.autoPage = false

				// #1136: copy the source slide's populated placeholders onto each overflow slide
				// (idx 0 is the source slide itself and already has them).
				if (idx > 0 && sourcePlaceholders.length > 0) {
					sourcePlaceholders.forEach(ph => newSlide._slideObjects.push(structuredClone(ph)))
				}

				// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
				createHyperlinkRels(newSlide, slide.rows)

				// Add rows to new slide. When `rowH` is an array it is keyed by *original* row index,
				// which no longer matches the per-slide physical row order after pagination; use the
				// per-slide heights the auto-pager resolved so each row keeps its configured height
				// instead of inheriting whatever row lands at the same index (#1145).
				// `slide.rowH` may contain `undefined` holes (auto-height rows); the table serializer
				// treats a falsy per-row height as "auto", so the cast to number[] is safe.
				newSlide.addTable(slide.rows, { ...opt, rowH: Array.isArray(opt.rowH) && slide.rowH ? (slide.rowH as number[]) : opt.rowH })

				// Add reference to the new slide so it can be returned, but don't add the first one because the user already has a reference to that one.
				if (idx > 0) newAutoPagedSlides.push(newSlide)
			}
		})
	}
	return newAutoPagedSlides
}

/**
 * Adds a text object to a slide definition.
 * @param {PresSlideInternal} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder whether this a placeholder object
 * @since: 1.0.0
 */
export function addTextDefinition(target: PresSlideInternal, text: TextProps[], opts: TextPropsOptions, isPlaceholder: boolean): void {
	const textObjects = !text || text.length === 0 ? [{ text: '' }] : text
	const objectOptions: ObjectOptions = opts || {}
	const newObject: ISlideObject = {
		_type: isPlaceholder ? SLIDE_OBJECT_TYPES.placeholder : SLIDE_OBJECT_TYPES.text,
		shape: opts.shape || SHAPE_TYPE.RECTANGLE,
		text: textObjects,
		options: objectOptions,
	}

	function cleanOpts(itemOpts: ObjectOptions): TextPropsOptions {
		// STEP 1: Set some options
		{
			// A.1: Color (placeholders should inherit their colors or override them, so don't default them)
			if (!itemOpts.placeholder) {
				// A hyperlink run with no color configured anywhere inherits the theme hyperlink color
				// (a:schemeClr hlink, and folHlink once visited), which PowerPoint applies automatically
				// when the run carries no explicit fill. Defaulting it to DEF_FONT_COLOR would emit a
				// solidFill plus hlinkClr="tx", pinning the link to black and suppressing the theme
				// hyperlink/visited colors (#1165). Only non-hyperlink text falls back to DEF_FONT_COLOR.
				itemOpts.color = itemOpts.color || objectOptions.color || target.color || (itemOpts.hyperlink || objectOptions.hyperlink ? undefined : DEF_FONT_COLOR)
			}

			// A.2: Placeholder should inherit their bullets or override them, so don't default them
			if (itemOpts.placeholder || isPlaceholder) {
				itemOpts.bullet = itemOpts.bullet || false
			}

			// A.3: Text targeting a placeholder need to inherit the placeholders options (eg: margin, valign, etc.) (Issue #640)
			if (itemOpts.placeholder && target._slideLayout && target._slideLayout._slideObjects) {
				const placeHold = target._slideLayout._slideObjects.filter(
					item => item._type === 'placeholder' && item.options && item.options.placeholder && item.options.placeholder === itemOpts.placeholder
				)[0]
				if (placeHold?.options) itemOpts = { ...itemOpts, ...placeHold.options }
			}

			// A.4: Other options
			itemOpts.objectName = itemOpts.objectName
				? encodeXmlEntities(validateObjectName(itemOpts.objectName, 'text'))
				: `Text ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.text).length}`

			// B:
			if (itemOpts.shape === SHAPE_TYPE.LINE) {
				const itemLine = typeof itemOpts.line === 'object' && itemOpts.line ? itemOpts.line : {}
				// ShapeLineProps defaults
				const newLineOpts: ShapeLineProps = {
					type: itemLine.type || 'solid',
					color: itemLine.color || DEF_SHAPE_LINE_COLOR,
					transparency: itemLine.transparency || 0,
					width: itemLine.width || 1,
					dashType: itemLine.dashType || 'solid',
					beginArrowType: itemLine.beginArrowType,
					endArrowType: itemLine.endArrowType,
				}
				if (typeof itemOpts.line === 'object') itemOpts.line = newLineOpts

				// 3: Handle line (lots of deprecated opts)
				if (typeof itemOpts.line === 'string') {
					const tmpOpts = newLineOpts
					if (typeof itemOpts.line === 'string') tmpOpts.color = itemOpts.line // @deprecated [remove in v4.0]
					// tmpOpts.color = itemOpts.line!.toString() // @deprecated `itemOpts.line`:[string] (was line color)
					itemOpts.line = tmpOpts
				}
				const lineOpts = itemOpts.line || newLineOpts
				itemOpts.line = lineOpts
				if (typeof itemOpts.lineSize === 'number') lineOpts.width = itemOpts.lineSize // @deprecated (part of `ShapeLineProps` now)
				if (typeof itemOpts.lineDash === 'string') lineOpts.dashType = itemOpts.lineDash // @deprecated (part of `ShapeLineProps` now)
				if (typeof itemOpts.lineHead === 'string') lineOpts.beginArrowType = itemOpts.lineHead // @deprecated (part of `ShapeLineProps` now)
				if (typeof itemOpts.lineTail === 'string') lineOpts.endArrowType = itemOpts.lineTail // @deprecated (part of `ShapeLineProps` now)
			}

			// C: Line opts
			itemOpts.line = itemOpts.line || {}
			itemOpts.lineSpacing = itemOpts.lineSpacing && !isNaN(itemOpts.lineSpacing) ? itemOpts.lineSpacing : undefined
			itemOpts.lineSpacingMultiple = itemOpts.lineSpacingMultiple && !isNaN(itemOpts.lineSpacingMultiple) ? itemOpts.lineSpacingMultiple : undefined

			// D: Transform text options to bodyProperties as thats how we build XML
			itemOpts._bodyProp = itemOpts._bodyProp || {}
			itemOpts._bodyProp.autoFit = itemOpts.autoFit || false // DEPRECATED: (3.3.0) If true, shape will collapse to text size (Fit To shape)
			itemOpts._bodyProp.anchor = !itemOpts.placeholder ? TEXT_VALIGN.ctr : undefined // VALS: [t,ctr,b]
			// `textDirection` is the documented public option; `vert` is a legacy/extended alias kept as an
			// escape hatch for the full ST_TextVerticalType range (eaVert, mongolianVert, wordArtVertRtl).
			// Both map directly to the `<a:bodyPr vert="…">` attribute, so prefer the documented one.
			itemOpts._bodyProp.vert = itemOpts.textDirection ?? itemOpts.vert // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			itemOpts._bodyProp.wrap = typeof itemOpts.wrap === 'boolean' ? itemOpts.wrap : true
			itemOpts._bodyProp.prstTxWarp = itemOpts.textWarp // preset text warp (`<a:prstTxWarp>`), e.g. 'textArchUp'

			// D.1: Text columns (`numCol` range is 1-16 per ECMA-376 ST_TextColumnCount)
			if (itemOpts.columns !== undefined) {
				if (typeof itemOpts.columns !== 'number' || isNaN(itemOpts.columns) || itemOpts.columns < 1 || itemOpts.columns > 16) {
					console.warn('Warning: text `columns` must be a number 1-16 (ignoring value)')
				} else {
					itemOpts._bodyProp.numCol = Math.round(itemOpts.columns)
				}
			}
			if (itemOpts.columnSpacing !== undefined) {
				if (typeof itemOpts.columnSpacing !== 'number' || isNaN(itemOpts.columnSpacing) || itemOpts.columnSpacing < 0) {
					console.warn('Warning: text `columnSpacing` must be a number >= 0 (ignoring value)')
				} else {
					itemOpts._bodyProp.spcCol = valToPts(itemOpts.columnSpacing)
				}
			}

			// E: Inset
			// @deprecated 3.10.0 (`inset` - use `margin`)
			if ((itemOpts.inset && !isNaN(Number(itemOpts.inset))) || itemOpts.inset === 0) {
				itemOpts._bodyProp.lIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.rIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.tIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.bIns = inch2Emu(itemOpts.inset)
			}

			// F: Transform @deprecated props
			if (typeof itemOpts.underline === 'boolean' && itemOpts.underline === true) itemOpts.underline = { style: 'sng' }
		}

		// STEP 2: Transform `align`/`valign` to XML values, store in _bodyProp for XML gen
		{
			const align = (itemOpts.align || '').toLowerCase()
			const valign = (itemOpts.valign || '').toLowerCase()
			if (align.startsWith('c')) itemOpts._bodyProp.align = TEXT_HALIGN.center
			else if (align.startsWith('l')) itemOpts._bodyProp.align = TEXT_HALIGN.left
			else if (align.startsWith('r')) itemOpts._bodyProp.align = TEXT_HALIGN.right
			else if (align.startsWith('j')) itemOpts._bodyProp.align = TEXT_HALIGN.justify

			if (valign.startsWith('b')) itemOpts._bodyProp.anchor = TEXT_VALIGN.b
			else if (valign.startsWith('m')) itemOpts._bodyProp.anchor = TEXT_VALIGN.ctr
			else if (valign.startsWith('t')) itemOpts._bodyProp.anchor = TEXT_VALIGN.t
		}

		// STEP 3: ROBUST: Set rational values for some shadow props if needed
		correctShadowOptions(itemOpts.shadow)

		return itemOpts
	}

	// STEP 1: Create/Clean object options
	newObject.options = cleanOpts(objectOptions)

	// STEP 1b: Standalone placeholder type (#1298 - accessibility "Missing Slide Title")
	// `placeholder` is documented as a placeholder *type* ('title', 'body', et. al.). When it
	// resolves to a layout placeholder the layout object supplies the <p:ph> at serialize time,
	// but with a blank/default layout there is no match and no <p:ph> was emitted - so PowerPoint's
	// accessibility checker reports the slide as having no title. Record the type here so a real
	// <p:ph type="..."/> is emitted on the slide shape even without a matching layout placeholder.
	if (!isPlaceholder && newObject.options.placeholder && !newObject.options._placeholderType) {
		newObject.options._placeholderType = newObject.options.placeholder as PLACEHOLDER_TYPE
	}

	// STEP 2: Create/Clean text options
	textObjects.forEach(item => (item.options = cleanOpts(item.options || {})))

	// STEP 3: Create hyperlinks
	createHyperlinkRels(target, textObjects)

	// STEP 4: Create picture-bullet image rels
	createBulletImageRels(target, newObject.options, textObjects)

	// STEP 5: Register an image fill (if any) as a media relationship for serialize-time blipFill
	if (typeof newObject.options.fill === 'object' && (newObject.options.fill.type === 'image' || newObject.options.fill.image)) {
		registerImageFillMedia(target, newObject.options.fill)
	}

	// LAST: Add object to Slide
	target._slideObjects.push(newObject)
}

/**
 * Register slide media relationships for any picture bullets (`bullet.image`) used by a text object.
 * Picture bullets render as `<a:buBlip><a:blip r:embed="rId.."/></a:buBlip>`, so the bullet image
 * needs the same media-rel + package-part plumbing as `addImage()`. The assigned `rId` is stored on
 * the bullet options object (`_rId`) so XML generation can reference it.
 * @param {PresSlideInternal} target - slide receiving the rels
 * @param {ObjectOptions} objectOptions - shape-level text options (bullet may live here)
 * @param {TextProps[]} textObjects - per-paragraph text options (bullet may live here too)
 */
function createBulletImageRels(target: PresSlideInternal, objectOptions: ObjectOptions, textObjects: TextProps[]): void {
	// Collect every bullet options object that requests a picture bullet (shape-level + per-paragraph).
	// Shape-level bullets are later shared by reference onto the first run, so the same object may appear
	// twice; the `_rId` guard below makes the registration idempotent.
	const bulletObjs: Array<{ image?: { path?: string, data?: string }, _rId?: number, _rIdSvg?: number }> = []
	const collect = (opts?: TextPropsOptions): void => {
		if (opts && typeof opts.bullet === 'object' && opts.bullet) bulletObjs.push(opts.bullet)
	}
	collect(objectOptions)
	textObjects.forEach(item => collect(item.options))

	bulletObjs.forEach(bullet => {
		const img = bullet.image
		if (!img || (!img.path && !img.data)) return

		// REALITY-CHECK: base64 `data` must carry a base64 header (mirror addImage())
		if (img.data && (typeof img.data !== 'string' || !img.data.toLowerCase().includes('base64,'))) {
			console.error('ERROR: bullet.image `data` value lacks a base64 header! Ex: \'image/png;base64,iVBOR[...]\'')
			return
		}

		// Auto-paging clones text objects onto new slides while sharing the bullet options object by
		// reference, so `_rId` may already be set from the originating slide. Skip when this slide already
		// carries the rel; otherwise (re-)register so the new slide's .rels and media part exist.
		if (bullet._rId && target._relsMedia.some(rel => rel.rId === bullet._rId)) return

		// Determine extension: path wins, else sniff the data: mime-type (mirror addImageDefinition())
		let strImgExtn = 'png'
		if (img.path) {
			const imagePathFile = img.path.slice(img.path.lastIndexOf('/') + 1).split('?')[0] || ''
			strImgExtn = ((imagePathFile.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()
		}
		const imageMimeMatch = /image\/(\w+);/.exec(img.data || '')
		if (img.data && imageMimeMatch) strImgExtn = imageMimeMatch[1]
		// `image/svg+xml` does not match the `\w+` sniff above (the `+`), so detect it explicitly (mirror addImageDefinition())
		else if (img.data?.toLowerCase().includes('image/svg+xml')) strImgExtn = 'svg'
		// Path-based SVG sniffing is already handled by the extension parse above.

		const relId = bullet._rId || getNewRelId(target)
		const mediaSlideKey = target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum

		if (strImgExtn === 'svg') {
			// SVG bullets consume *TWO* rels, mirroring addImage(): a PNG preview (referenced by the
			// `<a:buBlip><a:blip r:embed>`) plus the SVG itself (referenced by the `asvg:svgBlip` ext).
			// The preview rel is flagged `isSvgPng` so the media pipeline generates its PNG fallback.
			target._relsMedia.push({
				path: img.path || img.data + 'png',
				type: 'image/png',
				extn: 'png',
				data: img.data || '',
				rId: relId,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.png`,
				isSvgPng: true,
			})
			target._relsMedia.push({
				path: img.path || img.data || 'preencoded.svg',
				type: 'image/svg+xml',
				extn: 'svg',
				data: img.data || '',
				rId: relId + 1,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.svg`,
			})
			bullet._rId = relId
			bullet._rIdSvg = relId + 1
		} else {
			target._relsMedia.push({
				path: img.path || 'preencoded.' + strImgExtn,
				type: imageContentType(strImgExtn),
				extn: strImgExtn,
				data: img.data || '',
				rId: relId,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
			})
			bullet._rId = relId
		}
	})
}

/**
 * Adds placeholder objects to slide
 * @param {PresSlideInternal} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: PresSlideInternal): void {
	if (!slide._slideLayout) return
	// Add all placeholders on this Slide that dont already exist
	(slide._slideLayout._slideObjects || []).forEach(slideLayoutObj => {
		if (slideLayoutObj._type === SLIDE_OBJECT_TYPES.placeholder) {
			const slideLayoutOptions = slideLayoutObj.options || {}
			// A: Search for this placeholder on Slide before we add
			// NOTE: Check to ensure a placeholder does not already exist on the Slide
			// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
			if (!slide._slideObjects.some(slideObj => slideObj.options && slideObj.options.placeholder === slideLayoutOptions.placeholder)) {
				addTextDefinition(slide, [{ text: '' }], slideLayoutOptions, true)
			}
		}
	})
}

/* -------------------------------------------------------------------------------- */

/**
 * Adds a background image or color to a slide definition.
 * @param {BackgroundProps} props - color string or an object with image definition
 * @param {PresSlideInternal} target - slide object that the background is set to
 */
export function addBackgroundDefinition(props: BackgroundProps, target: SlideLayoutInternal): void {
	// A: @deprecated
	if (target.bkgd) {
		if (!target.background) target.background = {}

		if (typeof target.bkgd === 'string') target.background.color = target.bkgd
		else {
			if (target.bkgd.data) target.background.data = target.bkgd.data
			if (target.bkgd.path) target.background.path = target.bkgd.path
			if (target.bkgd.src) target.background.path = target.bkgd.src // @deprecated (drop in 4.x)
		}
	}
	if (target.background?.fill) target.background.color = target.background.fill

	// B: Handle media
	if (props && (props.path || props.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		props.path = props.path || 'preencoded.png'
		let strImgExtn = (props.path.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
		if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		target._relsMedia = target._relsMedia || []
		const intRels = target._relsMedia.length + 1
		// NOTE: `Target` cannot have spaces (eg:"Slide 1-image-1.jpg") or a "presentation is corrupt" warning comes up
		target._relsMedia.push({
			path: props.path,
			type: imageContentType(strImgExtn),
			extn: strImgExtn,
			data: props.data || undefined,
			rId: intRels,
			Target: `../media/${(target._name || '').replace(/\s+/gi, '-')}-image-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		target._bkgdImgRid = intRels
	}
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {PresSlideInternal} target - slide object that any hyperlinks will be be added to
 * @param {number | string | TextProps | TextProps[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(
	target: PresSlideInternal,
	text: number | string | ISlideObject | TextProps | TextProps[] | TableCell[] | TableCell[][],
	options?: TextPropsOptions[],
): void {
	let textObjs: Array<HyperlinkTextObject | TableCell[]> = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	else if (typeof text === 'object') textObjs = [text as HyperlinkTextObject]

	textObjs.forEach((text: HyperlinkTextObject | TableCell[], idx: number) => {
		// NOTE: `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
		if (Array.isArray(text)) {
			const cellOpts: TextPropsOptions[] = []
			text.forEach((tablecell) => {
				if (tablecell.options) {
					cellOpts.push(tablecell.options)
				}
			})
			createHyperlinkRels(target, text, cellOpts)
			return
		}

		// IMPORTANT: `options` are lost due to recursion/copy!
		if (options && options[idx] && options[idx].hyperlink) text.options = { ...text.options, ...options[idx] }
		if (Array.isArray(text.text)) {
			createHyperlinkRels(target, text.text, options && options[idx] ? [options[idx]] : undefined)
		} else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink._rId) {
			const hyperlink = text.options.hyperlink
			if (typeof hyperlink !== 'object') {
				console.log('ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:\'https://github.com\'}` ')
			}
			else if (!hyperlink.url && !hyperlink.slide) {
				console.log('ERROR: \'hyperlink requires either: `url` or `slide`\'')
			}
			else {
				const relId = getNewRelId(target)

				target._rels.push({
					type: SLIDE_OBJECT_TYPES.hyperlink,
					data: hyperlink.slide ? 'slide' : 'dummy',
					rId: relId,
					Target: hyperlink.url ? encodeXmlEntities(hyperlink.url) : String(hyperlink.slide),
				})

				hyperlink._rId = relId
			}
		}
		else if (text && typeof text === 'object' && text.options && text.options.hyperlink && text.options.hyperlink._rId) {
			const hyperlink = text.options.hyperlink
			const hyperlinkRelId = hyperlink._rId
			// NOTE: auto-paging will create new slides, but skip above as _rId exists, BUT this is a new slide, so add rels!
			if (hyperlinkRelId && !target._rels.some(rel => rel.rId === hyperlinkRelId)) {
				target._rels.push({
					type: SLIDE_OBJECT_TYPES.hyperlink,
					data: hyperlink.slide ? 'slide' : 'dummy',
					rId: hyperlinkRelId,
					Target: hyperlink.url ? encodeXmlEntities(hyperlink.url) : String(hyperlink.slide),
				})
			}
		}
	})
}
