/**
 * PptxGenJS Enums
 * NOTE: `enum` wont work for objects, so use `Object.freeze`
 */

export type ZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
export type WRITE_OUTPUT_TYPE = ZIP_OUTPUT_TYPE | 'STREAM'
/**
 * Public chart-type name accepted by `addChart()`. Derived from the internal
 * `ChartType` enum (see below) so the two never drift: adding a member to the
 * enum extends this union automatically.
 */
export type CHART_NAME = `${ChartType}`
export type SCHEME_COLORS =
	'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'

export enum TextAnchor {
	b = 'b',
	ctr = 'ctr',
	t = 't',
}

// ENUM
export enum OutputType {
	arraybuffer = 'arraybuffer',
	base64 = 'base64',
	binarystring = 'binarystring',
	blob = 'blob',
	nodebuffer = 'nodebuffer',
	uint8array = 'uint8array',
}
export enum ChartType {
	area = 'area',
	bar = 'bar',
	bar3d = 'bar3D',
	bubble = 'bubble',
	bubble3d = 'bubble3D',
	doughnut = 'doughnut',
	line = 'line',
	pie = 'pie',
	radar = 'radar',
	scatter = 'scatter',
	waterfall = 'waterfall',
	funnel = 'funnel',
	treemap = 'treemap',
	sunburst = 'sunburst',
	histogram = 'histogram',
	pareto = 'pareto',
}
export enum ShapeType {
	accentBorderCallout1 = 'accentBorderCallout1',
	accentBorderCallout2 = 'accentBorderCallout2',
	accentBorderCallout3 = 'accentBorderCallout3',
	accentCallout1 = 'accentCallout1',
	accentCallout2 = 'accentCallout2',
	accentCallout3 = 'accentCallout3',
	actionButtonBackPrevious = 'actionButtonBackPrevious',
	actionButtonBeginning = 'actionButtonBeginning',
	actionButtonBlank = 'actionButtonBlank',
	actionButtonDocument = 'actionButtonDocument',
	actionButtonEnd = 'actionButtonEnd',
	actionButtonForwardNext = 'actionButtonForwardNext',
	actionButtonHelp = 'actionButtonHelp',
	actionButtonHome = 'actionButtonHome',
	actionButtonInformation = 'actionButtonInformation',
	actionButtonMovie = 'actionButtonMovie',
	actionButtonReturn = 'actionButtonReturn',
	actionButtonSound = 'actionButtonSound',
	arc = 'arc',
	bentArrow = 'bentArrow',
	bentUpArrow = 'bentUpArrow',
	bevel = 'bevel',
	blockArc = 'blockArc',
	borderCallout1 = 'borderCallout1',
	borderCallout2 = 'borderCallout2',
	borderCallout3 = 'borderCallout3',
	bracePair = 'bracePair',
	bracketPair = 'bracketPair',
	callout1 = 'callout1',
	callout2 = 'callout2',
	callout3 = 'callout3',
	can = 'can',
	chartPlus = 'chartPlus',
	chartStar = 'chartStar',
	chartX = 'chartX',
	chevron = 'chevron',
	chord = 'chord',
	circularArrow = 'circularArrow',
	cloud = 'cloud',
	cloudCallout = 'cloudCallout',
	corner = 'corner',
	cornerTabs = 'cornerTabs',
	cube = 'cube',
	curvedDownArrow = 'curvedDownArrow',
	curvedLeftArrow = 'curvedLeftArrow',
	curvedRightArrow = 'curvedRightArrow',
	curvedUpArrow = 'curvedUpArrow',
	custGeom = 'custGeom',
	decagon = 'decagon',
	diagStripe = 'diagStripe',
	diamond = 'diamond',
	dodecagon = 'dodecagon',
	donut = 'donut',
	doubleWave = 'doubleWave',
	downArrow = 'downArrow',
	downArrowCallout = 'downArrowCallout',
	ellipse = 'ellipse',
	ellipseRibbon = 'ellipseRibbon',
	ellipseRibbon2 = 'ellipseRibbon2',
	flowChartAlternateProcess = 'flowChartAlternateProcess',
	flowChartCollate = 'flowChartCollate',
	flowChartConnector = 'flowChartConnector',
	flowChartDecision = 'flowChartDecision',
	flowChartDelay = 'flowChartDelay',
	flowChartDisplay = 'flowChartDisplay',
	flowChartDocument = 'flowChartDocument',
	flowChartExtract = 'flowChartExtract',
	flowChartInputOutput = 'flowChartInputOutput',
	flowChartInternalStorage = 'flowChartInternalStorage',
	flowChartMagneticDisk = 'flowChartMagneticDisk',
	flowChartMagneticDrum = 'flowChartMagneticDrum',
	flowChartMagneticTape = 'flowChartMagneticTape',
	flowChartManualInput = 'flowChartManualInput',
	flowChartManualOperation = 'flowChartManualOperation',
	flowChartMerge = 'flowChartMerge',
	flowChartMultidocument = 'flowChartMultidocument',
	flowChartOfflineStorage = 'flowChartOfflineStorage',
	flowChartOffpageConnector = 'flowChartOffpageConnector',
	flowChartOnlineStorage = 'flowChartOnlineStorage',
	flowChartOr = 'flowChartOr',
	flowChartPredefinedProcess = 'flowChartPredefinedProcess',
	flowChartPreparation = 'flowChartPreparation',
	flowChartProcess = 'flowChartProcess',
	flowChartPunchedCard = 'flowChartPunchedCard',
	flowChartPunchedTape = 'flowChartPunchedTape',
	flowChartSort = 'flowChartSort',
	flowChartSummingJunction = 'flowChartSummingJunction',
	flowChartTerminator = 'flowChartTerminator',
	foldedCorner = 'foldedCorner',
	frame = 'frame',
	funnel = 'funnel',
	gear6 = 'gear6',
	gear9 = 'gear9',
	halfFrame = 'halfFrame',
	heart = 'heart',
	heptagon = 'heptagon',
	hexagon = 'hexagon',
	homePlate = 'homePlate',
	horizontalScroll = 'horizontalScroll',
	irregularSeal1 = 'irregularSeal1',
	irregularSeal2 = 'irregularSeal2',
	leftArrow = 'leftArrow',
	leftArrowCallout = 'leftArrowCallout',
	leftBrace = 'leftBrace',
	leftBracket = 'leftBracket',
	leftCircularArrow = 'leftCircularArrow',
	leftRightArrow = 'leftRightArrow',
	leftRightArrowCallout = 'leftRightArrowCallout',
	leftRightCircularArrow = 'leftRightCircularArrow',
	leftRightRibbon = 'leftRightRibbon',
	leftRightUpArrow = 'leftRightUpArrow',
	leftUpArrow = 'leftUpArrow',
	lightningBolt = 'lightningBolt',
	line = 'line',
	lineInv = 'lineInv',
	mathDivide = 'mathDivide',
	mathEqual = 'mathEqual',
	mathMinus = 'mathMinus',
	mathMultiply = 'mathMultiply',
	mathNotEqual = 'mathNotEqual',
	mathPlus = 'mathPlus',
	moon = 'moon',
	noSmoking = 'noSmoking',
	nonIsoscelesTrapezoid = 'nonIsoscelesTrapezoid',
	notchedRightArrow = 'notchedRightArrow',
	octagon = 'octagon',
	parallelogram = 'parallelogram',
	pentagon = 'pentagon',
	pie = 'pie',
	pieWedge = 'pieWedge',
	plaque = 'plaque',
	plaqueTabs = 'plaqueTabs',
	plus = 'plus',
	quadArrow = 'quadArrow',
	quadArrowCallout = 'quadArrowCallout',
	rect = 'rect',
	ribbon = 'ribbon',
	ribbon2 = 'ribbon2',
	rightArrow = 'rightArrow',
	rightArrowCallout = 'rightArrowCallout',
	rightBrace = 'rightBrace',
	rightBracket = 'rightBracket',
	round1Rect = 'round1Rect',
	round2DiagRect = 'round2DiagRect',
	round2SameRect = 'round2SameRect',
	roundRect = 'roundRect',
	rtTriangle = 'rtTriangle',
	smileyFace = 'smileyFace',
	snip1Rect = 'snip1Rect',
	snip2DiagRect = 'snip2DiagRect',
	snip2SameRect = 'snip2SameRect',
	snipRoundRect = 'snipRoundRect',
	squareTabs = 'squareTabs',
	star10 = 'star10',
	star12 = 'star12',
	star16 = 'star16',
	star24 = 'star24',
	star32 = 'star32',
	star4 = 'star4',
	star5 = 'star5',
	star6 = 'star6',
	star7 = 'star7',
	star8 = 'star8',
	stripedRightArrow = 'stripedRightArrow',
	sun = 'sun',
	swooshArrow = 'swooshArrow',
	teardrop = 'teardrop',
	trapezoid = 'trapezoid',
	triangle = 'triangle',
	upArrow = 'upArrow',
	upArrowCallout = 'upArrowCallout',
	upDownArrow = 'upDownArrow',
	upDownArrowCallout = 'upDownArrowCallout',
	uturnArrow = 'uturnArrow',
	verticalScroll = 'verticalScroll',
	wave = 'wave',
	wedgeEllipseCallout = 'wedgeEllipseCallout',
	wedgeRectCallout = 'wedgeRectCallout',
	wedgeRoundRectCallout = 'wedgeRoundRectCallout',
}
export enum SchemeColor {
	text1 = 'tx1',
	text2 = 'tx2',
	background1 = 'bg1',
	background2 = 'bg2',
	accent1 = 'accent1',
	accent2 = 'accent2',
	accent3 = 'accent3',
	accent4 = 'accent4',
	accent5 = 'accent5',
	accent6 = 'accent6',
}
export enum AlignH {
	left = 'left',
	center = 'center',
	right = 'right',
	justify = 'justify',
}
export enum AlignV {
	top = 'top',
	middle = 'middle',
	bottom = 'bottom',
}

export type SHAPE_NAME =
	| 'accentBorderCallout1'
	| 'accentBorderCallout2'
	| 'accentBorderCallout3'
	| 'accentCallout1'
	| 'accentCallout2'
	| 'accentCallout3'
	| 'actionButtonBackPrevious'
	| 'actionButtonBeginning'
	| 'actionButtonBlank'
	| 'actionButtonDocument'
	| 'actionButtonEnd'
	| 'actionButtonForwardNext'
	| 'actionButtonHelp'
	| 'actionButtonHome'
	| 'actionButtonInformation'
	| 'actionButtonMovie'
	| 'actionButtonReturn'
	| 'actionButtonSound'
	| 'arc'
	| 'bentArrow'
	| 'bentUpArrow'
	| 'bevel'
	| 'blockArc'
	| 'borderCallout1'
	| 'borderCallout2'
	| 'borderCallout3'
	| 'bracePair'
	| 'bracketPair'
	| 'callout1'
	| 'callout2'
	| 'callout3'
	| 'can'
	| 'chartPlus'
	| 'chartStar'
	| 'chartX'
	| 'chevron'
	| 'chord'
	| 'circularArrow'
	| 'cloud'
	| 'cloudCallout'
	| 'corner'
	| 'cornerTabs'
	| 'cube'
	| 'curvedDownArrow'
	| 'curvedLeftArrow'
	| 'curvedRightArrow'
	| 'curvedUpArrow'
	| 'custGeom'
	| 'decagon'
	| 'diagStripe'
	| 'diamond'
	| 'dodecagon'
	| 'donut'
	| 'doubleWave'
	| 'downArrow'
	| 'downArrowCallout'
	| 'ellipse'
	| 'ellipseRibbon'
	| 'ellipseRibbon2'
	| 'flowChartAlternateProcess'
	| 'flowChartCollate'
	| 'flowChartConnector'
	| 'flowChartDecision'
	| 'flowChartDelay'
	| 'flowChartDisplay'
	| 'flowChartDocument'
	| 'flowChartExtract'
	| 'flowChartInputOutput'
	| 'flowChartInternalStorage'
	| 'flowChartMagneticDisk'
	| 'flowChartMagneticDrum'
	| 'flowChartMagneticTape'
	| 'flowChartManualInput'
	| 'flowChartManualOperation'
	| 'flowChartMerge'
	| 'flowChartMultidocument'
	| 'flowChartOfflineStorage'
	| 'flowChartOffpageConnector'
	| 'flowChartOnlineStorage'
	| 'flowChartOr'
	| 'flowChartPredefinedProcess'
	| 'flowChartPreparation'
	| 'flowChartProcess'
	| 'flowChartPunchedCard'
	| 'flowChartPunchedTape'
	| 'flowChartSort'
	| 'flowChartSummingJunction'
	| 'flowChartTerminator'
	| 'foldedCorner'
	| 'frame'
	| 'funnel'
	| 'gear6'
	| 'gear9'
	| 'halfFrame'
	| 'heart'
	| 'heptagon'
	| 'hexagon'
	| 'homePlate'
	| 'horizontalScroll'
	| 'irregularSeal1'
	| 'irregularSeal2'
	| 'leftArrow'
	| 'leftArrowCallout'
	| 'leftBrace'
	| 'leftBracket'
	| 'leftCircularArrow'
	| 'leftRightArrow'
	| 'leftRightArrowCallout'
	| 'leftRightCircularArrow'
	| 'leftRightRibbon'
	| 'leftRightUpArrow'
	| 'leftUpArrow'
	| 'lightningBolt'
	| 'line'
	| 'lineInv'
	| 'mathDivide'
	| 'mathEqual'
	| 'mathMinus'
	| 'mathMultiply'
	| 'mathNotEqual'
	| 'mathPlus'
	| 'moon'
	| 'noSmoking'
	| 'nonIsoscelesTrapezoid'
	| 'notchedRightArrow'
	| 'octagon'
	| 'parallelogram'
	| 'pentagon'
	| 'pie'
	| 'pieWedge'
	| 'plaque'
	| 'plaqueTabs'
	| 'plus'
	| 'quadArrow'
	| 'quadArrowCallout'
	| 'rect'
	| 'ribbon'
	| 'ribbon2'
	| 'rightArrow'
	| 'rightArrowCallout'
	| 'rightBrace'
	| 'rightBracket'
	| 'round1Rect'
	| 'round2DiagRect'
	| 'round2SameRect'
	| 'roundRect'
	| 'rtTriangle'
	| 'smileyFace'
	| 'snip1Rect'
	| 'snip2DiagRect'
	| 'snip2SameRect'
	| 'snipRoundRect'
	| 'squareTabs'
	| 'star10'
	| 'star12'
	| 'star16'
	| 'star24'
	| 'star32'
	| 'star4'
	| 'star5'
	| 'star6'
	| 'star7'
	| 'star8'
	| 'stripedRightArrow'
	| 'sun'
	| 'swooshArrow'
	| 'teardrop'
	| 'trapezoid'
	| 'triangle'
	| 'upArrow'
	| 'upArrowCallout'
	| 'upDownArrow'
	| 'upDownArrowCallout'
	| 'uturnArrow'
	| 'verticalScroll'
	| 'wave'
	| 'wedgeEllipseCallout'
	| 'wedgeRectCallout'
	| 'wedgeRoundRectCallout'
	| CONNECTOR_PRESET_NAME

/**
 * Valid ECMA-376 `ST_ShapeType` presets that are not surfaced with a friendly
 * `ShapeType` name. They are still legal geometries PowerPoint renders, so the
 * preset-validation set must accept them.
 */
const EXTRA_SHAPE_PRESETS = [
	'straightConnector1',
	'bentConnector2',
	'bentConnector3',
	'bentConnector4',
	'bentConnector5',
	'curvedConnector2',
	'curvedConnector3',
	'curvedConnector4',
	'curvedConnector5',
] as const

/**
 * Connector preset geometries (`straightConnector1` / `bentConnector{2-5}` /
 * `curvedConnector{2-5}`) that `addShape` can serialize but that are not surfaced
 * as friendly `ShapeType` names. Derived from `EXTRA_SHAPE_PRESETS` so the public
 * `SHAPE_NAME` union and the runtime `VALID_SHAPE_PRESETS` set share one source of
 * truth and cannot drift. Use `addConnector` for live, endpoint-bound connectors;
 * pass these to `addShape` for a static box-positioned connector geometry.
 */
export type CONNECTOR_PRESET_NAME = (typeof EXTRA_SHAPE_PRESETS)[number]

/**
 * Every shape geometry name PptxGenJS can serialize without corrupting the
 * package: the OOXML preset geometries (`ST_ShapeType` — `ShapeType` values
 * plus the unexposed connectors above) and `custGeom` (freeform paths, emitted
 * as `<a:custGeom>` rather than `<a:prstGeom>`). Used to reject bogus presets
 * before they become an invalid `<a:prstGeom prst="...">` that triggers
 * PowerPoint's "needs repair" dialog and drops the shape.
 */
export const VALID_SHAPE_PRESETS: ReadonlySet<string> = new Set<string>([
	...Object.values(ShapeType),
	...EXTRA_SHAPE_PRESETS,
])

/**
 * Narrow a public `CHART_NAME` to the internal `ChartType` enum. Because
 * `CHART_NAME` is derived from `ChartType`, every `CHART_NAME` value is by
 * construction a valid enum value, so this conversion is total and safe — it is
 * the single boundary cast that lets internal code compare against the enum
 * (enum-to-enum) instead of enum-to-string.
 */
export function asChartType(name: CHART_NAME): ChartType {
	return name as ChartType
}

/**
 * The chartEx (`cx:` / Office 2016 chart-extension) chart types. Unlike the classic 2007 catalog
 * these do NOT emit a `<c:chartSpace>` chart part: they emit a separate `chartEx{N}.xml` part in
 * the `http://schemas.microsoft.com/office/drawing/2014/chartex` namespace, are referenced from
 * the slide through `<mc:AlternateContent>`, and render only in PowerPoint 2016+/Microsoft 365
 * (older Office, Google Slides, Keynote and LibreOffice show the `<mc:Fallback>` shape instead).
 * See `gen/chart/chartex-xml.ts`.
 */
const CHARTEX_TYPES: ReadonlySet<string> = new Set<string>([
	ChartType.waterfall,
	ChartType.funnel,
	ChartType.treemap,
	ChartType.sunburst,
	ChartType.histogram,
	ChartType.pareto,
])

/**
 * Is `type` a chartEx (cx:) chart? Accepts a `ChartType`/`CHART_NAME` string; anything else —
 * including a combo-chart `ChartMulti[]` array or `undefined` — is not a chartEx type (combos are
 * classic-only), so those return `false`.
 */
export function isChartExType(type: unknown): boolean {
	return typeof type === 'string' && CHARTEX_TYPES.has(type)
}

export enum SlideObjectType {
	chart = 'chart',
	connector = 'connector',
	group = 'group',
	hyperlink = 'hyperlink',
	image = 'image',
	media = 'media',
	online = 'online',
	placeholder = 'placeholder',
	table = 'table',
	tablecell = 'tablecell',
	text = 'text',
	notes = 'notes',
}

/**
 * Maps a friendly `ConnectorType` to its OOXML connector preset *family*. `straight`
 * has a single fixed preset; `elbow` and `curved` select a member by bend count
 * (`bentConnector{3,4,5}` / `curvedConnector{3,4,5}`) — see `connectorPresetFor`.
 */
export const CONNECTOR_PRESETS = {
	straight: 'straightConnector1',
	elbow: 'bentConnector',
	curved: 'curvedConnector',
} as const satisfies Record<'straight' | 'elbow' | 'curved', string>

/**
 * Resolve the concrete connector preset geometry name for a `type` + bend count.
 * `straight` ignores `bends` (it has no adjustable jog). For `elbow` / `curved`, the
 * preset suffix is `bends + 2`, so `bends` of 1/2/3 maps to `…Connector3/4/5`, each
 * of which exposes exactly `bends` adjustable guides (`adj1…adjN`).
 * @param {'straight' | 'elbow' | 'curved'} type - friendly connector routing style
 * @param {number} bends - number of adjustable bends (1–3); ignored for `straight`
 * @return {CONNECTOR_PRESET_NAME} OOXML preset geometry name (e.g. `bentConnector4`)
 */
export function connectorPresetFor(type: 'straight' | 'elbow' | 'curved', bends: number): CONNECTOR_PRESET_NAME {
	if (type === 'straight') return CONNECTOR_PRESETS.straight
	return `${CONNECTOR_PRESETS[type]}${bends + 2}` as CONNECTOR_PRESET_NAME
}
export enum PlaceholderType {
	title = 'title',
	body = 'body',
	image = 'pic',
	chart = 'chart',
	table = 'tbl',
	media = 'media',
}
export type PLACEHOLDER_TYPE = 'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media'

/**
 * Bullet glyph presets. Only `DEFAULT` is currently wired through to output; the remaining entries are
 * reserved and not yet selectable via options.
 */
export enum BulletType {
	DEFAULT = '&#x2022;',
	CHECK = '&#x2713;',
	STAR = '&#x2605;',
	TRIANGLE = '&#x25B6;',
}

/**
 * Built-in PowerPoint table style IDs.
 *
 * Each value is the GUID that goes into `<a:tableStyleId>` inside `<a:tblPr>`.
 * These 74 GUIDs are the complete set defined in [MS-OE376] §5.1.6.10 and are
 * resolved by PowerPoint without embedding any style definition in the file.
 *
 * Only these enum members are supported. Raw GUID strings are not accepted;
 * use `TableStyle` members exclusively.
 */
export enum TableStyle {
	// ── No style ────────────────────────────────────────────────────────────
	NO_STYLE_NO_GRID = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}',
	NO_STYLE_TABLE_GRID = '{5940675A-B579-460E-94D1-54222C63F5DA}',

	// ── Themed Style 1 ──────────────────────────────────────────────────────
	THEMED_STYLE_1_ACCENT_1 = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}',
	THEMED_STYLE_1_ACCENT_2 = '{284E427A-3D55-4303-BF80-6455036E1DE7}',
	THEMED_STYLE_1_ACCENT_3 = '{69C7853C-536D-4A76-A0AE-DD22124D55A5}',
	THEMED_STYLE_1_ACCENT_4 = '{775DCB02-9BB8-47FD-8907-85C794F793BA}',
	THEMED_STYLE_1_ACCENT_5 = '{35758FB7-9AC5-4552-8A53-C91805E547FA}',
	THEMED_STYLE_1_ACCENT_6 = '{08FB837D-C827-4EFA-A057-4D05807E0F7C}',

	// ── Themed Style 2 ──────────────────────────────────────────────────────
	THEMED_STYLE_2_ACCENT_1 = '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}',
	THEMED_STYLE_2_ACCENT_2 = '{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}',
	THEMED_STYLE_2_ACCENT_3 = '{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}',
	THEMED_STYLE_2_ACCENT_4 = '{E269D01E-BC32-4049-B463-5C60D7B0CCD2}',
	THEMED_STYLE_2_ACCENT_5 = '{327F97BB-C833-4FB7-BDE5-3F7075034690}',
	THEMED_STYLE_2_ACCENT_6 = '{638B1855-1B75-4FBE-930C-398BA8C253C6}',

	// ── Light Style 1 ───────────────────────────────────────────────────────
	LIGHT_STYLE_1 = '{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}',
	LIGHT_STYLE_1_ACCENT_1 = '{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}',
	LIGHT_STYLE_1_ACCENT_2 = '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
	LIGHT_STYLE_1_ACCENT_3 = '{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}',
	LIGHT_STYLE_1_ACCENT_4 = '{D27102A9-8310-4765-A935-A1911B00CA55}',
	LIGHT_STYLE_1_ACCENT_5 = '{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}',
	LIGHT_STYLE_1_ACCENT_6 = '{68D230F3-CF80-4859-8CE7-A43EE81993B5}',

	// ── Light Style 2 ───────────────────────────────────────────────────────
	LIGHT_STYLE_2 = '{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}',
	LIGHT_STYLE_2_ACCENT_1 = '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}',
	LIGHT_STYLE_2_ACCENT_2 = '{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}',
	LIGHT_STYLE_2_ACCENT_3 = '{F2DE63D5-997A-4646-A377-4702673A728D}',
	LIGHT_STYLE_2_ACCENT_4 = '{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}',
	LIGHT_STYLE_2_ACCENT_5 = '{5A111915-BE36-4E01-A7E5-04B1672EAD32}',
	LIGHT_STYLE_2_ACCENT_6 = '{912C8C85-51F0-491E-9774-3900AFEF0FD7}',

	// ── Light Style 3 ───────────────────────────────────────────────────────
	LIGHT_STYLE_3 = '{616DA210-FB5B-4158-B5E0-FEB733F419BA}',
	LIGHT_STYLE_3_ACCENT_1 = '{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}',
	LIGHT_STYLE_3_ACCENT_2 = '{5DA37D80-6434-44D0-A028-1B22A696006F}',
	LIGHT_STYLE_3_ACCENT_3 = '{8799B23B-EC83-4686-B30A-512413B5E67A}',
	LIGHT_STYLE_3_ACCENT_4 = '{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}',
	LIGHT_STYLE_3_ACCENT_5 = '{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}',
	LIGHT_STYLE_3_ACCENT_6 = '{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}',

	// ── Medium Style 1 ──────────────────────────────────────────────────────
	MEDIUM_STYLE_1 = '{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}',
	MEDIUM_STYLE_1_ACCENT_1 = '{B301B821-A1FF-4177-AEE7-76D212191A09}',
	MEDIUM_STYLE_1_ACCENT_2 = '{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}',
	MEDIUM_STYLE_1_ACCENT_3 = '{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}',
	MEDIUM_STYLE_1_ACCENT_4 = '{1E171933-4619-4E11-9A3F-F7608DF75F80}',
	MEDIUM_STYLE_1_ACCENT_5 = '{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}',
	MEDIUM_STYLE_1_ACCENT_6 = '{10A1B5D5-9B99-4C35-A422-299274C87663}',

	// ── Medium Style 2 ──────────────────────────────────────────────────────
	MEDIUM_STYLE_2 = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
	MEDIUM_STYLE_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
	MEDIUM_STYLE_2_ACCENT_2 = '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}',
	MEDIUM_STYLE_2_ACCENT_3 = '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}',
	MEDIUM_STYLE_2_ACCENT_4 = '{00A15C55-8517-42AA-B614-E9B94910E393}',
	MEDIUM_STYLE_2_ACCENT_5 = '{7DF18680-E054-41AD-8BC1-D1AEF772440D}',
	MEDIUM_STYLE_2_ACCENT_6 = '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}',

	// ── Medium Style 3 ──────────────────────────────────────────────────────
	MEDIUM_STYLE_3 = '{8EC20E35-A176-4012-BC5E-935CFFF8708E}',
	MEDIUM_STYLE_3_ACCENT_1 = '{6E25E649-3F16-4E02-A733-19D2CDBF48F0}',
	MEDIUM_STYLE_3_ACCENT_2 = '{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}',
	MEDIUM_STYLE_3_ACCENT_3 = '{EB344D84-9AFB-497E-A393-DC336BA19D2E}',
	MEDIUM_STYLE_3_ACCENT_4 = '{EB9631B5-78F2-41C9-869B-9F39066F8104}',
	MEDIUM_STYLE_3_ACCENT_5 = '{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}',
	MEDIUM_STYLE_3_ACCENT_6 = '{2A488322-F2BA-4B5B-9748-0D474271808F}',

	// ── Medium Style 4 ──────────────────────────────────────────────────────
	MEDIUM_STYLE_4 = '{D7AC3CCA-C797-4891-BE02-D94E43425B78}',
	MEDIUM_STYLE_4_ACCENT_1 = '{69CF1AB2-1976-4502-BF36-3FF5EA218861}',
	MEDIUM_STYLE_4_ACCENT_2 = '{8A107856-5554-42FB-B03E-39F5DBC370BA}',
	MEDIUM_STYLE_4_ACCENT_3 = '{0505E3EF-67EA-436B-97B2-0124C06EBD24}',
	MEDIUM_STYLE_4_ACCENT_4 = '{C4B1156A-380E-4F78-BDF5-A606A8083BF9}',
	MEDIUM_STYLE_4_ACCENT_5 = '{22838BEF-8BB2-4498-84A7-C5851F593DF1}',
	MEDIUM_STYLE_4_ACCENT_6 = '{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}',

	// ── Dark Style 1 ────────────────────────────────────────────────────────
	DARK_STYLE_1 = '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}',
	DARK_STYLE_1_ACCENT_1 = '{125E5076-3810-47DD-B79F-674D7AD40C01}',
	DARK_STYLE_1_ACCENT_2 = '{37CE84F3-28C3-443E-9E96-99CF82512B78}',
	DARK_STYLE_1_ACCENT_3 = '{D03447BB-5D67-496B-8E87-E561075AD55C}',
	DARK_STYLE_1_ACCENT_4 = '{E929F9F4-4A8F-4326-A1B4-22849713DDAB}',
	DARK_STYLE_1_ACCENT_5 = '{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}',
	DARK_STYLE_1_ACCENT_6 = '{AF606853-7671-496A-8E4F-DF71F8EC918B}',

	// ── Dark Style 2 ────────────────────────────────────────────────────────
	DARK_STYLE_2 = '{5202B0CA-FC54-4496-8BCA-5EF66A818D29}',
	/** Header uses Accent 2, body rows use Accent 1 */
	DARK_STYLE_2_ACCENT_1_ACCENT_2 = '{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}',
	/** Header uses Accent 4, body rows use Accent 3 */
	DARK_STYLE_2_ACCENT_3_ACCENT_4 = '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}',
	/** Header uses Accent 6, body rows use Accent 5 */
	DARK_STYLE_2_ACCENT_5_ACCENT_6 = '{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}',
}
