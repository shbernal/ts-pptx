/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * PptxGenJS Interfaces
 */

import type { CHART_NAME, PLACEHOLDER_TYPE, SHAPE_NAME, SLIDE_OBJECT_TYPES, TABLE_STYLE, TEXT_HALIGN, TEXT_VALIGN, WRITE_OUTPUT_TYPE } from './core-enums.js'

// Core Types
// ==========

/**
 * Coordinate value. A bare `number` is **always inches** — there is no magnitude-based unit
 * guessing. For other units use an explicit string suffix:
 * - `number` → inches (e.g. `10.25`)
 * - `"<n>%"` → percentage of the slide axis (e.g. `"75%"`)
 * - `"<n>in"` → inches (e.g. `"10.25in"`)
 * - `"<n>pt"` → points (e.g. `"72pt"` = 1 inch)
 * - `"<n>px"` → CSS reference pixels at 96 px/inch (e.g. `"96px"` = 1 inch), for authoring against
 *   a known web/source canvas size
 * - `"<n>emu"` → raw EMU, the escape hatch for exact OOXML units (e.g. `"914400emu"` = 1 inch)
 *
 * @example 10.25 // inches
 * @example '75%' // percentage of slide size
 * @example '72pt' // points
 * @example '960px' // pixels at 96 DPI = 10 inches
 * @example '914400emu' // raw EMU
 */
export type Coord = number | `${number}%` | `${number}in` | `${number}pt` | `${number}px` | `${number}emu`
export interface PositionProps {
	/**
	 * Horizontal position
	 * - inches or percentage
	 * @example 10.25 // position in inches
	 * @example '75%' // position as percentage of slide size
	 */
	x?: Coord
	/**
	 * Vertical position
	 * - inches or percentage
	 * @example 10.25 // position in inches
	 * @example '75%' // position as percentage of slide size
	 */
	y?: Coord
	/**
	 * Height
	 * - inches or percentage
	 * @example 10.25 // height in inches
	 * @example '75%' // height as percentage of slide size
	 */
	h?: Coord
	/**
	 * Width
	 * - inches or percentage
	 * @example 10.25 // width in inches
	 * @example '75%' // width as percentage of slide size
	 */
	w?: Coord
}
/**
 * Reusable optional data/path fields.
 * Use `DataOrPathRequiredProps` for APIs that require at least one source.
 */
export interface DataOrPathProps {
	/**
	 * URL or relative path
	 *
	 * @example 'https://onedrives.com/myimg.png` // retrieve image via URL
	 * @example '/home/gitbrent/images/myimg.png` // retrieve image via local path
	 */
	path?: string
	/**
	 * base64-encoded string
	 * - Useful for avoiding potential path/server issues
	 *
	 * @example 'image/png;base64,iVtDafDrBF[...]=' // pre-encoded image in base-64
	 */
	data?: string
}
export type DataOrPathRequiredProps =
	| (DataOrPathProps & { data: string })
	| (DataOrPathProps & { path: string })
export interface BackgroundProps extends DataOrPathProps, ShapeFillProps {
	/**
	 * Color (hex format)
	 * @deprecated v3.6.0 - use `ShapeFillProps` instead
	 */
	fill?: HexColor

	/**
	 * source URL
	 * @deprecated v3.6.0 - use `DataOrPathProps` instead - remove in v4.0.0
	 */
	src?: string
}
/**
 * Color in Hex format
 * @example 'FF3399'
 */
export type HexColor = string
export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
export type Color = HexColor | ThemeColor
export interface GradientStopProps {
	/**
	 * Stop position as a percentage.
	 * - range: 0-100
	 */
	position: number
	/**
	 * Stop color.
	 * - `HexColor` or `ThemeColor`
	 */
	color: Color
	/**
	 * Transparency (percent)
	 * - range: 0-100
	 * @default 0
	 */
	transparency?: number
	/**
	 * Transparency (percent)
	 * @deprecated v3.3.0 - use `transparency`
	 */
	alpha?: number
}
export interface LinearGradientFillProps {
	/**
	 * Gradient type.
	 */
	kind: 'linear'
	/**
	 * Gradient angle in degrees. Values are normalized into 0-359.999...
	 * @default 0
	 */
	angle?: number
	/**
	 * Whether the gradient angle scales with the fill region.
	 */
	scaled?: boolean
	/**
	 * Whether the fill rotates with the shape.
	 * @default true
	 */
	rotateWithShape?: boolean
	/**
	 * Gradient stops. Stops are serialized in ascending `position` order.
	 */
	stops: GradientStopProps[]
}
export interface RadialGradientFillProps {
	/**
	 * Gradient type. A circular gradient radiating from a focus point: the stop at
	 * `position: 0` sits at the center and later stops fan outward to the edges.
	 */
	kind: 'radial'
	/**
	 * Focus point of the radial gradient as percentages of the fill box, where
	 * `{ x: 50, y: 50 }` (the default) centers it. Lower/higher values push the
	 * bright center toward an edge.
	 * @default { x: 50, y: 50 }
	 */
	center?: { x: number, y: number }
	/**
	 * Whether the fill rotates with the shape.
	 * @default true
	 */
	rotateWithShape?: boolean
	/**
	 * Gradient stops. Stops are serialized in ascending `position` order; the
	 * first (`position: 0`) is the center color.
	 */
	stops: GradientStopProps[]
}
export type GradientFillProps = LinearGradientFillProps | RadialGradientFillProps

/** OOXML ST_PresetPatternVal — preset pattern names for `<a:pattFill prst="...">` */
export type PatternPreset =
	| 'pct5' | 'pct10' | 'pct20' | 'pct25' | 'pct30' | 'pct40' | 'pct50'
	| 'pct60' | 'pct70' | 'pct75' | 'pct80' | 'pct90'
	| 'horz' | 'vert' | 'ltHorz' | 'ltVert' | 'dkHorz' | 'dkVert'
	| 'narHorz' | 'narVert' | 'dashHorz' | 'dashVert'
	| 'cross' | 'dnDiag' | 'upDiag' | 'ltDnDiag' | 'ltUpDiag'
	| 'dkDnDiag' | 'dkUpDiag' | 'wdDnDiag' | 'wdUpDiag'
	| 'dashDnDiag' | 'dashUpDiag' | 'diagCross'
	| 'smCheck' | 'lgCheck' | 'smGrid' | 'lgGrid' | 'dotGrid'
	| 'smConfetti' | 'lgConfetti' | 'horzBrick' | 'diagBrick'
	| 'solidDmnd' | 'openDmnd' | 'dotDmnd' | 'plaid' | 'sphere'
	| 'weave' | 'divot' | 'shingle' | 'wave' | 'trellis' | 'zigZag'

export interface PatternFillProps {
	/** OOXML preset pattern (`prst` attribute on `<a:pattFill>`). */
	preset: PatternPreset
	/** Foreground color. Defaults to black (`000000`) if omitted. */
	fgColor?: Color
	/** Background color. Defaults to white (`FFFFFF`) if omitted. */
	bgColor?: Color
}

/**
 * Native PPTX picture (image) fill — fills a shape's interior with a bitmap
 * (`<a:blipFill>`). Provide exactly one of `path` or `data`; raster formats only
 * (PNG/JPEG/GIF/BMP/WebP). SVG is not yet supported as a fill source.
 */
export interface ImageFillProps {
	/** Image file path (Node filesystem path or URL). */
	path?: string
	/** Pre-encoded base64 data URI, e.g. `'image/png;base64,iVBOR...'`. */
	data?: string
}

export type Margin = number | [number, number, number, number]
export type HAlign = 'left' | 'center' | 'right' | 'justify'
export type VAlign = 'top' | 'middle' | 'bottom'
/**
 * Text body `vert` attribute — flow/rotation direction of the text within its box.
 * Maps to `<a:bodyPr vert="…">` (ECMA-376 `ST_TextVerticalType`).
 */
export type TextVertType = 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'

/**
 * A single node of a freeform (`custGeom`) path.
 * - coordinates are authored in the object's own inch/EMU space (0..width, 0..height), not slide-relative and not normalized
 * - used by shapes (`pptx.shapes.CUSTOM_GEOMETRY`) and by images (clips the picture to the path)
 */
export type GeometryPoint =
	| { x: Coord, y: Coord, moveTo?: boolean }
	| { x: Coord, y: Coord, curve: { type: 'arc', hR: Coord, wR: Coord, stAng: number, swAng: number } }
	| { x: Coord, y: Coord, curve: { type: 'cubic', x1: Coord, y1: Coord, x2: Coord, y2: Coord } }
	| { x: Coord, y: Coord, curve: { type: 'quadratic', x1: Coord, y1: Coord } }
	| { close: true }

// used by charts, shape, text
export interface BorderProps {
	/**
	 * Border type
	 * @default solid
	 */
	type?: 'none' | 'dash' | 'solid'
	/**
	 * Border color (hex)
	 * @example 'FF3399'
	 * @default '666666'
	 */
	color?: HexColor

	// TODO: add `transparency` prop to Borders (0-100%)

	// TODO: add `width` - deprecate `pt`
	/**
	 * Border size (points)
	 * @default 1
	 */
	pt?: number
	/**
	 * Line end cap style
	 * @default 'flat'
	 */
	cap?: LineCap
}
// used by: image, object, text,
export interface HyperlinkProps {
	_rId?: number
	/**
	 * Slide number to link to
	 */
	slide?: number
	/**
	 * Url to link to
	 */
	url?: string
	/**
	 * Hyperlink Tooltip
	 */
	tooltip?: string
}
// used by: chart, text, image
export interface ShadowProps {
	/**
	 * shadow type
	 * @default 'none'
	 */
	type: 'outer' | 'inner' | 'none'
	/**
	 * opacity (percent)
	 * - range: 0.0-1.0
	 * @example 0.5 // 50% opaque
	 */
	opacity?: number // TODO: "Transparency (0-100%)" in PPT // TODO: deprecate and add `transparency`
	/**
	 * blur (points)
	 * - range: 0-100
	 * @default 0
	 */
	blur?: number
	/**
	 * angle (degrees)
	 * - range: 0-359
	 * @default 0
	 */
	angle?: number
	/**
	 * shadow offset (points)
	 * - range: 0-200
	 * @default 0
	 */
	offset?: number // TODO: "Distance" in PPT
	/**
	 * shadow color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * whether to rotate shadow with shape
	 * @default false
	 */
	rotateWithShape?: boolean
}
// used by: shape, table, text
export interface ShapeFillProps {
	/**
	 * Fill color
	 * - `HexColor` or `ThemeColor`
	 * @example 'FF0000' // hex color (red)
	 * @example pptx.SchemeColor.text1 // Theme color (Text1)
	 */
	color?: Color
	/**
	 * Transparency (percent)
	 * - MS-PPT > Format Shape > Fill & Line > Fill > Transparency
	 * - range: 0-100
	 * @default 0
	 */
	transparency?: number
	/**
	 * Fill type
	 * @default 'solid'
	 */
	type?: 'none' | 'solid' | 'gradient' | 'pattern' | 'image'

	/**
	 * Native PPTX gradient fill options.
	 */
	gradient?: GradientFillProps

	/**
	 * Native PPTX pattern fill options.
	 */
	pattern?: PatternFillProps

	/**
	 * Native PPTX picture fill options. Setting this (or `type: 'image'`) fills the
	 * shape interior with a stretched bitmap instead of a color.
	 */
	image?: ImageFillProps

	/**
	 * Transparency (percent)
	 * @deprecated v3.3.0 - use `transparency`
	 */
	alpha?: number

	/**
	 * Resolved media relationship id for an image fill, assigned at add-time.
	 * @internal
	 */
	_imgRid?: number
}
/**
 * Line (stroke) options.
 *
 * A stroke is painted like a fill: in addition to a solid `color`, it accepts the
 * inherited `gradient`/`pattern`/`image` fill options (DrawingML allows the same
 * fill group inside `<a:ln>`). Setting `gradient` (or `type: 'gradient'`) paints a
 * gradient stroke, e.g. `line: { width: 1, gradient: { kind: 'linear', angle: 0,
 * stops: [{ position: 0, color: 'accent3' }, { position: 100, color: 'accent4' }] } }`.
 */
export interface ShapeLineProps extends ShapeFillProps {
	/**
	 * Line width (pt)
	 * @default 1
	 */
	width?: number
	/**
	 * Dash type
	 * @default 'solid'
	 */
	dashType?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
	/**
	 * Line end cap style
	 * @default 'flat'
	 */
	cap?: LineCap
	/**
	 * Begin arrow type
	 * @since v3.3.0
	 */
	beginArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	/**
	 * End arrow type
	 * @since v3.3.0
	 */
	endArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	// FUTURE: beginArrowSize (1-9)
	// FUTURE: endArrowSize (1-9)

	/**
	 * Dash type
	 * @deprecated v3.3.0 - use `dashType`
	 */
	lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
	/**
	 * @deprecated v3.3.0 - use `beginArrowType`
	 */
	lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	/**
	 * @deprecated v3.3.0 - use `endArrowType`
	 */
	lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	/**
	 * Line width (pt)
	 * @deprecated v3.3.0 - use `width`
	 */
	pt?: number
	/**
	 * Line size (pt)
	 * @deprecated v3.3.0 - use `width`
	 */
	size?: number
}
/**
 * Connector routing style. Maps to a connector preset geometry. The exact preset
 * also depends on `bends` (number of jogs):
 * `straight`→`straightConnector1`; `elbow`→`bentConnector{3,4,5}`;
 * `curved`→`curvedConnector{3,4,5}`.
 */
export type ConnectorType = 'straight' | 'elbow' | 'curved'
/**
 * A connector is a line drawn between two points, emitted as a PowerPoint connector
 * (`<p:cxnSp>`) so the app treats it as a connector (selectable/reroutable) rather than a
 * plain line shape. Endpoints are given directly; the bounding box and flip flags are derived.
 */
export interface ConnectorProps {
	/**
	 * Routing style
	 * @default 'straight'
	 */
	type?: ConnectorType
	/**
	 * Number of adjustable bends (jogs) for an `elbow` / `curved` connector. Selects the
	 * preset variant and how many `adj` values it accepts:
	 * - `1` → `bentConnector3` / `curvedConnector3` (one jog) — the default
	 * - `2` → `bentConnector4` / `curvedConnector4` (two jogs)
	 * - `3` → `bentConnector5` / `curvedConnector5` (three jogs)
	 *
	 * Ignored for `type: 'straight'` (a straight connector has no bends).
	 * @default 1
	 */
	bends?: 1 | 2 | 3
	/**
	 * Bend position(s) as a percent of the connector box (`0`–`100`), one value per bend.
	 * A single number sets the sole jog of a one-bend `elbow` / `curved`; an array sets each
	 * jog of a multi-bend connector and its length must equal `bends`. Values outside `0`–`100`
	 * are allowed (they place the bend beyond the endpoint box, as PowerPoint itself does when
	 * endpoints flip). When omitted, PowerPoint uses the preset default (50%).
	 *
	 * Emitted as `<a:gd name="adj1…" fmla="val …"/>` adjust guides (OOXML 1000ths-of-a-percent).
	 */
	adj?: number | number[]
	/**
	 * Bind the connector's start point to a shape on the **same slide**, referenced by that
	 * shape's `objectName`. Emits `<a:stCxn id=… idx=…>`, so PowerPoint treats the endpoint as
	 * attached: it reroutes when the shape moves and its elbow auto-router can engage.
	 * The shape's `objectName` must be set and unique on the slide. `x1`/`y1` remain the static
	 * fallback geometry (and are used if the name can't be resolved).
	 */
	startShape?: string
	/**
	 * Connection-site index on `startShape` (0-based; the valid range is preset-dependent — a
	 * shape's `<a:cxnLst>` enumerates its sites). Ignored without `startShape`.
	 * @default 0
	 */
	startShapeIdx?: number
	/** Bind the connector's end point to a shape on the same slide (by `objectName`). Emits `<a:endCxn>`. See {@link startShape}. */
	endShape?: string
	/**
	 * Connection-site index on `endShape`. Ignored without `endShape`.
	 * @default 0
	 */
	endShapeIdx?: number
	/** Start point X — inches, or a `Coord` such as `'50%'` / `'2in'` */
	x1: Coord
	/** Start point Y */
	y1: Coord
	/** End point X */
	x2: Coord
	/** End point Y */
	y2: Coord
	/**
	 * Line color (6-digit hex, no `#`)
	 * @default '000000'
	 */
	color?: HexColor
	/**
	 * Line width (pt)
	 * @default 1
	 */
	width?: number
	/** Dash style */
	dashType?: ShapeLineProps['dashType']
	/** Arrowhead at the start point */
	beginArrowType?: ShapeLineProps['beginArrowType']
	/** Arrowhead at the end point */
	endArrowType?: ShapeLineProps['endArrowType']
	/** Selection Pane object name */
	objectName?: string
	/** Accessibility alt text */
	altText?: string
}
// used by: chart, slide, table, text
export interface TextBaseProps {
	/**
	 * Horizontal alignment
	 * @default 'left'
	 */
	align?: HAlign
	/**
	 * Bold style
	 * @default false
	 */
	bold?: boolean
	/**
	 * Add a line-break
	 * @default false
	 */
	breakLine?: boolean
	/**
	 * Preset text warp / WordArt shape (`<a:bodyPr><a:prstTxWarp prst="..">`), which
	 * bends the text along a preset path (arch, circle, wave, …). The value is an
	 * OOXML `ST_TextShapeType` preset name.
	 * @since v4.0.0
	 * @example 'textArchUp' // bend text along an upward arch (e.g. a label following a ring/arc)
	 * @example 'textCircle'
	 */
	textWarp?: string
	/**
	 * Add standard or custom bullet
	 * - use `true` for standard bullet
	 * - pass object options for custom bullet
	 * @default false
	 */
	bullet?:
	| boolean
	| {
		/**
		 * Bullet type
		 * @default bullet
		 */
		type?: 'bullet' | 'number'
		/**
		 * Bullet character code (unicode)
		 * @since v3.3.0
		 * @example '25BA' // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
		 */
		characterCode?: string
		/**
		 * Bullet glyph font typeface (`<a:buFont/>`), e.g. for symbol-font bullets
		 * @since v4.0.0
		 * @example 'Wingdings' // render `characterCode` using the Wingdings font
		 */
		fontFace?: string
		/**
		 * Bullet glyph size as a percentage of the run's text size (25–400)
		 * @since v4.0.0
		 * @default 100
		 * @example 80 // bullet glyph is 80% of the text size
		 */
		size?: number
		/**
		 * Indentation (space between bullet and text) (points)
		 * @since v3.3.0
		 * @default 27 // DEF_BULLET_MARGIN
		 * @example 10 // Indents text 10 points from bullet
		 */
		indent?: number
		/**
		 * Number type
		 * @since v3.3.0
		 * @example 'romanLcParenR' // roman numerals lower-case with paranthesis right
		 */
		numberType?:
		| 'alphaLcParenBoth'
		| 'alphaLcParenR'
		| 'alphaLcPeriod'
		| 'alphaUcParenBoth'
		| 'alphaUcParenR'
		| 'alphaUcPeriod'
		| 'arabicParenBoth'
		| 'arabicParenR'
		| 'arabicPeriod'
		| 'arabicPlain'
		| 'romanLcParenBoth'
		| 'romanLcParenR'
		| 'romanLcPeriod'
		| 'romanUcParenBoth'
		| 'romanUcParenR'
		| 'romanUcPeriod'
		/**
		 * Number bullets start at
		 * @since v3.3.0
		 * @default 1
		 * @example 10 // numbered bullets start with 10
		 */
		numberStartAt?: number
		/**
		 * Image to use as the bullet glyph ("picture bullet", `<a:buBlip>`)
		 * - supply an image `path` (filesystem/URL) or base64 `data` (same forms as `addImage()`)
		 * - raster formats (PNG/JPG/GIF) and SVG are supported; use `size` to scale relative to the text height
		 * - SVG bullets embed a PNG preview plus the SVG (the same dual-rel handling as `addImage()`)
		 * - takes precedence over `type`/`characterCode` when set
		 * @since v4.0.0
		 * @example image: { path: 'images/star.png' }
		 * @example image: { data: 'image/png;base64,iVBOR...' }
		 * @example image: { path: 'images/star.svg' }
		 */
		image?: { path?: string, data?: string }
		/**
		 * Relationship id assigned to a picture-bullet image (`<a:blip r:embed>`)
		 * - for SVG bullets this is the PNG-preview rel; the SVG rel is `_rIdSvg`
		 * @internal populated by `addText()`; do not set directly
		 */
		_rId?: number
		/**
		 * Relationship id of the SVG image for an SVG picture bullet (`<asvg:svgBlip r:embed>`)
		 * @internal populated by `addText()`; do not set directly
		 */
		_rIdSvg?: number

		// DEPRECATED

		/**
		 * Bullet code (unicode)
		 * @deprecated v3.3.0 - use `characterCode`
		 */
		code?: string
		/**
		 * Margin between bullet and text
		 * @since v3.2.1
		 * @deprecated v3.3.0 - use `indent`
		 */
		marginPt?: number
		/**
		 * Number to start with (only applies to type:number)
		 * @deprecated v3.3.0 - use `numberStartAt`
		 */
		startAt?: number
		/**
		 * Number type
		 * @deprecated v3.3.0 - use `numberType`
		 */
		style?: string
		/**
		 * Bullet glyph color (separate from the text run color)
		 * @since v4.0.0
		 * @example 'FF0000' // red bullet
		 */
		color?: HexColor
	}
	/**
	 * Text capitalization
	 * - `'all'` = ALL CAPS
	 * - `'small'` = Small Caps
	 * - `'none'` = no override (default)
	 * - PowerPoint: Font > Effects > All Caps / Small Caps
	 */
	caps?: 'none' | 'small' | 'all'
	/**
	 * Text color
	 * - `HexColor` or `ThemeColor`
	 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Color
	 * @example 'FF0000' // hex color (red)
	 * @example pptx.SchemeColor.text1 // Theme color (Text1)
	 */
	color?: Color
	/**
	 * Font face name
	 *
	 * Applied to the Latin (`<a:latin>`) and complex-script (`<a:cs>`) font slots, matching
	 * how PowerPoint writes a font picked from the UI. The East Asian slot (`<a:ea>`) is left
	 * to inherit from the theme unless `fontFaceEA` is set — forcing a Latin-only face into the
	 * East Asian slot duplicates/ghosts text in Office 365.
	 * @example 'Arial' // Arial font
	 */
	fontFace?: string
	/**
	 * East Asian font face name (`<a:ea>` slot), used to render CJK (Chinese/Japanese/Korean) glyphs
	 *
	 * Set this when the East Asian font differs from `fontFace`. When omitted, `<a:ea>` inherits the
	 * theme East Asian font, which is what PowerPoint does for Latin fonts.
	 * @example '微軟正黑體' // render East Asian glyphs with Microsoft JhengHei
	 */
	fontFaceEA?: string
	/**
	 * Font size
	 * @example 12 // Font size 12
	 */
	fontSize?: number
	/**
	 * Text highlight color (hex format)
	 * @example 'FFFF00' // yellow
	 */
	highlight?: HexColor
	/**
	 * italic style
	 * @default false
	 */
	italic?: boolean
	/**
	 * language
	 * - ISO 639-1 standard language code
	 * @default 'en-US' // english US
	 * @example 'fr-CA' // french Canadian
	 */
	lang?: string
	/**
	 * Add a soft line-break (shift+enter) before line text content
	 * @default false
	 * @since v3.5.0
	 */
	softBreakBefore?: boolean
	/**
	 * tab stops
	 * - PowerPoint: Paragraph > Tabs > Tab stop position
	 * @example [{ position:1 }, { position:3 }] // Set first tab stop to 1 inch, set second tab stop to 3 inches
	 */
	tabStops?: Array<{ position: number, alignment?: 'l' | 'r' | 'ctr' | 'dec' }>
	/**
	 * text direction
	 * `horz` = horizontal
	 * `vert` = rotate 90^
	 * `vert270` = rotate 270^
	 * `wordArtVert` = stacked
	 * @default 'horz'
	 */
	textDirection?: 'horz' | 'vert' | 'vert270' | 'wordArtVert'
	/**
	 * Transparency (percent)
	 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Transparency
	 * - range: 0-100
	 * @default 0
	 */
	transparency?: number
	/**
	 * underline properties
	 * - PowerPoint: Font > Color & Underline > Underline Style/Underline Color
	 * @default (none)
	 */
	underline?: {
		style?:
		| 'dash'
		| 'dashHeavy'
		| 'dashLong'
		| 'dashLongHeavy'
		| 'dbl'
		| 'dotDash'
		| 'dotDashHeave'
		| 'dotDotDash'
		| 'dotDotDashHeavy'
		| 'dotted'
		| 'dottedHeavy'
		| 'heavy'
		| 'none'
		| 'sng'
		| 'wavy'
		| 'wavyDbl'
		| 'wavyHeavy'
		color?: Color
	}
	/**
	 * vertical alignment
	 * @default 'top'
	 */
	valign?: VAlign
}
export interface PlaceholderProps extends PositionProps, TextBaseProps, ObjectNameProps {
	name: string
	type: PLACEHOLDER_TYPE
	/**
	 * margin (points)
	 */
	margin?: Margin
	/**
	 * Preset shape geometry for this placeholder (e.g. `'roundRect'`)
	 * @default 'rect'
	 */
	shape?: SHAPE_NAME
	/**
	 * Rounded rectangle corner radius (inches) when `shape: 'roundRect'`
	 * - range: 0.0 to slide height/2
	 */
	rectRadius?: number
}
export interface ObjectNameProps {
	/**
	 * Object name
	 * - used instead of default "Object N" name
	 * - PowerPoint: Home > Arrange > Selection Pane...
	 * @since v3.10.0
	 * @default 'Object 1'
	 * @example 'Antenna Design 9'
	 */
	objectName?: string
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - serialized to the generated object's `p:cNvPr` `descr` attribute
	 * - PowerPoint: [right-click on the object] > "Edit Alt Text..."
	 * @since v4.0.0
	 * @example 'Quarterly revenue bar chart'
	 */
	altText?: string
	/**
	 * Object lock flags (DrawingML `a:spLocks` / `a:picLocks` / `a:graphicFrameLocks`)
	 * - restrict how the object can be manipulated in PowerPoint (e.g. prevent moving, resizing, or grouping)
	 * - each flag maps 1:1 to the OOXML attribute of the same name; only flags set to `true` are emitted
	 * - PowerPoint UI: Selection Pane / right-click protections (most locks are honored at edit time, not as a password)
	 * - flags only apply to the object types that support them (see each flag); flags set on an unsupported
	 *   object type are ignored with a console warning
	 * @since v4.0.0
	 * @example { noMove: true, noResize: true } // pin an object in place
	 * @example { noGrp: true } // exclude from grouping
	 */
	objectLock?: ObjectLockProps
}
/**
 * Object lock flags. Maps to DrawingML locking elements:
 * - shapes / text boxes / placeholders → `a:spLocks`
 * - images / media → `a:picLocks`
 * - tables → `a:graphicFrameLocks`
 *
 * Each property mirrors the OOXML attribute name. A flag is only serialized for object types whose
 * locking element defines it (noted per-flag); setting an unsupported flag logs a warning and is ignored.
 * @since v4.0.0
 */
export interface ObjectLockProps {
	/** Disallow grouping/ungrouping with other objects. (shapes, images, tables) */
	noGrp?: boolean
	/** Disallow selecting the object. (shapes, images, tables) */
	noSelect?: boolean
	/** Disallow moving the object. (shapes, images, tables) */
	noMove?: boolean
	/** Disallow resizing the object. (shapes, images, tables) */
	noResize?: boolean
	/** Disallow changing the aspect ratio. (shapes, images, tables) */
	noChangeAspect?: boolean
	/** Disallow rotating the object. (shapes, images) */
	noRot?: boolean
	/** Disallow editing the freeform/custom-geometry points. (shapes, images) */
	noEditPoints?: boolean
	/** Disallow moving the shape's adjustment handles. (shapes, images) */
	noAdjustHandles?: boolean
	/** Disallow changing arrowheads. (shapes, images) */
	noChangeArrowheads?: boolean
	/** Disallow changing the shape type (preset geometry). (shapes, images) */
	noChangeShapeType?: boolean
	/** Disallow editing the text body. (shapes / text boxes) */
	noTextEdit?: boolean
	/** Disallow cropping the picture. (images) */
	noCrop?: boolean
	/** Disallow drilling down into the graphical object (e.g. chart data). (tables) */
	noDrilldown?: boolean
}
/**
 * Theme color scheme overrides. Each slot maps to one `<a:clrScheme>` entry in `theme1.xml`;
 * any slot left unset keeps the default Office value. Provide 6-digit hex (no `#`).
 *
 * Slot names use the OOXML scheme names. The PowerPoint UI labels them as:
 * `dk1`=Text/Dark 1, `lt1`=Background/Light 1, `dk2`=Text/Dark 2, `lt2`=Background/Light 2,
 * `accent1`-`accent6`=Accent 1-6, `hlink`=Hyperlink, `folHlink`=Followed Hyperlink.
 * @example { accent1: 'C00000', dk2: '1F3864', hlink: '0070C0' }
 */
export interface ThemeColorScheme {
	/** Text/Dark 1 (default Office: black via `windowText`) */
	dk1?: HexColor
	/** Background/Light 1 (default Office: white via `window`) */
	lt1?: HexColor
	/** Text/Dark 2 (default Office: `44546A`) */
	dk2?: HexColor
	/** Background/Light 2 (default Office: `E7E6E6`) */
	lt2?: HexColor
	/** Accent 1 (default Office: `4472C4`) */
	accent1?: HexColor
	/** Accent 2 (default Office: `ED7D31`) */
	accent2?: HexColor
	/** Accent 3 (default Office: `A5A5A5`) */
	accent3?: HexColor
	/** Accent 4 (default Office: `FFC000`) */
	accent4?: HexColor
	/** Accent 5 (default Office: `5B9BD5`) */
	accent5?: HexColor
	/** Accent 6 (default Office: `70AD47`) */
	accent6?: HexColor
	/** Hyperlink (default Office: `0563C1`) */
	hlink?: HexColor
	/** Followed Hyperlink (default Office: `954F72`) */
	folHlink?: HexColor
}
export interface ThemeProps {
	/**
	 * Headings font face name
	 * @example 'Arial Narrow'
	 * @default 'Calibri Light'
	 */
	headFontFace?: string
	/**
	 * Body font face name
	 * @example 'Arial'
	 * @default 'Calibri'
	 */
	bodyFontFace?: string
	/**
	 * Headings East Asian font face — theme `<a:ea>` slot of the major font.
	 * Used for CJK (Chinese/Japanese/Korean) runs that fall back to the theme font.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Yu Gothic'
	 */
	headFontFaceEA?: string
	/**
	 * Body East Asian font face — theme `<a:ea>` slot of the minor font.
	 * Used for CJK (Chinese/Japanese/Korean) runs that fall back to the theme font.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Yu Gothic'
	 */
	bodyFontFaceEA?: string
	/**
	 * Headings complex-script font face — theme `<a:cs>` slot of the major font.
	 * Used for complex scripts such as Arabic, Hebrew, Thai, and Devanagari.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Arial'
	 */
	headFontFaceCS?: string
	/**
	 * Body complex-script font face — theme `<a:cs>` slot of the minor font.
	 * Used for complex scripts such as Arabic, Hebrew, Thai, and Devanagari.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Arial'
	 */
	bodyFontFaceCS?: string
	/**
	 * Theme color scheme overrides written to `ppt/theme/theme1.xml`.
	 * - any unset slot keeps its default Office value
	 * - references such as `pptx.SchemeColor.accent1` resolve against these values
	 * @example { accent1: 'C00000', accent2: '00B050', hlink: '0070C0' }
	 */
	colorScheme?: ThemeColorScheme
}

// image / media ==================================================================================
export type MediaType = 'audio' | 'online' | 'video'

interface ImageBaseProps extends PositionProps, ObjectNameProps {
	/**
	 * Sizing note (`w`/`h` inherited from {@link PositionProps}):
	 * - When a `data` (base64) image is supplied and `w`/`h` are omitted, the natural pixel
	 *   size is read from the image header (PNG/JPEG/GIF/BMP/WebP) and used at 96 DPI
	 *   (natural pixels / 96 = inches).
	 * - When only one of `w`/`h` is given, the other is derived from the natural aspect ratio.
	 * - `path` images and vector (SVG) data cannot be measured synchronously, so an omitted
	 *   dimension falls back to 1 inch.
	 */
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - PowerPoint: [right-click on an image] > "Edit Alt Text..."
	 */
	altText?: string
	/**
	 * Flip horizontally?
	 * @default false
	 */
	flipH?: boolean
	/**
	 * Flip vertical?
	 * @default false
	 */
	flipV?: boolean
	hyperlink?: HyperlinkProps
	/**
	 * Border line (`<a:ln>` outline) drawn around the image
	 * - same options as a shape outline; a picture supports a single outline, not per-side borders
	 * - MS-PPT: Format Picture > Line
	 * @example { color: '0088CC', width: 2 }                   // 2pt blue border
	 * @example { color: '666666', width: 1, dashType: 'dash' } // dashed gray border
	 */
	line?: ShapeLineProps
	/**
	 * Name of a picture placeholder defined on the slide layout/master to populate
	 * - when it matches a layout/master placeholder, the image inherits that placeholder's
	 *   position and size for any of `x`/`y`/`w`/`h` not supplied explicitly (issue #1258);
	 *   explicit values always win
	 * @example 'picph'
	 * @see https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
	 */
	placeholder?: string
	/**
	 * Image rotation (degrees)
	 * - range: -360 to 360
	 * @default 0
	 * @example 180 // rotate image 180 degrees
	 */
	rotate?: number
	/**
	 * Enable image rounding (clips the image to a circle/ellipse)
	 * - shorthand for `shape: 'ellipse'`; `shape` takes precedence when both are set
	 * @default false
	 */
	rounding?: boolean
	/**
	 * Clip the image to a preset shape geometry ("fit image into shape")
	 * - accepts any PowerPoint preset geometry name, e.g. `'roundRect'`, `'hexagon'`, `'ellipse'`
	 * - combine with `sizing: { type: 'cover', ... }` for an aspect-correct fill of the shape box
	 * - use `rectRadius` to set the corner radius for `'roundRect'`
	 * @example 'roundRect' // rounded-rectangle avatar
	 * @example 'hexagon'   // hexagonal photo
	 */
	shape?: SHAPE_NAME
	/**
	 * Clip the image to an arbitrary freeform path (`custGeom`)
	 * - takes precedence over `shape` / `rounding` when present
	 * - coordinates are authored in the image's own inch/EMU space (0..w, 0..h), not slide-relative and not normalized
	 * - supports the same path DSL as freeform shapes: `moveTo` / `lnTo` / `cubicBezTo` / `quadBezTo` / `arcTo` / `close`
	 * @example [{ x: 1, y: 0 }, { x: 2, y: 2 }, { x: 0, y: 2 }, { close: true }] // triangular photo clip
	 */
	points?: GeometryPoint[]
	/**
	 * Rounded rectangle corner radius (inches) when `shape: 'roundRect'`
	 * - values: 0.0 to 1.0
	 * @default 0
	 */
	rectRadius?: number
	/**
	 * Preset-geometry adjustment handles (`<a:avLst>` guides) for the clip `shape`.
	 * - tune adjustment handles that lack a dedicated option, e.g. chevron point depth
	 * - accepts a single guide or an array; each `value` is a `0.0–1.0` fraction (see {@link ShapeAdjustValue})
	 * @since v4.0.0
	 * @example { name: 'adj', value: 0.25 }
	 */
	shapeAdjust?: ShapeAdjustValue | ShapeAdjustValue[]
	/**
	 * Shadow Props
	 * - MS-PPT > Format Picture > Shadow
	 * @example
	 * { type: 'outer', color: '000000', opacity: 0.5, blur: 20,  offset: 20, angle: 270 }
	 */
	shadow?: ShadowProps
	/**
	 * Crop the source image to a sub-region by percentage edge insets, emitted verbatim
	 * as OOXML `<a:srcRect>`.
	 * - each value is the percent (0–100) trimmed off that edge of the *source* image, the
	 *   same model PowerPoint's Picture Format > Crop uses; `l`+`r` and `t`+`b` must each be < 100
	 * - the remaining sub-region is stretched to fill the picture's displayed `w`×`h` box, so this
	 *   is the faithful way to reproduce a deck that maps several icons out of one composite raster
	 * - operates on the source image directly (independent of natural-pixel measurement), so unlike
	 *   `sizing: 'crop'` (which crops in *displayed inches*) it works for SVG and unmeasurable formats
	 * - mutually exclusive with {@link sizing}; if both are set `crop` wins and `sizing` is ignored
	 * @example { l: 0, t: 0, r: 50, b: 50 } // keep the top-left quadrant of the source image
	 */
	crop?: {
		/** Percent (0–100) trimmed from the left edge of the source image. @default 0 */
		l?: number
		/** Percent (0–100) trimmed from the top edge of the source image. @default 0 */
		t?: number
		/** Percent (0–100) trimmed from the right edge of the source image. @default 0 */
		r?: number
		/** Percent (0–100) trimmed from the bottom edge of the source image. @default 0 */
		b?: number
	}
	/**
	 * Image sizing options
	 */
	sizing?: {
		/**
		 * Sizing type
		 * - `cover` / `contain` fit the image into the `w`×`h` box using the image's *natural*
		 *   pixel aspect ratio (read from the embedded PNG/JPEG/GIF/BMP/WebP header). If the
		 *   natural size cannot be determined (e.g. SVG or an unrecognized format) the displayed
		 *   `w`/`h` ratio is used as a fallback and a warning is logged.
		 * - `crop` cuts a window out of the displayed image using the `x`/`y`/`w`/`h` offsets.
		 */
		type: 'contain' | 'cover' | 'crop'
		/**
		 * Image width
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		w: Coord
		/**
		 * Image height
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		h: Coord
		/**
		 * Offset from left to crop image
		 * - `crop` only
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		x?: Coord
		/**
		 * Offset from top to crop image
		 * - `crop` only
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		y?: Coord
	}
	/**
	 * Transparency (percent)
	 * - MS-PPT > Format Picture > Picture > Picture Transparency > Transparency
	 * - range: 0-100
	 * @default 0
	 * @example 25 // 25% transparent
	 */
	transparency?: number
	/**
	 * Recolor the image as a two-tone (duotone) effect
	 * - maps the image's shadows to `shadow` and its highlights to `highlight`
	 * - serializes `<a:duotone>` inside the picture's `<a:blip>` (MS-PPT > Format Picture > Picture Color > Recolor)
	 * - colors accept `HexColor` or `ThemeColor`, same as fills
	 * - the classic brand treatment: tint stock photography into a single brand hue
	 * @example { shadow: '250F6B', highlight: 'FFFFFF' } // deep-blue duotone
	 */
	duotone?: {
		/** Color mapped to the image's dark/shadow tones. */
		shadow: Color
		/** Color mapped to the image's light/highlight tones. */
		highlight: Color
	}
	/**
	 * Raw SVG markup to embed as the image source
	 * - convenience for `data: 'data:image/svg+xml;base64,...'`; PptxGenJS encodes it for you
	 * - ignored when `data` or `path` is also provided
	 * @example '<svg viewBox="0 0 24 24">...</svg>'
	 */
	svg?: string
}
export type ImageProps = ImageBaseProps & (DataOrPathRequiredProps | (DataOrPathProps & { svg: string }))
/**
 * Add media (audio/video) to slide
 * Requires either `data` or `path`; online media requires `link`.
 */
interface MediaBaseProps extends PositionProps, ObjectNameProps {
	/**
	 * Cover image
	 * @since 3.9.0
	 * @default "play button" image, gray background
	 */
	cover?: string
	/**
	 * media file extension
	 * - use when the media file path does not already have an extension, ex: "/folder/SomeSong"
	 * @since 3.9.0
	 * @default extension from file provided
	 */
	extn?: string
	/**
	 * Loop playback indefinitely (PowerPoint "Playback > Loop until Stopped")
	 * - emits a slide timing tree so the embedded audio/video repeats when played
	 * - not supported for `type: 'online'` (e.g. YouTube) embeds
	 * @since 4.0.0
	 * @default false
	 */
	loop?: boolean
	/**
	 * Total number of times to play the media (a finite loop), ex: `3` plays it three times
	 * - ignored when `loop` is `true` (which repeats forever)
	 * - not supported for `type: 'online'` (e.g. YouTube) embeds
	 * @since 4.0.0
	 */
	loopCount?: number
}
export type MediaProps = MediaBaseProps &
	(
		| (DataOrPathRequiredProps & {
			/**
			 * Media type
			 */
			type: Exclude<MediaType, 'online'>
			/**
			 * Optional video embed link metadata.
			 */
			link?: string
		})
		| (DataOrPathProps & {
			/**
			 * Use 'online' to embed a YouTube video (only supported in recent versions of PowerPoint)
			 */
			type: 'online'
			/**
			 * video embed link
			 * - works with YouTube
			 * - other sites may not show correctly in PowerPoint
			 * @example 'https://www.youtube.com/embed/Dph6ynRVyUc' // embed a youtube video
			 */
			link: string
		})
	)

// shapes =========================================================================================

/**
 * A single preset-geometry adjustment guide (`<a:gd>` inside `<a:avLst>`).
 * - `name` is the guide name the preset defines, e.g. `'adj'`, `'adj1'`, `'adj2'`.
 *   PowerPoint shows these handles as the yellow drag dots on a selected shape.
 * - `value` is a fraction `0.0–1.0` of the handle's range, emitted as a percentage
 *   guide formula (`val`, in 1/100000 units, so `0.25` → `fmla="val 25000"`).
 *   Most adjustment handles (corner radius, chevron point, callout depth, bevel
 *   width, …) are percentage-based and map directly; some shapes accept values
 *   beyond `1.0`. For angle-based handles, prefer the `angleRange` shortcut.
 * @since v4.0.0
 */
export interface ShapeAdjustValue {
	name: string
	value: number
}

export interface ShapeProps extends PositionProps, ObjectNameProps {
	/**
	 * Horizontal alignment
	 * @default 'left'
	 */
	align?: HAlign
	/**
	 * Radius (only for pptx.shapes.PIE, pptx.shapes.ARC, pptx.shapes.BLOCK_ARC)
	 * - In the case of pptx.shapes.BLOCK_ARC you have to setup the arcThicknessRatio
	 * - values: [0-359, 0-359]
	 * @since v3.4.0
	 * @default [270, 0]
	 */
	angleRange?: [number, number]
	/**
	 * Preset-geometry adjustment handles (`<a:avLst>` guides) for any preset shape.
	 * - Use this to tune adjustment handles that lack a dedicated shortcut option,
	 *   e.g. chevron/arrow point depth, callout pointer, bevel/frame thickness.
	 * - Accepts a single guide or an array; each `value` is a `0.0–1.0` fraction of
	 *   the handle's range (see {@link ShapeAdjustValue}).
	 * - `rectRadius` / `angleRange` remain friendly shortcuts; any `shapeAdjust`
	 *   guide that does not collide with a shortcut name is emitted in addition.
	 * @since v4.0.0
	 * @example { name: 'adj', value: 0.25 } // set the single adjust handle to 25%
	 * @example [{ name: 'adj1', value: 0.5 }, { name: 'adj2', value: 0.25 }] // two handles
	 */
	shapeAdjust?: ShapeAdjustValue | ShapeAdjustValue[]
	/**
	 * Radius (only for pptx.shapes.BLOCK_ARC)
	 * - You have to setup the angleRange values too
	 * - values: 0.0-1.0
	 * @since v3.4.0
	 * @default 0.5
	 */
	arcThicknessRatio?: number
	/**
	 * Shape fill color properties
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:pptx.SchemeColor.accent1 } // Theme color Accent1
	 */
	fill?: ShapeFillProps
	/**
	 * Flip shape horizontally?
	 * @default false
	 */
	flipH?: boolean
	/**
	 * Flip shape vertical?
	 * @default false
	 */
	flipV?: boolean
	/**
	 * Add hyperlink to shape
	 * @example hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
	 */
	hyperlink?: HyperlinkProps
	/**
	 * Line options
	 */
	line?: ShapeLineProps
	/**
	 * Points (only for pptx.shapes.CUSTOM_GEOMETRY)
	 * - type: 'arc'
	 * - `hR` Shape Arc Height Radius
	 * - `wR` Shape Arc Width Radius
	 * - `stAng` Shape Arc Start Angle
	 * - `swAng` Shape Arc Swing Angle
	 * @see http://www.datypic.com/sc/ooxml/e-a_arcTo-1.html
	 * @example [{ x: 0, y: 0 }, { x: 10, y: 10 }] // draw a line between those two points
	 */
	points?: GeometryPoint[]
	/**
	 * Rounded rectangle radius (only for pptx.shapes.ROUNDED_RECTANGLE)
	 * - values: 0.0 to 1.0
	 * @default 0
	 */
	rectRadius?: number
	/**
	 * Rotation (degrees)
	 * - range: -360 to 360
	 * @default 0
	 * @example 180 // rotate 180 degrees
	 */
	rotate?: number
	/**
	 * Shadow options
	 * TODO: need new demo.js entry for shape shadow
	 */
	shadow?: ShadowProps

	/**
	 * @deprecated v3.3.0
	 */
	lineSize?: number
	/**
	 * @deprecated v3.3.0
	 */
	lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
	/**
	 * @deprecated v3.3.0
	 */
	lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
	/**
	 * @deprecated v3.3.0
	 */
	lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
	/**
	 * Shape name (used instead of default "Shape N" name)
	 * @deprecated v3.10.0 - use `objectName`
	 */
	shapeName?: string
}

// tables =========================================================================================

export interface TableToSlidesProps extends TableProps {
	_arrObjTabHeadRows?: TableRow[]
	// _masterSlide?: SlideLayout

	/**
	 * Add an image to slide(s) created during autopaging
	 * - `image` prop requires either `path` or `data`
	 * - see `DataOrPathProps` for details on `image` props
	 * - see `PositionProps` for details on `options` props
	 */
	addImage?: { image: DataOrPathProps, options: PositionProps }
	/**
	 * Add a shape to slide(s) created during autopaging
	 */
	addShape?: { shapeName: SHAPE_NAME, options: ShapeProps }
	/**
	 * Add a table to slide(s) created during autopaging
	 */
	addTable?: { rows: TableRow[], options: TableProps }
	/**
	 * Add a text object to slide(s) created during autopaging
	 */
	addText?: { text: TextProps[], options: TextPropsOptions }
	/**
	 * Whether to enable auto-paging
	 * - auto-paging creates new slides as content overflows a slide
	 * @default true
	 */
	autoPage?: boolean
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * Whether to repeat head row(s) on new tables created by autopaging
	 * @since v3.3.0
	 * @default false
	 */
	autoPageRepeatHeader?: boolean
	/**
	 * The `y` location to use on subsequent slides created by autopaging
	 * @default (top margin of Slide)
	 */
	autoPageSlideStartY?: number
	/**
	 * Column widths (inches)
	 */
	colW?: number | number[]
	/**
	 * Master slide name
	 * - define a master slide to have your auto-paged slides have corporate design, etc.
	 * @see https://gitbrent.github.io/PptxGenJS/docs/masters.html
	 */
	masterSlideName?: string
	/**
	 * Slide margin
	 * - this margin will be across all slides created by auto-paging
	 */
	slideMargin?: Margin

	/**
	 * @deprecated v3.3.0 - use `autoPageRepeatHeader`
	 */
	addHeaderToEach?: boolean
	/**
	 * @deprecated v3.3.0 - use `autoPageSlideStartY`
	 */
	newSlideStartY?: number
}
export interface TableCellProps extends TextBaseProps {
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * Cell border
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
	/**
	 * Cell colspan
	 */
	colspan?: number
	/**
	 * Fill color
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
	 */
	fill?: ShapeFillProps
	hyperlink?: HyperlinkProps
	/**
	 * Cell margin (inches)
	 * @default 0
	 */
	margin?: Margin
	/**
	 * Cell rowspan
	 */
	rowspan?: number
	/**
	 * Shrink cell text to fit when it would overflow the cell's fixed height.
	 * - `'shrink'` measures the wrapped text and bakes a **reduced literal font size**
	 *   onto the cell's runs so the text fits — PowerPoint does not support text
	 *   autofit (`normAutofit`) inside table cells, so there is no font-scale flag to
	 *   set; the size itself is lowered, which both PowerPoint and LibreOffice render
	 *   identically with no edit/resize.
	 * - Requires the cell font registered via {@link PptxGenJS.registerFontMetrics};
	 *   without metrics it is a no-op (the cell keeps its authored size) and warns once.
	 * - Only triggers when the cell's row has a **fixed** height that the text exceeds.
	 *   With auto-height rows (no `rowH`/`h`), the row simply grows, so nothing shrinks.
	 * - Only `'shrink'` is acted on for cells. `'resize'` and the object form are ignored
	 *   here: a table row already auto-grows to fit its tallest cell (the cell equivalent
	 *   of `spAutoFit`), so there is nothing to bake. (The wider union is shared with
	 *   {@link TextPropsOptions.fit} so table-level `fit` can cascade to cells.)
	 * @since v4.0.0
	 * @example 'shrink' // measured when the cell font is registered; else no-op
	 */
	fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps
}
/**
 * Styling for one region of a custom table style (maps to a `CT_TablePartStyle`).
 * A region (e.g. the header row or banded rows) is shown only when the matching
 * `TableProps` flag is set — `firstRow` needs `hasHeader`, `band1H`/`band2H` need
 * `hasBandedRows`, and so on.
 * @see TableStyleProps
 */
export interface TableStyleRegionProps {
	/**
	 * Solid cell fill color (hex).
	 * - `HexColor` only; theme references are not supported in custom styles
	 * @example '1A2B3C'
	 */
	fill?: HexColor
	/**
	 * Text color (hex).
	 * @example 'FFFFFF'
	 */
	color?: HexColor
	/** Bold text. */
	bold?: boolean
	/** Italic text. */
	italic?: boolean
	/**
	 * Cell border(s).
	 * - single value is applied to all four sides plus the interior grid lines
	 * - array of values in TRBL order styles only the four outer sides
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
}
/**
 * A reusable custom table style written to `ppt/tableStyles.xml`.
 * Pass to `pptx.defineTableStyle()`, which registers it and returns a GUID to use
 * as `TableProps.tableStyle`. Unlike the fixed built-in `TABLE_STYLE` set, a custom
 * style can use arbitrary brand colors, is editable in PowerPoint's Table Styles
 * gallery, and bands correctly across any row/column count (including auto-paged tables).
 * @example
 * const brand = pptx.defineTableStyle({
 *   name: 'Brand Banded',
 *   wholeTbl: { border: { type:'solid', color:'D9D9D9', pt:0.5 } },
 *   firstRow: { fill:'1A2B3C', color:'FFFFFF', bold:true },
 *   band1H:   { fill:'EAF1F8' },
 *   band2H:   { fill:'FFFFFF' },
 * })
 * slide.addTable(rows, { tableStyle: brand, hasHeader:true, hasBandedRows:true })
 */
export interface TableStyleProps {
	/** Display name shown in PowerPoint's Table Styles gallery. */
	name: string
	/** Base styling applied to every cell. */
	wholeTbl?: TableStyleRegionProps
	/** Header (first) row — activated by `TableProps.hasHeader`. */
	firstRow?: TableStyleRegionProps
	/** Footer (last) row — activated by `TableProps.hasFooter`. */
	lastRow?: TableStyleRegionProps
	/** First column — activated by `TableProps.hasFirstColumn`. */
	firstCol?: TableStyleRegionProps
	/** Last column — activated by `TableProps.hasLastColumn`. */
	lastCol?: TableStyleRegionProps
	/** Odd horizontal band — activated by `TableProps.hasBandedRows`. */
	band1H?: TableStyleRegionProps
	/** Even horizontal band — activated by `TableProps.hasBandedRows`. */
	band2H?: TableStyleRegionProps
	/** Odd vertical band — activated by `TableProps.hasBandedColumns`. */
	band1V?: TableStyleRegionProps
	/** Even vertical band — activated by `TableProps.hasBandedColumns`. */
	band2V?: TableStyleRegionProps
}
/**
 * Internal record pairing a registered custom table style with its generated GUID.
 */
export interface TableStyleInternal {
	/** Braced, upper-case GUID emitted as both `styleId` and `<a:tableStyleId>`. */
	guid: string
	def: TableStyleProps
}
export interface TableProps extends PositionProps, TextBaseProps, ObjectNameProps {
	_arrObjTabHeadRows?: TableRow[]

	/**
	 * Name of a table/content placeholder defined on the slide layout/master to bind this table to.
	 * - when it matches a layout/master placeholder, the table's `<p:graphicFrame>` emits that
	 *   placeholder's `<p:ph>` (idx/type) so PowerPoint treats the table as filling the placeholder
	 *   (e.g. a "Title and Content" content placeholder)
	 * - the table also inherits the placeholder's position/size for any of x/y/w/h left unset
	 * @example 'body' // bind to the layout placeholder named 'body'
	 */
	placeholder?: string

	/**
	 * Whether to enable auto-paging
	 * - auto-paging creates new slides as content overflows a slide
	 * @default false
	 */
	autoPage?: boolean
	/**
	 * Auto-paging character weight
	 * - adjusts how many characters are used before lines wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
	 */
	autoPageCharWeight?: number
	/**
	 * Auto-paging line weight
	 * - adjusts how many lines are used before slides wrap
	 * - range: -1.0 to 1.0
	 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
	 * @default 0.0
	 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
	 */
	autoPageLineWeight?: number
	/**
	 * Whether table header row(s) should be repeated on each new slide creating by autoPage.
	 * Use `autoPageHeaderRows` to designate how many rows comprise the table header (1+).
	 * @default false
	 * @since v3.3.0
	 */
	autoPageRepeatHeader?: boolean
	/**
	 * Number of rows that comprise table headers
	 * - required when `autoPageRepeatHeader` is set to true.
	 * @example 2 - repeats the first two table rows on each new slide created
	 * @default 1
	 * @since v3.3.0
	 */
	autoPageHeaderRows?: number
	/**
	 * The `y` location to use on subsequent slides created by autopaging
	 * @default (top margin of Slide)
	 */
	autoPageSlideStartY?: number
	/**
	 * Whether populated placeholders on the source slide (e.g. a title set via
	 * `addText(text, { placeholder })`) are copied onto each overflow slide created by autoPage.
	 * - new slides otherwise inherit only the layout's empty placeholders, so a title set on the
	 *   first slide would not appear on continuation slides (upstream gitbrent/PptxGenJS#1136).
	 * @default false
	 */
	autoPagePlaceholder?: boolean
	/**
	 * Table border
	 * - single value is applied to all 4 sides
	 * - array of values in TRBL order for individual sides
	 */
	border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
	/**
	 * Width of table columns (inches)
	 * - single value is applied to every column equally based upon `w`
	 * - array of values in applied to each column in order
	 * @default columns of equal width based upon `w`
	 */
	colW?: number | number[]
	/**
	 * Mark the first row as a header row.
	 * Emits `firstRow="1"` on `<a:tblPr>`, activating the first-row style region of
	 * the table style and satisfying the PowerPoint accessibility checker's "table header" rule.
	 * @default false
	 */
	hasHeader?: boolean
	/**
	 * Mark the last row as a footer row.
	 * Emits `lastRow="1"` on `<a:tblPr>`, activating the last-row style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasFooter?: boolean
	/**
	 * Enable alternating row (band) shading.
	 * Emits `bandRow="1"` on `<a:tblPr>`, activating band1H/band2H style regions.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasBandedRows?: boolean
	/**
	 * Enable alternating column (band) shading.
	 * Emits `bandCol="1"` on `<a:tblPr>`, activating band1V/band2V style regions.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasBandedColumns?: boolean
	/**
	 * Apply special styling to the first column.
	 * Emits `firstCol="1"` on `<a:tblPr>`, activating the firstCol style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasFirstColumn?: boolean
	/**
	 * Apply special styling to the last column.
	 * Emits `lastCol="1"` on `<a:tblPr>`, activating the lastCol style region.
	 * Requires `tableStyle` to have a visible effect.
	 * @default false
	 */
	hasLastColumn?: boolean
	/**
	 * Lay the table out right-to-left.
	 * Emits `rtl="1"` on `<a:tblPr>`, which mirrors the column order so the first
	 * column renders on the right — the correct layout for RTL scripts (Arabic, Hebrew).
	 * This controls only the table/column direction; per-cell text direction is set
	 * with each cell's `rtlMode` option.
	 * @default false
	 * @since v4.0.0
	 */
	rtl?: boolean
	/**
	 * Table style to apply, either a built-in `TABLE_STYLE` member or the GUID
	 * returned by `pptx.defineTableStyle()` for a custom style.
	 * Emits `<a:tableStyleId>` inside `<a:tblPr>` with the corresponding GUID.
	 * Style flags (`hasHeader`, `hasFooter`, `hasBandedRows`, etc.) select which
	 * regions of the chosen style are activated; they have no visible effect without
	 * a `tableStyle` set.
	 *
	 * @example tableStyle: pptx.TABLE_STYLE.MEDIUM_STYLE_2_ACCENT_1 // built-in
	 * @example const brand = pptx.defineTableStyle({ name:'Brand', firstRow:{ fill:'1A2B3C', color:'FFFFFF', bold:true } }); tableStyle: brand
	 */
	tableStyle?: TABLE_STYLE | string
	/**
	 * Cell background color
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
	 */
	fill?: ShapeFillProps
	/**
	 * Cell margin (inches)
	 * - affects all table cells, is superceded by cell options
	 */
	margin?: Margin
	/**
	 * Height of table rows (inches)
	 * - single value is applied to every row equally based upon `h`
	 * - array of values in applied to each row in order
	 * @default rows of equal height based upon `h`
	 */
	rowH?: number | number[]
	/**
	 * DEV TOOL: Verbose Mode (to console)
	 * - tell the library to provide an almost ridiculous amount of detail during auto-paging calculations
	 * @default false // obviously
	 */
	verbose?: boolean // Undocumented; shows verbose output

	/**
	 * @deprecated v3.3.0 - use `autoPageSlideStartY`
	 */
	newSlideStartY?: number
}
export interface TableCell {
	_type?: SLIDE_OBJECT_TYPES.tablecell
	/** lines in this cell (autoPage) */
	_lines?: TableCell[][]
	/** `text` prop but guaranteed to hold "TableCell[]" */
	_tableCells?: TableCell[]
	/** height in EMU */
	_lineHeight?: number
	_hmerge?: boolean
	_vmerge?: boolean
	_rowContinue?: number
	/** origin cell of a colspan/rowspan span, set on the dummy `_hmerge`/`_vmerge` cells so they can
	 * inherit the origin's border/fill and render the merged region's outer edges (Issue #680) */
	_spanOrigin?: TableCell
	_optImp?: any

	text?: string | number | TableCell[] // TODO: FUTURE: 20210815: ONly allow `TableCell[]` dealing with string|TableCell[] *SUCKS*
	options?: TableCellProps
}
export interface TableRowSlide {
	rows: TableRow[]
	/**
	 * Per-row height (inches) aligned 1:1 with `rows`, derived from the original `rowH` array.
	 * Auto-paging splits/reorders rows across slides and inserts repeated headers, so the caller's
	 * `rowH[i]` (keyed by *original* row index) can no longer be applied by physical row index on
	 * each generated slide. This carries each output row's resolved height so a configured height
	 * follows its source row instead of being re-applied to whatever lands at that index (#1145).
	 * Entries are `undefined` where no explicit height was configured (auto-distributed height).
	 */
	rowH?: Array<number | undefined>
}
export type TableRow = TableCell[]

// text ===========================================================================================
export interface TextGlowProps {
	/**
	 * Border color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * opacity (0.0 - 1.0)
	 * @example 0.5
	 * 50% opaque
	 */
	opacity?: number
	/**
	 * size (points)
	 */
	size: number
}

export interface TextFitShrinkProps {
	/**
	 * Shrink text on overflow (`<a:normAutofit>`)
	 */
	type: 'shrink'
	/**
	 * Font scale as a percent (0-100), mapped to `<a:normAutofit fontScale="..">`.
	 *
	 * PowerPoint normally calculates this dynamically when text overflows; set it
	 * explicitly to bake the scale into the generated file.
	 * @example 85 // render text at 85% of its nominal size
	 * @default undefined // attribute omitted (PowerPoint defaults to 100%)
	 */
	fontScale?: number
	/**
	 * Line-space reduction as a percent (0-100), mapped to `<a:normAutofit lnSpcReduction="..">`.
	 * @example 20 // reduce line spacing by 20%
	 * @default undefined // attribute omitted (PowerPoint defaults to 0%)
	 */
	lnSpcReduction?: number
}

export interface TextPropsOptions extends PositionProps, DataOrPathProps, TextBaseProps, ObjectNameProps {
	_bodyProp?: {
		// Note: Many of these duplicated as user options are transformed to _bodyProp options for XML processing
		autoFit?: boolean
		align?: TEXT_HALIGN
		anchor?: TEXT_VALIGN
		lIns?: number
		rIns?: number
		tIns?: number
		bIns?: number
		numCol?: number
		spcCol?: number
		vert?: TextVertType
		wrap?: boolean
		prstTxWarp?: string
	}
	_lineIdx?: number

	baseline?: number
	/**
	 * Character spacing
	 */
	charSpacing?: number
	/**
	 * Number of text columns in the text body
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Columns > "Number"
	 * - range: 1-16
	 * @since v5.3.0
	 * @default 1
	 * @example 2 // flow text into two columns
	 */
	columns?: number
	/**
	 * Spacing between text columns (points)
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Columns > "Spacing"
	 * - only applies when `columns` > 1
	 * @since v5.3.0
	 * @default 0
	 * @example 10 // 10pt gap between columns
	 */
	columnSpacing?: number
	/**
	 * Text fit options
	 *
	 * MS-PPT > Format Shape > Shape Options > Text Box > "[unlabeled group]": [3 options below]
	 * - 'none' = Do not Autofit
	 * - 'shrink' = Shrink text on overflow
	 * - 'resize' = Resize shape to fit text
	 *
	 * **Measured fit:** if you register the box's font with
	 * {@link PptxGenJS.registerFontMetrics}, both `'shrink'` and `'resize'` are
	 * **measured at export time**, so the text renders correctly in headless renderers
	 * and on plain file-open (no edit/resize needed):
	 * - `'shrink'` computes the largest `fontScale` at which the wrapped text fits and
	 *   bakes `<a:normAutofit fontScale=…/>`.
	 * - `'resize'` computes the height the text needs and bakes it into the shape's
	 *   `a:ext/@cy` (adjusting `a:off/@y` per vertical anchor), the marker being
	 *   `<a:spAutoFit/>`.
	 * Without registered metrics they fall back to the bare flag (`<a:normAutofit/>` /
	 * `<a:spAutoFit/>`, which only PowerPoint recomputes on edit) and warn once.
	 *
	 * **Note** Bare `'shrink'`/`'resize'` (no metrics) only take effect after editing
	 * text / resizing the shape; PowerPoint calculates the result then. The object form
	 * of `'shrink'` always bakes the explicit values you pass.
	 * @since v3.3.0
	 * @example 'shrink' // measured when metrics are registered; else bare <a:normAutofit/>
	 * @example 'resize' // measured when metrics are registered; else bare <a:spAutoFit/>
	 * @example { type: 'shrink', fontScale: 85, lnSpcReduction: 20 } // pre-shrink with explicit values
	 * @default "none"
	 */
	fit?: 'none' | 'shrink' | 'resize' | TextFitShrinkProps
	/**
	 * Shape fill
	 * @example { color:'FF0000' } // hex color (red)
	 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
	 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
	 */
	fill?: ShapeFillProps
	/**
	 * Flip shape horizontally?
	 * @default false
	 */
	flipH?: boolean
	/**
	 * Flip shape vertical?
	 * @default false
	 */
	flipV?: boolean
	glow?: TextGlowProps
	hyperlink?: HyperlinkProps
	indentLevel?: number
	isTextBox?: boolean
	line?: ShapeLineProps
	/**
	 * Line spacing (pt)
	 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Exactly"
	 * @example 28 // 28pt
	 */
	lineSpacing?: number
	/**
	 * line spacing multiple (percent)
	 * - range: 0.0-9.99
	 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Multiple"
	 * @example 1.5 // 1.5X line spacing
	 * @since v3.5.0
	 */
	lineSpacingMultiple?: number
	// TODO: [20220219] powerpoint uses inches but library has always been pt... @future @deprecated - update in v4.0? [range: 0.0-22.0]
	/**
	 * Margin (points)
	 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Left/Right/Top/Bottom margin
	 * @default "Normal" margin in PowerPoint [3.5, 7.0, 3.5, 7.0] // (this library sets no value, but PowerPoint defaults to "Normal" [0.05", 0.1", 0.05", 0.1"])
	 * @example 0 // Top/Right/Bottom/Left margin 0 [0.0" in powerpoint]
	 * @example 10 // Top/Right/Bottom/Left margin 10 [0.14" in powerpoint]
	 * @example [10,5,10,5] // Top margin 10, Right margin 5, Bottom margin 10, Left margin 5
	 */
	margin?: Margin
	outline?: { color: Color, size: number }
	paraSpaceAfter?: number
	paraSpaceBefore?: number
	/**
	 * Placeholder type
	 * - when the value matches a placeholder defined on the slide layout/master, this text
	 *   inherits that placeholder's position and formatting
	 * - otherwise the text shape is promoted to a standalone placeholder of this type, emitting
	 *   a real `<p:ph type="...">`. Use `placeholder: 'title'` to give a slide an accessible
	 *   title (PowerPoint's accessibility checker otherwise reports "Missing Slide Title")
	 * - values: 'title' | 'body' | et. al.
	 * @example 'title'
	 * @see https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
	 */
	placeholder?: string
	/**
	 * Rounded rectangle radius (only for pptx.shapes.ROUNDED_RECTANGLE)
	 * - values: 0.0 to 1.0
	 * @default 0
	 */
	rectRadius?: number
	/**
	 * Rotation (degrees)
	 * - range: -360 to 360
	 * @default 0
	 * @example 180 // rotate 180 degrees
	 */
	rotate?: number
	/**
	 * Whether to enable right-to-left mode
	 * @default false
	 */
	rtlMode?: boolean
	shadow?: ShadowProps
	shape?: SHAPE_NAME
	strike?: boolean | 'dblStrike' | 'sngStrike'
	subscript?: boolean
	superscript?: boolean
	/**
	 * Vertical alignment
	 * @default middle
	 */
	valign?: VAlign
	vert?: TextVertType
	/**
	 * Text wrap
	 * @since v3.3.0
	 * @default true
	 */
	wrap?: boolean

	/**
	 * Whether "Fit to Shape?" is enabled
	 * @deprecated v3.3.0 - use `fit`
	 */
	autoFit?: boolean
	/**
	 * Whather "Shrink Text on Overflow?" is enabled
	 * @deprecated v3.3.0 - use `fit`
	 */
	shrinkText?: boolean
	/**
	 * Inset
	 * @deprecated v3.10.0 - use `margin`
	 */
	inset?: number
	/**
	 * Dash type
	 * @deprecated v3.3.0 - use `line.dashType`
	 */
	lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
	/**
	 * @deprecated v3.3.0 - use `line.beginArrowType`
	 */
	lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	/**
	 * @deprecated v3.3.0 - use `line.width`
	 */
	lineSize?: number
	/**
	 * @deprecated v3.3.0 - use `line.endArrowType`
	 */
	lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
}
export interface TextProps {
	text?: string | number
	options?: TextPropsOptions
	/**
	 * Raw OMML (Office MathML) for a native, editable PowerPoint equation, emitted as its own
	 * display-math paragraph (`<a14:m><m:oMathPara><m:oMath>…`). When set, this item is a math
	 * paragraph and `text` is ignored. Accepts either the inner OMML (children of `<m:oMath>`),
	 * a full `<m:oMath>…</m:oMath>`, or a full `<m:oMathPara>…</m:oMathPara>`; the `m:` prefix is
	 * resolved by the wrapper, so the markup does not need its own namespace declarations.
	 * This is the raw-OMML entry point; LaTeX/MathML→OMML conversion is not yet provided.
	 * @example { math: '<m:r><m:t>x^2+1=y</m:t></m:r>' } // (use real OMML, not LaTeX)
	 * @since v5.4.0
	 */
	math?: string
}

/**
 * Options for layout-time text measurement ({@link PptxGenJS.measureText}).
 * Inches for width, points for type/spacing — the consumer-facing units. The
 * measured face must have metrics registered via {@link PptxGenJS.registerFontMetrics}
 * (a named face without exact metrics uses a conservative heuristic; an unnamed
 * theme-default face is unmeasurable).
 * @since v6.1.0
 */
export interface MeasureTextOptions {
	/** Available text width in inches (the box width minus L/R inset, unless `insetIn` is given). */
	wIn: number
	/** Font size in points. */
	fontSize: number
	/** Font family name, as used in `fontFace`. Required for an exact measure; an unnamed face is unmeasurable. */
	fontFace?: string
	bold?: boolean
	italic?: boolean
	/** Character spacing in points. */
	charSpacing?: number
	/** Exact line spacing in points (overrides `lineSpacingMultiple`). */
	lineSpacing?: number
	/** Line spacing as a multiple of single (e.g. `1.5`). */
	lineSpacingMultiple?: number
	/** Space before each paragraph, in points. */
	paraSpaceBefore?: number
	/** Space after each paragraph, in points. */
	paraSpaceAfter?: number
	/** L/R text inset in inches; when set, subtracted from `wIn` on both sides (pass a raw box width). */
	insetIn?: number
}

/**
 * Result of {@link PptxGenJS.measureText}. Heights err **tall** (conservative) —
 * they match the value the export-time autofit bake uses, so the laid-out height is
 * ≥ what PowerPoint/LibreOffice render. Use it to grow a container; for an overflow
 * check it may slightly over-report (good for a warning, not a hard gate).
 * @since v6.1.0
 */
export interface TextMeasurement {
	/** Laid-out height in inches at the given `fontSize` (conservative/tall). */
	heightIn: number
	/** Number of wrapped lines (conservative — the model wraps marginally early). */
	lineCount: number
	/**
	 * Width in inches of the widest laid-out line (conservative — the model wraps
	 * marginally early, so this errs slightly wide). With an unconstrained `wIn` it
	 * is the natural single-line width; constrained, it is the widest wrapped line.
	 * A box set to this width will not re-wrap the text.
	 * @since v7.0.0
	 */
	widestLineIn: number
	/** `false` only for an unnamed theme-default face that could not be measured. */
	measurable: boolean
	/** True if the text fits a box of inner height `hIn` (inches) at full size. */
	fitsBox: (hIn: number) => boolean
	/** The `fontScale` (percent) that fits inner height `hIn`; `100` if it already fits, never below the shrink floor. */
	shrinkScaleFor: (hIn: number) => number
}

/** Options for {@link PptxGenJS.overflowsBox}: a measure plus the box inner height to test against. @since v6.1.0 */
export interface OverflowBoxOptions extends MeasureTextOptions {
	/** Box inner height in inches to test for overflow. */
	hIn: number
}

/**
 * One cell's computed rectangle from {@link PptxGenJS.tableLayout}. All values are
 * inches; `x`/`y` are absolute (offset from the table's `x`/`y`). For a merged cell,
 * `row`/`col` are the top-left origin and `wIn`/`hIn` cover the whole span; the
 * cells it covers are not emitted separately.
 * @since v7.1.0
 */
export interface TableCellLayout {
	/** Zero-based grid row of the cell's top-left origin. */
	row: number
	/** Zero-based grid column of the cell's top-left origin. */
	col: number
	/** Number of rows the cell spans (1 if not merged). */
	rowSpan: number
	/** Number of columns the cell spans (1 if not merged). */
	colSpan: number
	/** Left edge in inches (absolute). */
	xIn: number
	/** Top edge in inches (absolute). */
	yIn: number
	/** Outer cell width in inches (sum of spanned column widths). */
	wIn: number
	/** Outer cell height in inches (sum of spanned row heights). */
	hIn: number
	/**
	 * `true` when `hIn`/`yIn` are pinned by an explicit `rowH` (array or scalar) or
	 * table `h`; `false` when the row is auto-height and the value is a conservative
	 * (tall) estimate from the same text model as {@link PptxGenJS.measureText}.
	 */
	heightExact: boolean
}

/**
 * Result of {@link PptxGenJS.tableLayout}: per-cell geometry plus overall table
 * bounds, for placing images/shapes over a table without rendering it. Geometry is
 * for a single, un-paginated table laid out at `opts.x`/`y`/`w`; `autoPage` paging
 * is not modeled. Widths are exact; auto-height row heights are conservative
 * estimates (see {@link TableCellLayout.heightExact}).
 * @since v7.1.0
 */
export interface TableLayoutResult {
	/** One entry per non-merged origin cell, in row-major order. */
	cells: TableCellLayout[]
	/** Overall table width in inches (sum of column widths). */
	widthIn: number
	/** Overall table height in inches (sum of row heights; may include estimates). */
	heightIn: number
	/** `false` if any row height was estimated (the total errs tall, like `measureText`). */
	heightExact: boolean
}

/**
 * Per-run options for a speaker-notes text run.
 * A focused subset of `TextPropsOptions`: inline formatting plus an (external URL) hyperlink.
 * Notes hyperlinks support `url` only; `slide` targets are not yet supported.
 */
export type NotesTextOptions = Pick<TextPropsOptions, 'hyperlink' | 'bold' | 'italic' | 'underline' | 'color' | 'fontSize' | 'fontFace'>

/** A single speaker-notes text run: text plus optional inline formatting / hyperlink. */
export interface NotesProps {
	text: string
	options?: NotesTextOptions
}

/** Factory for a single inline text run. Prevents `as never` casts when building mixed-style run arrays. */
export function textRun(text: string | number, options?: TextPropsOptions): TextProps {
	return options !== undefined ? { text, options } : { text }
}

/** Wraps a run array so TypeScript accepts it as `TextProps[]` without a cast. */
export function textRuns(runs: TextProps[]): TextProps[] {
	return runs
}

// charts =========================================================================================
// FUTURE: BREAKING-CHANGE: (soln: use `OptsDataLabelPosition|string` until 3.5/4.0)
/*
export interface OptsDataLabelPosition {
	pie: 'ctr' | 'inEnd' | 'outEnd' | 'bestFit'
	scatter: 'b' | 'ctr' | 'l' | 'r' | 't'
	// TODO: add all othere chart types
}
*/

export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
/**
 * Line end cap style. Maps to the OOXML `cap` attribute on `<a:ln>` (`flat`/`sq`/`rnd`).
 */
export type LineCap = 'flat' | 'round' | 'square'
/** @deprecated use `LineCap` (the cap type is not chart-specific) */
export type ChartLineCap = LineCap
export type ChartLineDash = 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'

export interface OptsChartData {
	_dataIndex?: number

	/**
	 * category labels
	 * @example ['Year 2000', 'Year 2010', 'Year 2020'] // single-level category axes labels
	 * @example [['Year 2000', 'Year 2010', 'Year 2020'], ['Decades', '', '']] // multi-level category axes labels
	 * @since `labels` string[][] type added v3.11.0
	 */
	labels?: string[] | string[][]
	/**
	 * series name
	 * @example 'Locations'
	 */
	name?: string
	/**
	 * bubble sizes
	 * @example [5, 1, 5, 1]
	 */
	sizes?: number[]
	/**
	 * category values
	 * @example [2000, 2010, 2020]
	 */
	values?: number[]
	/**
	 * Custom text label per data point, replacing the auto-generated value label.
	 * Index aligns with `values[]`. Empty string or missing entries fall back to the chart-level label settings.
	 * Supported for BAR, LINE, AREA, RADAR, PIE, and DOUGHNUT chart types.
	 * @example ['Low', '', 'High']  // only points 0 and 2 get custom labels
	 */
	customLabels?: string[]
	/**
	 * Per-data-point visual overrides (border / fill), index-aligned with `values[]`.
	 * Empty (`{}`) or missing entries fall back to series/chart styling.
	 * Supported for BAR, LINE, AREA, SCATTER, PIE, and DOUGHNUT chart types.
	 * @example
	 * pointStyles: [
	 *   { border: { pt: 2, color: 'FF0000' } }, // point 0: red 2pt border
	 *   {},                                     // point 1: default
	 *   { fill: '00B050', border: { type: 'dash', color: '404040' } }, // point 2
	 * ]
	 * @since v5.3.0
	 */
	pointStyles?: ChartDataPointStyle[]
	/**
	 * Error bars for this series (`<c:errBars>`).
	 * - Supported for BAR, BAR3D, LINE, AREA, and SCATTER chart types (RADAR has no error bars in the schema).
	 * - Pass a single config, or an array to draw both X and Y error bars (SCATTER/AREA only; BAR/LINE use the first entry).
	 * @example { valueType: 'percentage', value: 5 } // ±5% error bars
	 * @example { valueType: 'fixedVal', value: 2, barType: 'plus', noEndCap: true }
	 * @example { valueType: 'cust', plusValues: [1, 2, 1], minusValues: [0.5, 1, 0.5] }
	 * @since v6.0.0
	 */
	errorBars?: ChartErrorBarOptions | ChartErrorBarOptions[]
	/**
	 * Override `chartColors`
	 */
	// color?: string // TODO: WIP: (Pull #727)
}
/**
 * Per-data-point style override for a chart series.
 * Each entry applies to the data point at the same index in `values[]`.
 * Unset fields fall back to the series/chart-level styling.
 */
export interface ChartDataPointStyle {
	/**
	 * Data-point border (line). Reuses {@link BorderProps}.
	 * - `type: 'none'` hides the border; `'dash'` draws a dashed border.
	 * @example { pt: 2, color: 'FF0000' }
	 */
	border?: BorderProps
	/**
	 * Data-point fill color (hex), overriding `chartColors[idx]`.
	 * Most meaningful on fill-based charts (BAR, AREA, PIE, DOUGHNUT).
	 * @example '00B050'
	 */
	fill?: HexColor
	/**
	 * Data-point pattern fill (`<a:pattFill>`), e.g. diagonal hatching, for the
	 * BAR/BAR3D and SCATTER charts that emit per-point `c:dPt`. Takes precedence
	 * over `fill` (OOXML allows only one fill per data point).
	 *
	 * When `pattern.fgColor` is omitted it defaults to this point's resolved fill
	 * color (`fill` or the varied `chartColors[idx]`), giving a hatched version of
	 * the bar color; if no point color is resolvable it falls back to black.
	 * `pattern.bgColor` defaults to white.
	 * @example { preset: 'ltUpDiag' }
	 * @example { preset: 'diagCross', fgColor: 'C00000', bgColor: 'FFFFFF' }
	 */
	pattern?: PatternFillProps
}
/**
 * Error-bar configuration for a chart series (`<c:errBars>`).
 * Maps onto OOXML `CT_ErrBars` (errDir / errBarType / errValType / noEndCap / plus / minus / val).
 */
export interface ChartErrorBarOptions {
	/**
	 * Axis the error bars measure along.
	 * - `'y'` (the value axis) for BAR/BAR3D/LINE/AREA; SCATTER may also use `'x'`.
	 * @default 'y'
	 */
	direction?: 'x' | 'y'
	/**
	 * Which sides of each marker draw a bar.
	 * @default 'both'
	 */
	barType?: 'both' | 'minus' | 'plus'
	/**
	 * How `value` (or `plusValues`/`minusValues`) is interpreted.
	 * - `'fixedVal'` — fixed amount in axis units
	 * - `'percentage'` — percent of each value (e.g. `value: 5` → ±5%)
	 * - `'stdDev'` — `value` standard deviations
	 * - `'stdErr'` — standard error (ignores `value`)
	 * - `'cust'` — explicit per-point amounts via `plusValues`/`minusValues`
	 * @default 'fixedVal'
	 */
	valueType?: 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr'
	/**
	 * Magnitude for `'fixedVal'`, `'percentage'`, or `'stdDev'`. Ignored for `'stdErr'` and `'cust'`.
	 * @default 1
	 */
	value?: number
	/** Per-point positive magnitudes; required when `valueType === 'cust'` (unless `barType: 'minus'`). Index-aligned with `values[]`. */
	plusValues?: number[]
	/** Per-point negative magnitudes; required when `valueType === 'cust'` (unless `barType: 'plus'`). Index-aligned with `values[]`. */
	minusValues?: number[]
	/**
	 * Hide the perpendicular end caps.
	 * @default false
	 */
	noEndCap?: boolean
	/** Error-bar line color (hex, e.g. `'FF0000'`). */
	color?: HexColor
	/** Error-bar line width (points). */
	size?: number
}
// Used internally, probably shouldn't be used by end users
export interface IOptsChartData extends OptsChartData {
	labels?: string[][]
}
export interface OptsChartGridLine {
	/**
	 * MS-PPT > Chart format > Format Major Gridlines > Line > Cap type
	 * - line cap type
	 * @default flat
	 */
	cap?: ChartLineCap
	/**
	 * Gridline color (hex)
	 * @example 'FF3399'
	 */
	color?: HexColor
	/**
	 * Gridline size (points)
	 */
	size?: number
	/**
	 * Gridline style
	 */
	style?: 'solid' | 'dash' | 'dot' | 'none'
}
// TODO: 202008: chart types remain with predicated with "I" in v3.3.0 (ran out of time!)
export interface IChartMulti {
	type: CHART_NAME
	data: OptsChartData[]
	options: IChartOptsLib
}
export interface IChartPropsFillLine {
	/**
	 * PowerPoint: Format Chart Area/Plot > Border ["Line"]
	 * @example border: {color: 'FF0000', pt: 1} // hex RGB color, 1 pt line
	 */
	border?: BorderProps
	/**
	 * PowerPoint: Format Chart Area/Plot Area > Fill
	 * @example fill: {color: '696969'} // hex RGB color value
	 * @example fill: {color: pptx.SchemeColor.background2} // Theme color value
	 * @example fill: {transparency: 50} // 50% transparency
	 */
	fill?: ShapeFillProps
}
export interface IChartAreaProps extends IChartPropsFillLine {
	/**
	 * Whether the chart area has rounded corners
	 * - only applies when either `fill` or `border` is used
	 * @default true
	 * @since v3.11
	 */
	roundedCorners?: boolean
}
export interface IChartPropsBase {
	/**
	 * Axis position
	 */
	axisPos?: 'b' | 'l' | 'r' | 't'
	chartColors?: HexColor[]
	/**
	 * opacity (0 - 100)
	 * @example 50 // 50% opaque
	 */
	chartColorsOpacity?: number
	dataBorder?: BorderProps
	displayBlanksAs?: 'gap' | 'span' | 'zero'
	invertedColors?: HexColor[]
	lang?: string
	layout?: PositionProps
	shadow?: ShadowProps
	/**
	 * Show each bubble's size value as a data label (bubble / bubble3D charts only).
	 * Has no effect on other chart types.
	 * @default false
	 */
	showBubbleSize?: boolean
	/**
	 * @default false
	 */
	showLabel?: boolean
	showLeaderLines?: boolean
	/**
	 * Leader line color (pie/doughnut data labels). Requires `showLeaderLines: true`.
	 * When omitted, PowerPoint applies its automatic leader-line color.
	 * @example 'FF0000' // red leader lines
	 */
	leaderLineColor?: HexColor
	/**
	 * Leader line width, in points (pie/doughnut data labels). Requires `showLeaderLines: true`.
	 * @default 0.75
	 * @example 1.5
	 */
	leaderLineSize?: number
	/**
	 * @default false
	 */
	showLegend?: boolean
	/**
	 * @default false
	 */
	showPercent?: boolean
	/**
	 * @default false
	 */
	showSerName?: boolean
	/**
	 * @default false
	 */
	showTitle?: boolean
	/**
	 * @default false
	 */
	showValue?: boolean
	/**
	 * 3D Perspecitve
	 * - range: 0-120
	 * @default 30
	 */
	v3DPerspective?: number
	/**
	 * Right Angle Axes
	 * - Shows chart from first-person perspective
	 * - Overrides `v3DPerspective` when true
	 * - PowerPoint: Chart Options > 3-D Rotation
	 * @default false
	 */
	v3DRAngAx?: boolean
	/**
	 * X Rotation
	 * - PowerPoint: Chart Options > 3-D Rotation
	 * - range: 0-359.9
	 * @default 30
	 */
	v3DRotX?: number
	/**
	 * Y Rotation
	 * - range: 0-359.9
	 * @default 30
	 */
	v3DRotY?: number

	/**
	 * PowerPoint: Format Chart Area (Fill & Border/Line)
	 * @since v3.11
	 */
	chartArea?: IChartAreaProps
	/**
	 * PowerPoint: Format Plot Area (Fill & Border/Line)
	 * @since v3.11
	 */
	plotArea?: IChartPropsFillLine

	/**
	 * @deprecated v3.11.0 - use `plotArea.border`
	 */
	border?: BorderProps
	/**
	 * @deprecated v3.11.0 - use `plotArea.fill`
	 */
	fill?: HexColor
	/**
	 * Per-series style overrides.
	 * Element at index N applies to the series at data[N].
	 * Missing indices or unset fields fall back to the chart-level option.
	 * @since v4.0.0
	 */
	seriesOptions?: IChartSeriesOpts[]
}
export interface IChartPropsAxisCat {
	/**
	 * Multi-Chart prop: array of cat axes
	 */
	catAxes?: IChartPropsAxisCat[]
	catAxisBaseTimeUnit?: string
	catAxisCrossesAt?: number | 'autoZero'
	catAxisHidden?: boolean
	catAxisLabelColor?: string
	catAxisLabelFontBold?: boolean
	catAxisLabelFontFace?: string
	catAxisLabelFontItalic?: boolean
	catAxisLabelFontSize?: number
	/**
	 * Number format code for the category (X) axis labels on scatter and bubble charts.
	 * Falls back to `valAxisLabelFormatCode` when not set.
	 * - Example: `'0.00'`, `'#,##0'`, `'mmm yyyy'`
	 * - PowerPoint: Format Axis > Number > Format Code
	 */
	catAxisLabelFormatCode?: string
	catAxisLabelFrequency?: string
	catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	catAxisLabelRotate?: number
	catAxisLineColor?: string
	catAxisLineShow?: boolean
	catAxisLineSize?: number
	catAxisLineStyle?: 'solid' | 'dash' | 'dot'
	catAxisMajorTickMark?: ChartAxisTickMark
	catAxisMajorTimeUnit?: string
	catAxisMajorUnit?: number
	catAxisMaxVal?: number
	catAxisMinorTickMark?: ChartAxisTickMark
	catAxisMinorTimeUnit?: string
	catAxisMinorUnit?: number
	catAxisMinVal?: number
	/** @since v3.11.0 */
	catAxisMultiLevelLabels?: boolean
	catAxisOrientation?: 'minMax' | 'maxMin'
	catAxisTitle?: string
	catAxisTitleColor?: string
	catAxisTitleFontFace?: string
	catAxisTitleFontSize?: number
	catAxisTitleRotate?: number
	catGridLine?: OptsChartGridLine
	catLabelFormatCode?: string
	/**
	 * Whether data should use secondary category axis (instead of primary)
	 * @default false
	 */
	secondaryCatAxis?: boolean
	showCatAxisTitle?: boolean
}
export interface IChartPropsAxisSer {
	serAxisBaseTimeUnit?: string
	serAxisHidden?: boolean
	serAxisLabelColor?: string
	serAxisLabelFontBold?: boolean
	serAxisLabelFontFace?: string
	serAxisLabelFontItalic?: boolean
	serAxisLabelFontSize?: number
	serAxisLabelFrequency?: string
	serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	serAxisLineColor?: string
	serAxisLineShow?: boolean
	serAxisMajorTimeUnit?: string
	serAxisMajorUnit?: number
	serAxisMinorTimeUnit?: string
	serAxisMinorUnit?: number
	serAxisOrientation?: string
	serAxisTitle?: string
	serAxisTitleColor?: string
	serAxisTitleFontFace?: string
	serAxisTitleFontSize?: number
	serAxisTitleRotate?: number
	serGridLine?: OptsChartGridLine
	serLabelFormatCode?: string
	showSerAxisTitle?: boolean
}
export interface IChartPropsAxisVal {
	/**
	 * Whether data should use secondary value axis (instead of primary)
	 * @default false
	 */
	secondaryValAxis?: boolean
	showValAxisTitle?: boolean
	/**
	 * Multi-Chart prop: array of val axes
	 */
	valAxes?: IChartPropsAxisVal[]
	valAxisCrossesAt?: number | 'autoZero'
	/**
	 * Controls where axis values are plotted relative to tick marks
	 * - `'between'` = values plotted between tick marks (default for bar/column/line)
	 * - `'midCat'` = values plotted on tick marks (default for scatter/area)
	 * - PowerPoint: Format Axis > Axis Options > Axis crosses > On tick marks / Between tick marks
	 */
	valAxisCrossBetween?: 'between' | 'midCat'
	valAxisDisplayUnit?: 'billions' | 'hundredMillions' | 'hundreds' | 'hundredThousands' | 'millions' | 'tenMillions' | 'tenThousands' | 'thousands' | 'trillions'
	valAxisDisplayUnitLabel?: boolean
	valAxisHidden?: boolean
	valAxisLabelColor?: string
	valAxisLabelFontBold?: boolean
	valAxisLabelFontFace?: string
	valAxisLabelFontItalic?: boolean
	valAxisLabelFontSize?: number
	valAxisLabelFormatCode?: string
	valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	valAxisLabelRotate?: number
	valAxisLineColor?: string
	valAxisLineShow?: boolean
	valAxisLineSize?: number
	valAxisLineStyle?: 'solid' | 'dash' | 'dot'
	/**
	 * PowerPoint: Format Axis > Axis Options > Logarithmic scale - Base
	 * - range: 2-99
	 * @since v3.5.0
	 */
	valAxisLogScaleBase?: number
	valAxisMajorTickMark?: ChartAxisTickMark
	valAxisMajorUnit?: number
	valAxisMaxVal?: number
	valAxisMinorTickMark?: ChartAxisTickMark
	valAxisMinVal?: number
	valAxisOrientation?: 'minMax' | 'maxMin'
	valAxisTitle?: string
	valAxisTitleColor?: string
	valAxisTitleFontFace?: string
	valAxisTitleFontSize?: number
	valAxisTitleRotate?: number
	valGridLine?: OptsChartGridLine
	/**
	 * Value label format code
	 * - this also directs Data Table formatting
	 * @since v3.3.0
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	valLabelFormatCode?: string
}
export interface IChartPropsChartBar {
	bar3DShape?: string
	barDir?: string
	barGapDepthPct?: number
	/**
	 * MS-PPT > Format chart > Format Data Point > Series Options >  "Gap Width"
	 * - width (percent)
	 * - range: `0`-`500`
	 * @default 150
	 */
	barGapWidthPct?: number
	barGrouping?: string
	/**
	 * MS-PPT > Format chart > Format Data Point > Series Options >  "Series Overlap"
	 * - overlap (percent)
	 * - range: `-100`-`100`
	 * @since v3.9.0
	 * @default 0
	 */
	barOverlapPct?: number
	/**
	 * Draw connector lines between data points across stacked bar/column series
	 * ("Series Lines" in PowerPoint). Emits `<c:serLines>` in the bar chart.
	 *
	 * - `true` uses PowerPoint's automatic line styling.
	 * - An {@link OptsChartGridLine} object customizes color/size/style/cap.
	 * - Omit (or pass an object with `style: 'none'`) to disable.
	 *
	 * Bar (`bar`) charts only; ignored for 3D bar charts.
	 * @default undefined
	 * @example true
	 * @example { color: '777777', size: 1, style: 'dash' }
	 */
	barSeriesLine?: boolean | OptsChartGridLine
}
export interface IChartPropsChartDoughnut {
	dataNoEffects?: boolean
	holeSize?: number
}
export interface IChartPropsChartLine {
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Cap type
	 * - line cap type
	 * @default flat
	 */
	lineCap?: ChartLineCap
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Dash type (chart-level default)
	 * - applies to every series that has no entry in `lineDashValues`
	 * @default solid
	 */
	lineDash?: ChartLineDash
	/**
	 * Per-series dash type overrides; index matches the series order in the `data` array.
	 * - entries shorter than the series count fall back to `lineDash`
	 * - example: `['solid', 'dash', 'lgDash']` gives each series its own dash pattern
	 */
	lineDashValues?: ChartLineDash[]
	/**
	 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Type
	 * - marker type
	 * @default circle
	 */
	lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
	/**
	 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Color
	 * - border color
	 * @default circle
	 */
	lineDataSymbolLineColor?: string
	/**
	 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Width
	 * - border width (points)
	 * @default 0.75
	 */
	lineDataSymbolLineSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Size
	 * - marker size
	 * - range: 2-72
	 * @default 6
	 */
	lineDataSymbolSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Width
	 * - line width (points)
	 * - range: 0-1584
	 * @default 2
	 */
	lineSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Smoothed line
	 * - "Smoothed line"
	 * @default false
	 */
	lineSmooth?: boolean
}
export interface IChartPropsChartPie {
	dataNoEffects?: boolean
	/**
	 * MS-PPT > Format chart > Format Data Series > Series Options >  "Angle of first slice"
	 * - angle (degrees)
	 * - range: 0-359
	 * @since v3.4.0
	 * @default 0
	 */
	firstSliceAng?: number
}
export interface IChartPropsChartRadar {
	/**
	 * MS-PPT > Chart Type > Waterfall
	 * - radar chart type
	 * @default standard
	 */
	radarStyle?: 'standard' | 'marker' | 'filled' // TODO: convert to 'radar'|'markers'|'filled' in 4.0 (verbatim with PPT app UI)
}
/**
 * Per-series style overrides for a chart.
 * Each entry applies to the series at the same index in the data array.
 * Unset fields fall back to the chart-level option.
 */
export interface IChartSeriesOpts {
	/** Series fill / line color (hex, e.g. `'FF0000'`) */
	color?: HexColor
	/** Data-label font color */
	dataLabelColor?: string
	/** Data-label font bold */
	dataLabelFontBold?: boolean
	/** Data-label typeface */
	dataLabelFontFace?: string
	/** Data-label font italic */
	dataLabelFontItalic?: boolean
	/** Data-label font size (points) */
	dataLabelFontSize?: number
	/**
	 * Data-label number format code for this series.
	 * Overrides the chart-level `dataLabelFormatCode` for this series only.
	 * @example '#,##0' // thousands separator
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataLabelFormatCode?: string
	/**
	 * Line/radar series line width (points).
	 * Pass `0` to hide the line.
	 */
	lineSize?: number
}

export interface IChartPropsDataLabel {
	dataLabelBkgrdColors?: boolean
	dataLabelColor?: string
	dataLabelFontBold?: boolean
	dataLabelFontFace?: string
	dataLabelFontItalic?: boolean
	dataLabelFontSize?: number
	/**
	 * Data label format code
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataLabelFormatCode?: string
	dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
	dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd'
}
export interface IChartPropsDataTable {
	dataTableFontSize?: number
	/**
	 * Data table format code
	 * @since v3.3.0
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataTableFormatCode?: string
	/**
	 * Whether to show a data table adjacent to the chart
	 * @default false
	 */
	showDataTable?: boolean
	showDataTableHorzBorder?: boolean
	showDataTableKeys?: boolean
	showDataTableOutline?: boolean
	showDataTableVertBorder?: boolean
}
export interface IChartPropsLegend {
	legendColor?: string
	legendFontFace?: string
	legendFontSize?: number
	/**
	 * Manual legend placement within the chart area.
	 *
	 * Each of `x`/`y`/`w`/`h` is a fraction (0-1) of the chart's width/height.
	 * `x`/`y` position the legend's top-left corner relative to the chart edge;
	 * `w`/`h` set its size. Each axis is independent: provide only `x` to move the
	 * legend horizontally while leaving vertical placement and size automatic.
	 * Setting this overrides the automatic placement implied by `legendPos`.
	 *
	 * Has no effect unless `showLegend` is `true`.
	 *
	 * @example { x: 0.7, y: 0.3, w: 0.25, h: 0.4 }
	 */
	legendLayout?: PositionProps
	legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
}
export interface IChartPropsTitle extends TextBaseProps {
	title?: string
	titleAlign?: string
	titleBold?: boolean
	titleColor?: string
	titleFontFace?: string
	titleFontSize?: number
	titleItalic?: boolean
	titleUnderline?: boolean
	/**
	 * Manual title position (inches), relative to the chart.
	 * Each axis is independent: omit `x` to keep automatic horizontal centering,
	 * or omit `y` to keep automatic vertical placement. Provide at least one.
	 */
	titlePos?: { x?: number, y?: number }
	titleRotate?: number
}
export interface IChartOpts
	extends IChartPropsAxisCat,
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
	IChartPropsLegend,
	IChartPropsTitle,
	ObjectNameProps,
	OptsChartGridLine,
	PositionProps {
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - PowerPoint: [right-click on a chart] > "Edit Alt Text..."
	 */
	altText?: string
}
export interface IChartOptsLib extends IChartOpts {
	_type?: CHART_NAME | IChartMulti[] // TODO: v3.4.0 - move to `IChartOpts`, remove `IChartOptsLib`
}
export interface ISlideRelChart extends OptsChartData {
	type: CHART_NAME | IChartMulti[]
	opts: IChartOptsLib
	data: IOptsChartData[]
	// internal below
	rId: number
	Target: string
	globalId: number
	fileName: string
}

// Core
// ====
// PRIVATE vvv
export interface ISlideRel {
	type: SLIDE_OBJECT_TYPES
	Target: string
	fileName?: string
	data: any[] | string
	opts?: IChartOpts
	path?: string
	extn?: string
	globalId?: number
	rId: number
}
export interface ISlideRelMedia {
	type: string
	opts?: MediaProps
	path?: string
	extn?: string
	data?: string | ArrayBuffer
	/** used to indicate that a media file has already been read/enocded (PERF) */
	isDuplicate?: boolean
	isSvgPng?: boolean
	svgSize?: { w: number, h: number }
	rId: number
	Target: string
}
export interface ISlideObject {
	_type: SLIDE_OBJECT_TYPES
	options?: ObjectOptions
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
	_groupObjects?: ISlideObject[]
}
// PRIVATE ^^^

export interface WriteBaseProps {
	/**
	 * Whether to DEFLATE-compress the package (PowerPoint itself always compresses;
	 * set `false` only if export time matters more than file size)
	 * @default true
	 * @since v3.5.0 (default changed false→true in v4.0.0)
	 */
	compression?: boolean
	/**
	 * How to handle a media asset (image/audio/video) that fails to load during export.
	 * - `'throw'` (default): reject the export with an error naming the failing asset. A deck
	 *   that silently embeds a broken-image placeholder is a degenerate result, so failing
	 *   loudly is the safe default.
	 * - `'placeholder'`: substitute a broken-image placeholder, emit a `console.warn`, and
	 *   continue. Useful for best-effort/batch jobs where one missing asset should not abort
	 *   the whole deck.
	 * @default 'throw'
	 */
	onMediaError?: 'throw' | 'placeholder'
}
export interface WriteProps extends WriteBaseProps {
	/**
	 * Output type
	 * - values: 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array' | 'STREAM'
	 * @default 'blob'
	 */
	outputType?: WRITE_OUTPUT_TYPE
}
export interface WriteFileProps extends WriteBaseProps {
	/**
	 * Export file name
	 * @default 'Presentation.pptx'
	 */
	fileName?: string
}
export interface SectionProps {
	/**
	 * Section title
	 */
	title: string
	/**
	 * Section order - uses to add section at any index
	 * - values: 1-n
	 */
	order?: number
}
export interface SectionInternalProps extends SectionProps {
	_type?: 'user' | 'default'
	_slides: PresSlideInternal[]
}
export interface PresLayout {
	_sizeW?: number
	_sizeH?: number

	/**
	 * Layout Name
	 * @example 'LAYOUT_WIDE'
	 */
	name: string
	width: number
	height: number
}
export interface SlideNumberProps extends PositionProps, TextBaseProps {
	/**
	 * margin (points)
	 */
	margin?: Margin // TODO: convert to inches in 4.0 (valid values are 0-22)
}
export interface SlideMasterChartProps {
	type: CHART_NAME | IChartMulti[]
	data: OptsChartData[]
	options?: IChartOptsLib
	opts?: IChartOptsLib
}
export type SlideMasterObject =
	| { chart: SlideMasterChartProps }
	| { image: ImageProps }
	| { line: ShapeProps }
	| { rect: ShapeProps }
	| { roundRect: ShapeProps }
	/**
	 * Any preset shape, addressed by `SHAPE_NAME` (e.g. `pptx.ShapeType.ellipse`).
	 * Generalizes the `line`/`rect`/`roundRect` shortcuts to every preset the
	 * `addShape()` serializer supports (ellipse, triangle, chevron, …).
	 * @example { shape: { type: 'ellipse', options: { x: 1, y: 1, w: 2, h: 2, fill: { color: 'FF0000' } } } }
	 */
	| { shape: { type: SHAPE_NAME, options?: ShapeProps } }
	| { text: { text: string | number | TextProps[], options?: TextPropsOptions } }
	| {
		placeholder: {
			options: PlaceholderProps
			/**
			 * Text to be shown in placeholder (shown until user focuses textbox or adds text)
			 * - Leave blank to have powerpoint show default phrase (ex: "Click to add title")
			 */
			text?: string
		}
	}
/**
 * A child object that can be placed inside a group via `slide.addGroup()`.
 *
 * Uses the same key-tagged descriptor shape as `SlideMasterObject`, but limited to the
 * object types the flat-group MVP supports. Charts, media, tables, placeholders, and nested
 * groups are intentionally excluded (see `addGroup`); passing one logs a warning and skips it.
 * @since v4.0.0
 */
export type GroupChildProps =
	| { image: ImageProps }
	| { line: ShapeProps }
	| { rect: ShapeProps }
	| { roundRect: ShapeProps }
	| { shape: { type: SHAPE_NAME, options?: ShapeProps } }
	| { text: { text: string | number | TextProps[], options?: TextPropsOptions } }
/**
 * Options for `slide.addGroup()`.
 *
 * The group is a *flat* group (identity child coordinate space): children keep their
 * slide-absolute `x/y/w/h`. When `x/y/w/h` are omitted the group's bounds are auto-computed
 * as the bounding box of its children.
 * @since v4.0.0
 */
export interface GroupProps extends PositionProps, ObjectNameProps {
	/** Rotation in degrees (applied to the whole group) */
	rotate?: number
	/** Flip the group horizontally */
	flipH?: boolean
	/** Flip the group vertically */
	flipV?: boolean
}
export interface SlideMasterProps {
	/**
	 * Unique name for this master
	 */
	title: string
	background?: BackgroundProps
	margin?: Margin
	slideNumber?: SlideNumberProps
	objects?: SlideMasterObject[]

	/**
	 * @deprecated v3.3.0 - use `background`
	 */
	bkgd?: string | BackgroundProps
}
export interface ObjectOptions extends ImageBaseProps, PositionProps, ShapeProps, TableCellProps, TextPropsOptions {
	_placeholderIdx?: number
	_placeholderType?: PLACEHOLDER_TYPE
	/** Connector adjust-guide values (OOXML 1000ths-of-a-percent), one per bend; emitted as `<a:gd name="adjN">` */
	_connectorAdj?: number[]
	/** Connector start-point binding: target shape `objectName` + connection-site index; resolved to `<a:stCxn>` at serialize time */
	_startCxn?: { name: string, idx: number }
	/** Connector end-point binding: target shape `objectName` + connection-site index; resolved to `<a:endCxn>` at serialize time */
	_endCxn?: { name: string, idx: number }
	/**
	 * Image: which dimensions were omitted by the user and should be derived from the image's
	 * natural pixel size at serialize time. Path-based images can't be measured synchronously in
	 * `addImage()` (bytes are loaded async during export), so the missing extent is backfilled
	 * once `_relsMedia[].data` is populated. `{ w, h }` true means "derive this side from the
	 * natural ratio". Base64 `data` images are measured eagerly in `addImage()` and never set this.
	 */
	_szAuto?: { w: boolean, h: boolean }

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
	tableStyle?: TABLE_STYLE | string // table
}
export interface SlideBaseProps {
	_bkgdImgRid?: number
	_margin?: Margin
	_name?: string
	_presLayout: PresLayout
	_rels: ISlideRel[]
	_relsChart: ISlideRelChart[] // needed as we use args:"PresSlide|SlideLayout" often
	_relsMedia: ISlideRelMedia[] // needed as we use args:"PresSlide|SlideLayout" often
	_relsNotes?: ISlideRel[] // hyperlink rels emitted in the notes-slide part (notesSlideN.xml.rels)
	_slideNum: number
	_slideNumberProps?: SlideNumberProps | null
	_slideObjects: ISlideObject[]

	background?: BackgroundProps
	/**
	 * @deprecated v3.3.0 - use `background`
	 */
	bkgd?: string | BackgroundProps
}
export interface SlideLayout {
	background?: BackgroundProps
	/**
	 * @deprecated v3.3.0 - use `background`
	 */
	bkgd?: string | BackgroundProps
}
export interface SlideLayoutInternal extends SlideBaseProps, SlideLayout {
	_slide?: {
		_bkgdImgRid?: number
		back: string
		color: string
		hidden?: boolean
	} | null
}
export interface PresSlide {
	addChart(type: CHART_NAME, data: OptsChartData[], options?: IChartOpts): PresSlide
	addChart(type: IChartMulti[], options?: IChartOpts): PresSlide
	addConnector: (options: ConnectorProps) => PresSlide
	addImage: (options: ImageProps) => PresSlide
	addMedia: (options: MediaProps) => PresSlide
	addNotes: (notes: string | NotesProps | NotesProps[]) => PresSlide
	addShape: (shapeName: SHAPE_NAME, options?: ShapeProps) => PresSlide
	addTable: (tableRows: TableRow[], options?: TableProps) => PresSlide
	addText: (text: string | number | TextProps[], options?: TextPropsOptions) => PresSlide

	readonly newAutoPagedSlides?: PresSlide[]

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
	 * @example { path: '/home/gitbrent/images/myimg.png` } - retrieve image via local path
	 * @example { data: 'image/png;base64,iVtDaDrF[...]=' } - base64 string
	 * @since v3.3.0
	 */
	background?: BackgroundProps
	/**
	 * @deprecated v3.3.0 - use `background`
	 */
	bkgd?: string | BackgroundProps
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
export interface PresSlideInternal extends SlideBaseProps, PresSlide {
	_rId: number
	_slideLayout: SlideLayoutInternal | null
	_slideId: number
}
export interface AddSlideProps {
	masterName?: string // TODO: 20200528: rename to "masterTitle" (createMaster uses `title` so lets be consistent)
	sectionTitle?: string
}
export type CustomPropertyValue = string | number | boolean | Date

export interface PresentationProps {
	author: string
	company: string
	layout: string
	masterSlide: PresSlide
	/**
	 * Presentation's layout
	 * read-only
	 */
	presLayout: PresLayout
	revision: string
	/**
	 * Slide number to assign to the first slide (affects the slide-number field displayed in placeholders).
	 * @default 1
	 */
	firstSlideNum: number
	/**
	 * Whether to enable right-to-left mode
	 * @default false
	 */
	rtlMode: boolean
	subject: string
	theme: ThemeProps
	title: string
}
// PRIVATE interface
export interface IPresentationProps extends PresentationProps {
	masterSlide: PresSlideInternal
	sections: SectionInternalProps[]
	slideLayouts: SlideLayoutInternal[]
	slides: PresSlideInternal[]
}
