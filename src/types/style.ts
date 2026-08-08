/**
 * Visual styling types shared across object kinds: borders, shadows, shape fill and line
 * (stroke) options, connectors and hyperlinks.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { LineCap } from './chart.js'
import type { Color, Coord, GradientFillProps, HexColor, ImageFillProps, PatternFillProps } from './core.js'
import type { PresetLineDashVal } from '../ooxml/st-enums.js'

// used by charts, shape, text
export interface BorderProps {
	/**
	 * Border type
	 * @default solid
	 */
	type?: 'none' | 'dash' | 'solid'
	/**
	 * Dash pattern, using the same vocabulary as {@link ShapeLineProps.dashType}.
	 *
	 * `type` is only a coarse three-way switch, so every dashed border it can express
	 * collapses onto `sysDash`. Set this to pick the exact `a:prstDash` preset instead:
	 * `dashType: 'lgDashDot'` emits `<a:prstDash val="lgDashDot"/>`.
	 *
	 * **Precedence.** When both are set, `dashType` wins over `type` for the dash pattern —
	 * `{ type: 'solid', dashType: 'sysDot' }` draws a dotted rule. The one exception is
	 * `type: 'none'`, which suppresses the border entirely and is decided before any dash
	 * pattern is chosen. An unrecognized value warns and falls back to what `type` implies.
	 * @default (derived from `type`: `sysDash` for `'dash'`, else `solid`)
	 * @example { type: 'solid', color: '999999', dashType: 'lgDash' }
	 */
	dashType?: ShapeLineProps['dashType']
	/**
	 * Border color (hex)
	 * @example 'FF3399'
	 * @default '666666'
	 */
	color?: HexColor

	/**
	 * Border width (points)
	 * - MS-PPT > Format Shape > Fill & Line > Line > Width
	 * @default 1
	 */
	width?: number
	/**
	 * Border transparency (percent)
	 * - MS-PPT > Format Shape > Fill & Line > Line > Transparency
	 * - range: 0-100
	 * @default 0
	 */
	transparency?: number
	/**
	 * Line end cap style
	 * @default 'flat'
	 */
	cap?: LineCap
}
/**
 * Slide-show navigation action for action-button shapes. Each value maps 1:1 onto
 * `ppaction://hlinkshowjump?jump=<value>`. A specific numbered slide is reached via
 * `HyperlinkProps.slide` instead (which emits `ppaction://hlinksldjump`).
 */
export type HyperlinkActionJump = 'firstslide' | 'previousslide' | 'nextslide' | 'lastslide' | 'lastslideviewed' | 'endshow'
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
	/**
	 * Slide-show navigation action for action-button shapes. Emits a relationship-less
	 * `<a:hlinkClick action="ppaction://hlinkshowjump?jump=…"/>` on the shape.
	 * @example hyperlink: { action: 'nextslide' }
	 */
	action?: HyperlinkActionJump
}
// used by: chart, text, image
export interface ShadowProps {
	/**
	 * shadow type
	 * @default 'none'
	 */
	type: 'outer' | 'inner' | 'none'
	/**
	 * shadow transparency (percent)
	 * - MS-PPT > Format Shape > Effects > Shadow > Transparency
	 * - range: 0-100 (0 = fully opaque, 100 = fully transparent)
	 * @example 25 // 25% transparent
	 */
	transparency?: number
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
	offset?: number // PowerPoint UI: "Distance"
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
	 * @example SchemeColor.text1 // Theme color (Text1)
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
	 * Dash type — the full `ST_PresetLineDashVal` set, so any dash a source deck can
	 * carry can also be authored and replicated.
	 * @default 'solid'
	 */
	dashType?: PresetLineDashVal
	/**
	 * Line end cap style
	 * @default 'flat'
	 */
	cap?: LineCap
	/**
	 * Begin arrow type
	 */
	beginArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	/**
	 * End arrow type
	 */
	endArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	// FUTURE: beginArrowSize (1-9)
	// FUTURE: endArrowSize (1-9)
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
	 *
	 * A shape inside a group is a valid target — group children are named on the same slide — though
	 * the connector itself cannot be a group child (see {@link Slide.addGroup}).
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
