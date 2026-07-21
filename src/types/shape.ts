/**
 * Preset/freeform shape types and their adjust values.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { Coord, GeometryPoint, HAlign, PositionProps } from './core.js'
import type { ObjectNameProps } from './object.js'
import type { HyperlinkProps, ShadowProps, ShapeFillProps, ShapeLineProps } from './style.js'

/**
 * A single preset-geometry adjustment guide (`<a:gd>` inside `<a:avLst>`).
 * - `name` is the guide name the preset defines, e.g. `'adj'`, `'adj1'`, `'adj2'`.
 *   PowerPoint shows these handles as the yellow drag dots on a selected shape.
 * - `value` is a fraction `0.0–1.0` of the handle's range, emitted as a percentage
 *   guide formula (`val`, in 1/100000 units, so `0.25` → `fmla="val 25000"`).
 *   Most adjustment handles (corner radius, chevron point, callout depth, bevel
 *   width, …) are percentage-based and map directly; some shapes accept values
 *   beyond `1.0`. For angle-based handles, prefer the `angleRange` shortcut.
 */
export interface ShapeAdjustValue {
	name: string
	value: number
}

/**
 * A geometry guide (`<a:gd>` inside a `custGeom`'s `<a:gdLst>`): a named construction formula
 * that connection sites, adjust handles, and the text rect can reference by name.
 * - `name` is a `ST_GeomGuideName` token, e.g. `'w2'`.
 * - `formula` is a `ST_GeomGuideFormula` — a space-delimited op + args string emitted verbatim,
 *   e.g. the multiply-divide op `"* / w 1 2"` written without the inner space (w × 1 ÷ 2), or
 *   `'val 10800'`. This is an advanced escape hatch: the string is not interpreted, only escaped.
 * @see http://www.datypic.com/sc/ooxml/e-a_gd-1.html
 */
export interface ShapeGuide {
	name: string
	formula: string
}

/**
 * A connection site (`<a:cxn>` inside a `custGeom`'s `<a:cxnLst>`): a point a connector can attach
 * to. A connector's `startShapeIdx`/`endShapeIdx` index into the target shape's connection sites.
 */
export interface ShapeConnectionSite {
	/** Incoming-connector angle in degrees (emitted to `ang`, `ST_AdjAngle` 60000ths). */
	ang: number
	/** X in the object's own space (like `points`), or a guide-name string emitted verbatim. */
	x: Coord | string
	/** Y in the object's own space (like `points`), or a guide-name string emitted verbatim. */
	y: Coord | string
}

/** An XY adjust handle (`<a:ahXY>`) — a draggable handle, optionally bound to guides. */
export interface ShapeAdjustHandleXY {
	x: Coord | string
	y: Coord | string
	gdRefX?: string
	minX?: Coord | string
	maxX?: Coord | string
	gdRefY?: string
	minY?: Coord | string
	maxY?: Coord | string
}

/** A polar adjust handle (`<a:ahPolar>`) — radius/angle rather than XY. */
export interface ShapeAdjustHandlePolar {
	x: Coord | string
	y: Coord | string
	gdRefR?: string
	minR?: Coord | string
	maxR?: Coord | string
	gdRefAng?: string
	minAng?: number
	maxAng?: number
}

export type ShapeAdjustHandle = ShapeAdjustHandleXY | ShapeAdjustHandlePolar

export interface ShapeProps extends PositionProps, ObjectNameProps {
	/**
	 * Horizontal alignment
	 * @default 'left'
	 */
	align?: HAlign
	/**
	 * Radius (only for pptx.ShapeType.pie, pptx.ShapeType.arc, pptx.ShapeType.blockArc)
	 * - In the case of pptx.ShapeType.blockArc you have to setup the arcThicknessRatio
	 * - values: [0-359, 0-359]
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
	 * @example { name: 'adj', value: 0.25 } // set the single adjust handle to 25%
	 * @example [{ name: 'adj1', value: 0.5 }, { name: 'adj2', value: 0.25 }] // two handles
	 */
	shapeAdjust?: ShapeAdjustValue | ShapeAdjustValue[]
	/**
	 * Radius (only for pptx.ShapeType.blockArc)
	 * - You have to setup the angleRange values too
	 * - values: 0.0-1.0
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
	 * @example hyperlink: { url: "https://example.com", tooltip: "Visit Homepage" },
	 */
	hyperlink?: HyperlinkProps
	/**
	 * Line options
	 */
	line?: ShapeLineProps
	/**
	 * Points (only for pptx.ShapeType.custGeom)
	 * - type: 'arc' (no end point — it is computed from the pen position, radii and sweep)
	 * - `hR` Shape Arc Height Radius
	 * - `wR` Shape Arc Width Radius
	 * - `stAng` Shape Arc Start Angle (degrees, not wrapped into 0..360)
	 * - `swAng` Shape Arc Swing Angle (degrees, not wrapped into 0..360)
	 * @see http://www.datypic.com/sc/ooxml/e-a_arcTo-1.html
	 * @example [{ x: 0, y: 0 }, { x: 10, y: 10 }] // draw a line between those two points
	 */
	points?: GeometryPoint[]
	/**
	 * Geometry guides (`<a:gdLst>`) for a `custGeom` shape — named construction formulas that
	 * connection sites / adjust handles can reference. Advanced: each `formula` is emitted verbatim.
	 * @example [{ name: 'w2', formula: 'val 10800' }]
	 */
	guides?: ShapeGuide[]
	/**
	 * Connection sites (`<a:cxnLst>`) for a `custGeom` shape — the points a connector can attach to.
	 * Positions are absolute in the object's own space (like `points`); a `startShapeIdx`/`endShapeIdx`
	 * on an `addConnector` indexes into this list.
	 * @example [{ ang: 0, x: 1, y: 0 }, { ang: 180, x: 0, y: 1 }]
	 */
	connectionSites?: ShapeConnectionSite[]
	/**
	 * Adjust handles (`<a:ahLst>`) for a `custGeom` shape — draggable handles (XY or polar),
	 * optionally bound to guides via `gdRef*`.
	 */
	adjustHandles?: ShapeAdjustHandle[]
	/**
	 * Rounded rectangle radius (only for pptx.ShapeType.roundRect)
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
	 */
	shadow?: ShadowProps
}
