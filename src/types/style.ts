/**
 * Visual styling types shared across object kinds: borders, shadows, shape fill and line
 * (stroke) options, connectors and hyperlinks.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { LineCap } from './chart.js'
import type { Color, Coord, GradientFillProps, HexColor, ImageFillProps, PatternFillProps } from './core.js'
import type { LineEndType, PresetLineDashVal } from '../ooxml/st-enums.js'

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
export type HyperlinkActionJump =
	| 'firstslide'
	| 'previousslide'
	| 'nextslide'
	| 'lastslide'
	| 'lastslideviewed'
	| 'endshow'
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
	 * - MS-PPT > Format Shape > Effects > Shadow > Blur
	 * - PowerPoint's own spinner stops at 100, but a larger blur is unusual rather than
	 *   invalid: it loads and paints. Only a *negative* blur is rejected (clamped to `0` with
	 *   a `shadow/blur-out-of-range` diagnostic), because the `blurRad` attribute is unsigned.
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
	 * - MS-PPT > Format Shape > Effects > Shadow > Distance
	 * - PowerPoint's own spinner stops at 200, but a larger distance is unusual rather than
	 *   invalid: it loads and paints. Only a *negative* offset is rejected (clamped to `0` with
	 *   a `shadow/offset-out-of-range` diagnostic), because the `dist` attribute is unsigned.
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
/**
 * What a fill option accepts: a {@link ShapeFillProps} object, or a bare {@link Color} as
 * shorthand for a solid fill in that colour.
 *
 * `'FF0000'` and `{ color: 'FF0000' }` are the same statement and emit the same
 * `<a:solidFill>` — the shorthand is lossless, not a lesser spelling, and every other key on
 * the object (`transparency`, `gradient`, `pattern`, `image`) is simply one the shorthand has
 * no way to say. Reach for the object as soon as you need any of them.
 *
 * The shorthand is deliberately **not** offered on {@link ShapeLineProps}. A stroke carries
 * width and dash alongside its paint, and those defaults are applied by rebuilding the line
 * object at definition time; a bare string arrives with nothing to rebuild, so it would paint
 * the colour and silently drop `w` and `prstDash`. A stroke says more than a colour, so it is
 * spelled as an object.
 * @example fill: 'FF0000' // solid red
 * @example fill: { color: 'FF0000', transparency: 50 } // the same red, half transparent
 */
export type FillOption = Color | ShapeFillProps

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
	 *
	 * `'none'` and `'inherit'` are different states, and both have to be spelled out
	 * because omitting `fill` is not one of them: the shape/text-box emitter defaults a
	 * missing `fill` to `<a:noFill/>`.
	 * - `'none'` emits `<a:noFill/>` — an explicit transparent interior.
	 * - `'inherit'` emits no fill child at all, leaving the interior to
	 *   `p:style/a:fillRef` or the placeholder. This is the only way to author that
	 *   state from `addShape`/`addText`.
	 *
	 * On a table cell (`TableCellProps.fill`) and a slide background the two spellings
	 * differ the same way, except that *omitting* the option there already means
	 * inherit, so `'inherit'` is merely the explicit form of the default.
	 *
	 * **Omitting `type` is not the same as omitting the fill.** Setting `gradient`,
	 * `pattern` or `image` selects that kind on its own — the sub-object is the statement,
	 * and `type` is how you say something the sub-objects cannot. Where both are given and
	 * they disagree, **the explicit `type` wins** and the sub-object is ignored, because
	 * that is the only rule under which `{ type: 'none', gradient }` can still mean
	 * transparent.
	 * @default 'solid'
	 */
	type?: 'none' | 'inherit' | 'solid' | 'gradient' | 'pattern' | 'image'

	/**
	 * Native PPTX gradient fill options. Setting this (or `type: 'gradient'`) paints the
	 * shape interior with a gradient instead of a color.
	 */
	gradient?: GradientFillProps

	/**
	 * Native PPTX pattern fill options. Setting this (or `type: 'pattern'`) paints the
	 * shape interior with a two-color preset pattern instead of a color.
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
 * A stroke is painted much like a fill: in addition to a solid `color`, it accepts the
 * inherited `gradient` and `pattern` fill options. Setting `gradient` (or
 * `type: 'gradient'`) paints a gradient stroke, e.g. `line: { width: 1, gradient: {
 * kind: 'linear', angle: 0, stops: [{ position: 0, color: 'accent3' }, { position: 100,
 * color: 'accent4' }] } }`.
 *
 * **A picture stroke is not one of them**, which is why `image` and `type: 'image'` are
 * subtracted from the inherited fill props rather than carried. `<a:ln>`'s paint child is
 * `EG_LineFillProperties` — `a:noFill`, `a:solidFill`, `a:gradFill`, `a:pattFill` and
 * nothing else — so unlike a shape interior a stroke has no `a:blipFill` slot to put a
 * bitmap in, and a package that emitted one is what PowerPoint answers with a repair
 * prompt. The interface carried `image` from `ShapeFillProps` regardless, and no call site
 * ever registered the media for it, so a picture stroke reached the emitter with no rel and
 * painted nothing. It is now spelled out as unsupported: a JS caller who sets it anyway
 * gets `line/image-fill-unsupported`.
 *
 * A stroke also does not take the bare-colour shorthand a fill does — see {@link FillOption}
 * for why. Spell it `line: { color: 'FF0000' }`.
 */
export interface ShapeLineProps extends Omit<ShapeFillProps, 'type' | 'image' | '_imgRid'> {
	/**
	 * Stroke paint kind. The same vocabulary as {@link ShapeFillProps.type} minus
	 * `'image'`, which `<a:ln>` cannot express (see the note on this interface).
	 * @default 'solid'
	 */
	type?: 'none' | 'inherit' | 'solid' | 'gradient' | 'pattern'
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
	 * Begin arrow type — the `ST_LineEndType` set (`a:headEnd/@type`).
	 */
	beginArrowType?: LineEndType
	/**
	 * End arrow type — the `ST_LineEndType` set (`a:tailEnd/@type`).
	 */
	endArrowType?: LineEndType
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
