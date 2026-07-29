/**
 * The value types the shape read-model returns.
 *
 * Declarations only — no element access, no theme resolution. They live apart from the classes
 * that produce them so a consumer (or a sibling reader) can name a shape's effect, geometry or
 * recolour result without pulling in the whole shape hierarchy.
 */

import type { ResolvedColor } from '../theme-context.js'
// Type-only, so it is erased and no runtime import cycle exists: `ConnectionSite` names the
// shape union whose members are defined in terms of these very types.
import type { AnyShape } from '../shapes.js'

/** Discriminator for the concrete `Shape` subclass. */
export type ShapeType = 'autoShape' | 'picture' | 'connector' | 'graphicFrame' | 'group'

/** One end of a connector/line (`a:ln/a:headEnd` or `a:tailEnd`), as read from a shape. */
export interface LineEnd {
	/** Arrowhead type (`@type`: `none`/`triangle`/`stealth`/`diamond`/`oval`/`arrow`). */
	type: string
	/** Width class (`@w`: `sm`/`med`/`lg`), or `null` when unset. */
	width: string | null
	/** Length class (`@len`: `sm`/`med`/`lg`), or `null` when unset. */
	length: string | null
}

/** Both ends of a shape's line/border, when either carries an arrowhead. */
export interface LineEnds {
	head: LineEnd | null
	tail: LineEnd | null
}

/**
 * One bound end of a connector (`p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn` or `a:endCxn`),
 * as read. A connector authored with `startShape`/`endShape` attaches each end to
 * a shape by that shape's `p:cNvPr/@id` plus a connection-site index; an unbound
 * end (the writer emits a bare `p:cNvCxnSpPr`) reports no site at all — see
 * {@link Connector.startConnection}. This is the read counterpart of the write
 * API's `startShape`/`startShapeIdx` split, which binds by `objectName` and
 * resolves the name → id at serialize time.
 */
export interface ConnectionSite {
	/** The bound shape's drawing id (`@id`, i.e. its `p:cNvPr/@id`). */
	shapeId: number
	/** Connection-site index on the bound shape (`@idx`; 0-based, preset-dependent). */
	siteIndex: number
	/**
	 * The bound shape resolved to a read-model shape via {@link Slide.shapeByIdDeep},
	 * which descends into groups — so a connector bound to a shape nested in a group
	 * resolves the same as one bound to a top-level shape. `null` only when no shape
	 * anywhere on the slide carries that id (a genuinely dangling binding), which is
	 * faithful degradation and does not throw.
	 */
	boundShape: AnyShape | null
}

/**
 * A shape's outer drop shadow (`spPr/a:effectLst/a:outerShdw`), as read from a
 * shape and resolved against the slide theme. Distances are in points (the EMU
 * source ÷ 12700) and the direction in degrees (the `60000`ths source ÷ 60000),
 * matching the write-side {@link ShadowProps} convention so it round-trips.
 */
export interface OuterShadow {
	/** Effective shadow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the shadow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Shadow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Blur radius in points (`@blurRad` ÷ 12700), or `undefined` when unset. */
	blurPt?: number
	/** Offset distance in points (`@dist` ÷ 12700), or `undefined` when unset. */
	offsetPt?: number
	/** Offset direction in degrees, clockwise from 3 o'clock (`@dir` ÷ 60000), or `undefined` when unset. */
	angleDeg?: number
}

/**
 * A shape's inner shadow (`spPr/a:effectLst/a:innerShdw`), resolved against the
 * slide theme. Identical fields to {@link OuterShadow} — CT_InnerShadowEffect
 * carries the same `blurRad`/`dist`/`dir` + colour — but a distinct type so a
 * consumer can tell an inset shadow from a drop shadow. Distances in points,
 * direction in degrees, matching the write-side `shadow: { type: 'inner' }`.
 */
export interface InnerShadow {
	/** Effective shadow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the shadow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Shadow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Blur radius in points (`@blurRad` ÷ 12700), or `undefined` when unset. */
	blurPt?: number
	/** Offset distance in points (`@dist` ÷ 12700), or `undefined` when unset. */
	offsetPt?: number
	/** Offset direction in degrees, clockwise from 3 o'clock (`@dir` ÷ 60000), or `undefined` when unset. */
	angleDeg?: number
}

/**
 * A shape's glow effect (`spPr/a:effectLst/a:glow`) — a coloured halo — resolved
 * against the slide theme. The write-side text glow (`glow: { size, color,
 * opacity }`) emits the same element, so {@link radiusPt} (`@rad` ÷ 12700) and the
 * colour round-trip.
 */
export interface Glow {
	/** Effective glow colour as 6-hex (theme-resolved, transforms applied), or `null`. */
	color: string | null
	/** Theme colour token when the glow colour was a scheme colour (e.g. `accent1`), else `undefined`. */
	colorToken?: string
	/** Glow opacity 0–1 (from the colour's `a:alpha`), or `undefined` when fully opaque. */
	alpha?: number
	/** Glow radius in points (`@rad` ÷ 12700), or `undefined` when unset. */
	radiusPt?: number
}

/**
 * A shape's reflection effect (`spPr/a:effectLst/a:reflection`) — a mirrored fade
 * beneath the shape. This library's writer authors none, so this is a **read-only**
 * surface: a consumer that finds one should carry the part verbatim rather than
 * regenerate it. Only the attributes a faithful replica needs are decoded —
 * distances in points (÷ 12700), directions in degrees (÷ 60000), and the start/end
 * alpha and position pairs as 0–1 fractions (the `1000`ths-of-a-percent source
 * ÷ 100000). All fields are optional; an attribute absent from the source is omitted.
 */
export interface Reflection {
	/** Blur radius in points (`@blurRad` ÷ 12700). */
	blurPt?: number
	/** Start opacity 0–1 (`@stA` ÷ 100000). */
	startAlpha?: number
	/** Start position along the reflection 0–1 (`@stPos` ÷ 100000). */
	startPos?: number
	/** End opacity 0–1 (`@endA` ÷ 100000). */
	endAlpha?: number
	/** End position along the reflection 0–1 (`@endPos` ÷ 100000). */
	endPos?: number
	/** Offset distance in points (`@dist` ÷ 12700). */
	distPt?: number
	/** Offset direction in degrees (`@dir` ÷ 60000). */
	angleDeg?: number
	/** Fade direction in degrees (`@fadeDir` ÷ 60000). */
	fadeAngleDeg?: number
}

/**
 * A shape's soft-edge effect (`spPr/a:effectLst/a:softEdge`) — a feathered border.
 * Like {@link Reflection}, the writer authors none, so carry it verbatim. `@rad`
 * (the feather radius) ÷ 12700 → points.
 */
export interface SoftEdge {
	/** Feather radius in points (`@rad` ÷ 12700). */
	radiusPt: number
}

/**
 * A shape's pattern fill (`spPr/a:pattFill`) — a two-colour preset hatch. The
 * write-side `fill: { type: 'pattern', pattern: { preset, fgColor, bgColor } }`
 * emits the same element, so the {@link preset} name and both colours round-trip.
 * Colours resolve against the slide theme (a scheme token → literal hex) the same
 * way {@link Shape.resolvedFill} resolves a solid fill.
 */
export interface PatternFill {
	/** Preset pattern name (`@prst`, e.g. `pct50`/`diagCross`/`ltUpDiag`), or `null` when unset. */
	preset: string | null
	/** Foreground colour (`a:fgClr`) resolved against the theme, or `null`. */
	foreground: ResolvedColor | null
	/** Background colour (`a:bgClr`) resolved against the theme, or `null`. */
	background: ResolvedColor | null
}

/**
 * One segment of a custom-geometry path (`a:path`), as read from a shape. The
 * command verbs mirror the write-side `GeometryPoint` DSL (`src/core-interfaces.ts`)
 * so a consumer can map a `GeometryCommand[]` to `GeometryPoint[]` one-to-one.
 *
 * Coordinates (`x`/`y`/`x1`…) are raw path-unit integers in the path's own
 * `0..w` / `0..h` space (see {@link CustomGeometryPath.w}); they are not EMU and
 * must be scaled by the consumer against the shape's box. `arcTo` angles are in
 * **degrees** (the raw `60000`ths-of-a-degree values divided by 60000), matching
 * the write DSL's degree input.
 */
export type GeometryCommand =
	| { cmd: 'moveTo'; x: number; y: number }
	| { cmd: 'lnTo'; x: number; y: number }
	| { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { cmd: 'quadBezTo'; x1: number; y1: number; x: number; y: number }
	| { cmd: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number }
	| { cmd: 'close' }

/** One `<a:path>` of a custom geometry, with its path-unit viewport and render attrs. */
export interface CustomGeometryPath {
	/** Path-unit width (`a:path/@w`); the `x` denominator for this path's coords. Default `0`. */
	w: number
	/** Path-unit height (`a:path/@h`); the `y` denominator for this path's coords. Default `0`. */
	h: number
	/** Fill mode (`a:path/@fill`, `ST_PathFillMode`, e.g. `norm`/`none`/`lighten`). Default `norm`. */
	fill: string
	/** Whether the path is stroked (`a:path/@stroke`). Default `true`. */
	stroke: boolean
	/** The path's segments in document order — order *is* the geometry. */
	commands: GeometryCommand[]
}

/**
 * Custom freeform geometry (`spPr/a:custGeom/a:pathLst`), as read from a shape.
 * Faithfully exposes every `a:path` rather than flattening to a single command
 * list: the schema allows repeatable `a:path`, each with independent
 * `fill`/`stroke`. Desktop PowerPoint's own freeforms only ever emit one
 * `a:path` (a hole is one path with two `moveTo`…`close` contours), but
 * multi-`a:path` input is schema-legal (e.g. SVG import) and preserved here.
 */
export interface CustomGeometry {
	paths: CustomGeometryPath[]
}

/** A shape's resolved position and size in slide-absolute EMU. */
export interface AbsoluteFrame {
	left: number
	top: number
	width: number
	height: number
	/** Effective clockwise rotation in degrees after composing enclosing group rotations. */
	rotation: number
	/** Effective horizontal flip after XOR-composing enclosing group flips. */
	flipH: boolean
	/** Effective vertical flip after XOR-composing enclosing group flips. */
	flipV: boolean
}

/**
 * A colour reference inside a picture recolour effect, split by colour model
 * (mirrors {@link GradientStop}). At most one field is non-`null`.
 */
export interface RecolorColor {
	/** Explicit RGB as 6-hex (`a:srgbClr/@val`), or `null`. */
	color: string | null
	/** Theme colour token (`a:schemeClr/@val`, e.g. `accent1`), or `null`. */
	schemeColor: string | null
	/** Preset colour name (`a:prstClr/@val`, e.g. `black`/`white` — the duotone icon-tint stops), or `null`. */
	presetColor: string | null
}

/**
 * A picture's blip recolour effect (`p:blipFill/a:blip` recolour child), as read.
 * A small discriminated union over the effects a faithful reader needs to
 * reproduce a recoloured image. Colours use the same `color`/`schemeColor`/
 * `presetColor` split as {@link GradientStop}, so theme tokens can resolve through
 * {@link Slide.themeContext}. `threshold`/`amount` are 0–1 fractions.
 */
export type Recolor =
	| { kind: 'duotone'; stops: RecolorColor[] }
	| { kind: 'clrChange'; from: RecolorColor | null; to: RecolorColor | null }
	| { kind: 'grayscale' }
	| { kind: 'biLevel'; threshold: number | null }
	| { kind: 'alphaModFix'; amount: number }
