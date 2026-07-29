/**
 * Core geometry and color types: `Coord`/`PositionProps`, the color model, gradient/pattern/image
 * fills, margins, alignment and freeform geometry points.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { ShapeFillProps } from './style.js'

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
	 * @example '/home/user/images/myimg.png` // retrieve image via local path
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
export type DataOrPathRequiredProps = (DataOrPathProps & { data: string }) | (DataOrPathProps & { path: string })
export interface BackgroundProps extends DataOrPathProps, ShapeFillProps {}
/**
 * Color in Hex format
 * @example 'FF3399'
 */
export type HexColor = string
export type ThemeColor =
	'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
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
	center?: { x: number; y: number }
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
	| 'pct5'
	| 'pct10'
	| 'pct20'
	| 'pct25'
	| 'pct30'
	| 'pct40'
	| 'pct50'
	| 'pct60'
	| 'pct70'
	| 'pct75'
	| 'pct80'
	| 'pct90'
	| 'horz'
	| 'vert'
	| 'ltHorz'
	| 'ltVert'
	| 'dkHorz'
	| 'dkVert'
	| 'narHorz'
	| 'narVert'
	| 'dashHorz'
	| 'dashVert'
	| 'cross'
	| 'dnDiag'
	| 'upDiag'
	| 'ltDnDiag'
	| 'ltUpDiag'
	| 'dkDnDiag'
	| 'dkUpDiag'
	| 'wdDnDiag'
	| 'wdUpDiag'
	| 'dashDnDiag'
	| 'dashUpDiag'
	| 'diagCross'
	| 'smCheck'
	| 'lgCheck'
	| 'smGrid'
	| 'lgGrid'
	| 'dotGrid'
	| 'smConfetti'
	| 'lgConfetti'
	| 'horzBrick'
	| 'diagBrick'
	| 'solidDmnd'
	| 'openDmnd'
	| 'dotDmnd'
	| 'plaid'
	| 'sphere'
	| 'weave'
	| 'divot'
	| 'shingle'
	| 'wave'
	| 'trellis'
	| 'zigZag'

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
 * - used by shapes (`ShapeType.custGeom`) and by images (clips the picture to the path)
 *
 * The `arc` node carries no `x`/`y`: an `<a:arcTo>` has no explicit end point. The renderer
 * derives it from the current pen position, the radii and the swept angle, so an end point
 * authored here would be discarded (a warning is emitted if one is supplied).
 */
export type GeometryPoint =
	| { x: Coord; y: Coord; moveTo?: boolean }
	| { curve: { type: 'arc'; hR: Coord; wR: Coord; stAng: number; swAng: number } }
	| { x: Coord; y: Coord; curve: { type: 'cubic'; x1: Coord; y1: Coord; x2: Coord; y2: Coord } }
	| { x: Coord; y: Coord; curve: { type: 'quadratic'; x1: Coord; y1: Coord } }
	| { close: true }
