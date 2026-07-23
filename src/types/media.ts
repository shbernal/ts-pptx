/**
 * Image and audio/video types, including the shared `ImageBaseProps` base that `ObjectOptions` extends.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { SHAPE_NAME } from '../core-enums.js'
import type { Color, Coord, DataOrPathProps, DataOrPathRequiredProps, GeometryPoint, PositionProps } from './core.js'
import type { ObjectNameProps } from './object.js'
import type { ShapeAdjustValue } from './shape.js'
import type { HyperlinkProps, ShadowProps, ShapeLineProps } from './style.js'

export type MediaType = 'audio' | 'online' | 'video'

export interface ImageBaseProps extends PositionProps, ObjectNameProps {
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
	 *   position and size for any of `x`/`y`/`w`/`h` not supplied explicitly;
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
	 * Recolor the image to grayscale (`<a:grayscl/>`)
	 * - maps every pixel to its luminance grey (MS-PPT > Format Picture > Picture Color > Recolor > Grayscale)
	 * - the cheapest recolour: no payload
	 * - mutually exclusive with the other recolour modes (`duotone`/`biLevel`/`clrChange`); if several are set,
	 *   the first present in document order wins on read-back
	 * @default false
	 * @example true // desaturate the image
	 */
	grayscale?: boolean
	/**
	 * Recolor the image to two-level black & white (`<a:biLevel thresh="…"/>`)
	 * - every pixel at/above the luminance threshold becomes white, everything below it black
	 *   (MS-PPT > Format Picture > Picture Color > Recolor > Black and White …%)
	 * @example { threshold: 0.5 } // split at 50% luminance
	 */
	biLevel?: {
		/** Luminance split point as a `0.0–1.0` fraction (serialized to `thresh`, the 0–1 fraction ×100000). */
		threshold: number
	}
	/**
	 * Recolor the image by mapping one source color to another (`<a:clrChange>`)
	 * - every pixel matching `from` is repainted `to` — the classic "swap the flat background out" recolour
	 *   (MS-PPT > Format Picture > Picture Color > Set Transparent Color is the single-colour sibling)
	 * - colors accept `HexColor` or `ThemeColor`, same as fills / `duotone`
	 * @example { from: '000000', to: 'FF0000' } // turn black pixels red
	 */
	clrChange?: {
		/** Source color to replace (`<a:clrFrom>`). */
		from: Color
		/** Replacement color (`<a:clrTo>`). */
		to: Color
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
	 * @default "play button" image, gray background
	 */
	cover?: string
	/**
	 * media file extension
	 * - use when the media file path does not already have an extension, ex: "/folder/SomeSong"
	 * @default extension from file provided
	 */
	extn?: string
	/**
	 * Loop playback indefinitely (PowerPoint "Playback > Loop until Stopped")
	 * - emits a slide timing tree so the embedded audio/video repeats when played
	 * - not supported for `type: 'online'` (e.g. YouTube) embeds
	 * @default false
	 */
	loop?: boolean
	/**
	 * Total number of times to play the media (a finite loop), ex: `3` plays it three times
	 * - ignored when `loop` is `true` (which repeats forever)
	 * - not supported for `type: 'online'` (e.g. YouTube) embeds
	 */
	loopCount?: number
}
/**
 * Add an embedded OLE object (PowerPoint's Insert ▸ Object ▸ Create from File) to a slide.
 *
 * The payload's bytes travel inside the `.pptx` (in `ppt/embeddings/`), so double-clicking the
 * object in PowerPoint opens the source document in place. Requires either `data` or `path`.
 * Linked objects (`<p:link>` to a file outside the package) are not supported.
 */
interface OleObjectBaseProps extends PositionProps, ObjectNameProps {
	/**
	 * Picture of the embedded document shown on the slide — a raster image `path` (Node/local)
	 * or base64 `data:` URI.
	 * - the library is Node-first and cannot render an Office document, so when this is omitted a
	 *   neutral gray placeholder is embedded; PowerPoint draws the live object over it, but every
	 *   other consumer (and PowerPoint's own `mc:Fallback` path) shows exactly what is supplied here
	 * - supply a real screenshot whenever the deck is meant to read correctly outside PowerPoint
	 * @example { path: 'assets/budget-preview.png' }
	 */
	cover?: { path?: string; data?: string }
	/**
	 * Payload file extension, used to name the embedded part and to pick the content type,
	 * relationship type, and default `progId`.
	 * - inferred from `data`'s MIME, else `path`'s extension, else `progId`
	 * - anything that is not a known Office package extension (`xlsx`/`xlsm`/`docx`/`docm`/`pptx`/`pptm`)
	 *   is embedded as a generic OLE blob part named `.bin`
	 * @example 'xlsx'
	 */
	extn?: string
	/**
	 * OLE server ProgID — what PowerPoint launches on double-click.
	 * - defaults from the resolved extension (`xlsx` → `Excel.Sheet.12`, `docx` → `Word.Document.12`,
	 *   `pptx` → `PowerPoint.Show.12`, anything else → `Package`)
	 * @example 'Excel.Sheet.12'
	 */
	progId?: string
	/**
	 * Display the object as its application icon instead of a document preview (`showAsIcon`).
	 * - the `cover` image is still what gets drawn, so supply an icon-looking preview to match
	 * @default false
	 */
	showAsIcon?: boolean
	/**
	 * Native size of the preview image in EMU (`imgW`/`imgH`), which PowerPoint uses to keep the
	 * object's aspect ratio when it re-renders the embedded document.
	 * - defaults to the object's own `w`/`h` converted to EMU
	 */
	imgW?: number
	/** @see {@link imgW} */
	imgH?: number
}
/**
 * Options for `slide.addOleObject()`. Requires either `data` (base64, with or without a
 * `data:...;base64,` header) or `path` (a local/remote file read at export time).
 *
 * Sizing note: `w`/`h` default to 4 × 3 inches rather than being measured, since the library
 * does not open the embedded document.
 */
export type OleObjectProps = OleObjectBaseProps & DataOrPathRequiredProps
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
