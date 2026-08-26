/**
 * A cheap, flat snapshot of what a `.pptx` contains: every element on every
 * slide, with its slide-absolute box and the handful of text/fill properties a
 * layout audit asks about.
 *
 * It is a **projection over `ts-pptx/read`**, not a second reader. `read` gives a
 * navigable, mutable model of a package and answers questions in the shape of the
 * OOXML tree; this answers one flat question — "what is on the slides, and where"
 * — in the shape a linter, an overlap check, or a diffing tool wants. The two used
 * to be independent implementations over two different XML parsers, which meant
 * every read-side fix had to be made twice and a divergence between them was
 * invisible to every test. Now the model below is the only reader, and this file
 * is the flattening.
 *
 * What it deliberately does NOT do, and why the projection is not just
 * `Presentation` re-exported:
 * - **Only geometry a shape carries itself.** A placeholder that inherits its box
 *   from the layout is omitted rather than resolved (`absoluteFrame`, not
 *   `resolvedFrame`) — the snapshot reports what the slide states.
 * - **A `p:graphicFrame` is a box, not its contents.** A table, chart or SmartArt
 *   graphic is reported as one element with its box, its kind, and the text a
 *   reader sees on the slide, but its structure is not flattened: no per-run
 *   formatting, no cells, no series. Walk `ts-pptx/read` for what is inside one.
 */
import { warn } from './diagnostics.js'
import { OpcPackage } from './read/opc/package.js'
import { Presentation } from './read/api/presentation.js'
import type { AnyShape } from './read/api/shapes.js'
import type { GraphicFrame } from './read/api/shapes/graphic-frame.js'
import type { Run, TextFrame } from './read/api/text.js'
import { ELEMENT_NODE, OOXML_NS, attr, firstChild, intValue, type Element } from './read/oxml/dom.js'
import { STANDARD_LAYOUTS, emuToInches } from './units.js'

/**
 * Input to the inspect surface. A `string` is a **filesystem path** (Node) read
 * from disk; pass `Uint8Array`/`ArrayBuffer`/`Blob`/`number[]` for an in-memory
 * archive. (A string is a path, not JSZip's latin1 binary content — see
 * {@link ZipInput}.)
 */
type PptxInspectInputValue = string | number[] | Uint8Array | ArrayBuffer | Blob

export type PptxInspectInput = PptxInspectInputValue | Promise<PptxInspectInputValue>

export interface PptxSlideSize {
	widthIn: number
	heightIn: number
}

export interface PptxBox {
	x: number
	y: number
	w: number
	h: number
}

export interface PptxTextRun {
	text: string
	fontSizePt: number | null
	color: string | null
	/** Run font face (`a:rPr > a:latin@typeface`), or null when it inherits from the theme/placeholder. */
	fontFace: string | null
	/** Bold (`a:rPr@b`). */
	bold: boolean
	/** Italic (`a:rPr@i`). */
	italic: boolean
	/** Strikethrough token (`a:rPr@strike`: `noStrike`/`sngStrike`/`dblStrike`), or null when unset. */
	strike: string | null
	/** Highlight colour hex (`a:rPr > a:highlight`), or null when unset. A theme token is resolved against the slide's theme. */
	highlight: string | null
	/** Character spacing in points (`a:rPr@spc`, authored in hundredths of a point), or null when unset. */
	charSpacingPt: number | null
}

/**
 * One paragraph (`a:p`): its runs in document order. Preserves the paragraph/line
 * boundaries that the flat {@link PptxSlideElement.textRuns} list discards — needed
 * to measure a `wrap="none"` frame line-by-line (each paragraph is one unwrapped line).
 */
export interface PptxParagraph {
	runs: PptxTextRun[]
}

/**
 * What an element is. `'group'` is a `p:grpSp` container: it has its own identity,
 * box, and fill, but no text of its own — its content is the elements that name it
 * in {@link PptxSlideElement.parentZIndex}. `'graphicFrame'` is a `p:graphicFrame`
 * hosting a table, a chart or a SmartArt graphic; {@link PptxSlideElement.graphicKind}
 * says which, and it keeps that kind even when it carries text (a `'text'` element is
 * always a `p:sp`).
 */
export type PptxSlideElementKind = 'text' | 'image' | 'shape' | 'group' | 'graphicFrame'

/**
 * What a `'graphicFrame'` element hosts, read from its `a:graphicData/@uri`:
 * a table (`a:tbl`), a classic chart (`c:chart`), a 2016-family chart
 * (`cx:chartSpace`: waterfall/funnel/treemap/…), or a SmartArt diagram. `'other'`
 * is a frame whose URI matches none of those — an OLE object, an ink part, a 3D
 * model — reported as a box so it is not silently missing from the slide.
 */
export type PptxGraphicKind = 'table' | 'chart' | 'chartEx' | 'diagram' | 'other'

/**
 * Vertical-autofit mode of a text frame, read from the `a:bodyPr` child element:
 * - `'none'`: no autofit (`a:noAutofit`, or no autofit child at all). The box has a
 *   fixed height the text must fit inside — a genuine overflow candidate.
 * - `'normAutofit'`: shrink text to fit (`a:normAutofit`, ts-pptx `fit: 'shrink'`).
 *   Text is downscaled rather than overflowing.
 * - `'spAutoFit'`: resize shape to fit text (`a:spAutoFit`, ts-pptx `fit: 'resize'`).
 *   The authored height is an output, not a constraint, so the box cannot overflow.
 */
export type PptxAutofitMode = 'none' | 'normAutofit' | 'spAutoFit'

/**
 * Text-frame body insets in inches (`a:bodyPr` `lIns`/`tIns`/`rIns`/`bIns`), with
 * PowerPoint defaults applied when an attribute is absent (0.1in left/right,
 * 0.05in top/bottom). Subtract these from {@link PptxBox} to get the inner text box.
 */
export interface PptxBodyInsets {
	left: number
	top: number
	right: number
	bottom: number
}

export interface PptxSlideElement {
	id: string | number
	name: string
	kind: PptxSlideElementKind
	/**
	 * What a `'graphicFrame'` element hosts, or `null` for every other kind. See
	 * {@link PptxGraphicKind}.
	 */
	graphicKind: PptxGraphicKind | null
	/**
	 * Paint order on the slide, `0`-based: the element's position in a depth-first
	 * walk of `p:spTree` in document order. A group is immediately followed by its
	 * own children, which is also the order PowerPoint paints them, so a higher
	 * `zIndex` draws on top. Unique within a slide, which is what
	 * {@link parentZIndex} / {@link childZIndices} key on.
	 */
	zIndex: number
	/**
	 * Position and size in **slide-absolute** inches, composing every enclosing
	 * group transform — directly comparable across elements whether or not they are
	 * grouped. For a rotated element this is the unrotated placement box (the box
	 * PowerPoint writes after Ungroup); see {@link rotation}.
	 */
	box: PptxBox
	/**
	 * Effective clockwise rotation in degrees, normalised to `[0, 360)`, after
	 * composing the element's own `a:xfrm@rot` with every enclosing group rotation.
	 */
	rotation: number
	/** Effective horizontal flip, XOR-composing the element's own `@flipH` with enclosing group flips. */
	flipH: boolean
	/** Effective vertical flip, XOR-composing the element's own `@flipV` with enclosing group flips. */
	flipV: boolean
	/** {@link zIndex} of the enclosing `'group'` element, or `null` when the element sits at slide level. */
	parentZIndex: number | null
	/** {@link zIndex} of each direct child, in document order. Empty unless {@link kind} is `'group'`. */
	childZIndices: number[]
	/**
	 * The element's text, whitespace-collapsed to one line. For a `'graphicFrame'`
	 * this is the text the structure puts on the slide (table cells in row order,
	 * SmartArt node text), which is what {@link PptxSlideInspection.text} and
	 * `wordCount` then count; a chart contributes nothing, matching `Slide.text` on
	 * the read model. That text has no {@link textRuns} or {@link paragraphs} here —
	 * per-run formatting inside a structure is what this surface does not flatten.
	 */
	text: string
	textWrap: string | null
	autofit: PptxAutofitMode | null
	/**
	 * Baked shrink scale from `<a:normAutofit@fontScale>` as a percent (62.5 = 62.5%),
	 * or null when the frame has no `normAutofit` or bakes no scale (a bare
	 * `<a:normAutofit/>`, which PowerPoint draws at 100% until edited).
	 */
	autofitFontScale: number | null
	bodyInsets: PptxBodyInsets | null
	textRuns: PptxTextRun[]
	/** Runs grouped by their source paragraph (`a:p`), in document order. */
	paragraphs: PptxParagraph[]
	fontSizes: number[]
	colors: string[]
	fill: string | null
	line: string | null
	shapeType: string | null
}

export interface PptxSlideInspection {
	index: number
	name: string
	path: string
	size: PptxSlideSize
	elements: PptxSlideElement[]
	text: string
	wordCount: number
}

export interface PptxInspection {
	slideSize: PptxSlideSize
	slides: PptxSlideInspection[]
}

export type PptxBoxAxis = 'x' | 'y'
export type PptxBoxAnchor = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'

export const DEFAULT_INSPECT_SLIDE_SIZE: PptxSlideSize = Object.freeze({
	widthIn: round(STANDARD_LAYOUTS.LAYOUT_WIDE.widthIn, 3),
	heightIn: STANDARD_LAYOUTS.LAYOUT_WIDE.heightIn,
})

/* ────────────────────────────────────────────────────────────────────────────
 * Package layer
 *
 * Thin conveniences over `OpcPackage`, kept because this surface speaks zip
 * paths (`ppt/slides/slide1.xml`) while OPC speaks partnames (`/ppt/...`), and
 * because reaching a part's bytes is the one thing a caller of a *flat* surface
 * still routinely needs.
 * ──────────────────────────────────────────────────────────────────────────── */

/**
 * Load a `.pptx` as an {@link OpcPackage} — the same package model `ts-pptx/read`
 * uses, so a caller that starts here can hand the result straight to
 * `Presentation.fromPackage()` without re-reading the bytes.
 *
 * The input must be a real OPC package: a zip that merely contains some slide XML
 * but no `[Content_Types].xml` is rejected with a `PackageReadError`.
 */
export async function loadPptxPackage(input: PptxInspectInput): Promise<OpcPackage> {
	return OpcPackage.load(await input)
}

/** Every part in the package, as zip paths (`ppt/slides/slide1.xml`), sorted. */
export function listPptxParts(pptxPackage: OpcPackage): string[] {
	return [...pptxPackage.parts.keys()].map(zipPathOf).sort()
}

/** Read a package part as UTF-8 text, or `null` when the part is absent. */
export async function readPptxTextPart(pptxPackage: OpcPackage, path: string): Promise<string | null> {
	const bytes = await readPptxBinaryPart(pptxPackage, path)
	return bytes ? textDecoder.decode(bytes) : null
}

/**
 * Read a package part as raw bytes — the binary sibling of {@link readPptxTextPart}
 * for embedded media (SVG/PNG/EMF blobs, fonts, …) that must not be UTF-8 decoded.
 * Returns `null` when the part is absent. The `Uint8Array` is browser-isomorphic;
 * Node consumers can wrap it with `Buffer.from(...)` if they need Buffer methods.
 */
export async function readPptxBinaryPart(pptxPackage: OpcPackage, path: string): Promise<Uint8Array | null> {
	return pptxPackage.part(partNameOf(path))?.bytes ?? null
}

const textDecoder = new TextDecoder('utf-8')

/** OPC partnames are absolute (`/ppt/…`); this surface speaks zip paths (`ppt/…`). */
function zipPathOf(partName: string): string {
	return partName.startsWith('/') ? partName.slice(1) : partName
}

function partNameOf(zipPath: string): string {
	return zipPath.startsWith('/') ? zipPath : `/${zipPath}`
}

/* ────────────────────────────────────────────────────────────────────────────
 * The surface
 * ──────────────────────────────────────────────────────────────────────────── */

export async function inspectPptx(input: PptxInspectInput): Promise<PptxInspection> {
	const pptxPackage = await loadPptxPackage(input)
	const slideSize = await readPresentationSize(pptxPackage)
	const slides = await extractSlides(pptxPackage, slideSize)
	return { slideSize, slides }
}

export async function readPresentationSize(
	pptxPackage: OpcPackage,
	fallback: PptxSlideSize = DEFAULT_INSPECT_SLIDE_SIZE
): Promise<PptxSlideSize> {
	const size = presentationOf(pptxPackage)?.slideSize
	if (!size) return fallback
	return { widthIn: round(size.widthIn, 3), heightIn: round(size.heightIn, 3) }
}

/**
 * Every slide in the deck, in **presentation order** (`p:sldIdLst`) — the order
 * PowerPoint shows them in, which is not the order their parts are named once a
 * deck has been reordered.
 */
export async function extractSlides(pptxPackage: OpcPackage, size?: PptxSlideSize): Promise<PptxSlideInspection[]> {
	const slideSize = size || (await readPresentationSize(pptxPackage))
	const presentation = presentationOf(pptxPackage)
	if (!presentation) return []

	return presentation.slides.map((slide, index) => {
		const path = zipPathOf(slide.partName)
		const harvested: HarvestedShape[] = []
		harvest(slide.shapes, null, harvested) // slide level: no enclosing group
		const elements = flatten(harvested, path)
		const text = elements
			.map((el) => el.text)
			.filter(Boolean)
			.join(' ')

		return {
			index,
			name: slide.name || `Slide ${index + 1}`,
			path,
			size: slideSize,
			elements,
			text,
			wordCount: countWords(text),
		}
	})
}

/**
 * The deck's `Presentation`, or `null` when the package holds no readable
 * `presentation.xml`. Callers here are auditing files they did not author, so a
 * package that is a valid OPC container but not a presentation reports an empty
 * deck rather than throwing. `presentationPart` is the throw site: it rejects a
 * package with no — or more than one — office-document relationship.
 */
function presentationOf(pptxPackage: OpcPackage): Presentation | null {
	const presentation = Presentation.fromPackage(pptxPackage)
	try {
		return presentation.presentationPart.isXmlPart ? presentation : null
	} catch {
		return null
	}
}

export function overlapArea(a: PptxBox, b: PptxBox): number {
	const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x))
	const y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y))
	return x * y
}

export function boxAnchor(box: PptxBox, anchor: PptxBoxAnchor, axis: PptxBoxAxis): number {
	if (axis === 'x') {
		if (anchor === 'left') return box.x
		if (anchor === 'right') return box.x + box.w
		return box.x + box.w / 2
	}
	if (anchor === 'top') return box.y
	if (anchor === 'bottom') return box.y + box.h
	return box.y + box.h / 2
}

/* ────────────────────────────────────────────────────────────────────────────
 * Shape-tree harvest
 * ──────────────────────────────────────────────────────────────────────────── */

/** One harvested shape, before it is flattened into a {@link PptxSlideElement}. */
interface HarvestedShape {
	shape: AnyShape
	parentZIndex: number | null
	childZIndices: number[]
}

/**
 * Flatten a shape list depth-first in document order: each shape is followed by
 * its descendants, which is both paint order and the order the
 * {@link PptxSlideElement.zIndex} contract promises. Every entry's index in `out`
 * is its zIndex. Returns the zIndices assigned at this level, which is what the
 * enclosing group records as its {@link PptxSlideElement.childZIndices}.
 */
function harvest(shapes: AnyShape[], parentZIndex: number | null, out: HarvestedShape[]): number[] {
	const zIndices: number[] = []
	for (const shape of shapes) {
		const entry: HarvestedShape = { shape, parentZIndex, childZIndices: [] }
		const zIndex = out.push(entry) - 1
		zIndices.push(zIndex)
		if (shape.shapeType === 'group') entry.childZIndices = harvest(shape.shapes, zIndex, out)
	}
	return zIndices
}

/**
 * Turn harvested shapes into elements, dropping the ones with no resolvable slide
 * position. A dropped shape keeps its zIndex — the numbers are positions in the
 * shape tree, and renumbering them would silently repoint every `parentZIndex`.
 */
function flatten(harvested: HarvestedShape[], slidePath: string): PptxSlideElement[] {
	return harvested
		.map((entry, zIndex) => toElement(entry, zIndex, slidePath))
		.filter((element): element is PptxSlideElement => Boolean(element))
}

function toElement(entry: HarvestedShape, zIndex: number, slidePath: string): PptxSlideElement | null {
	const { shape } = entry
	const frame = shape.absoluteFrame
	if (!frame) {
		reportUnresolvable(shape, zIndex, slidePath)
		return null
	}

	// Only `p:sp` holds a text body, so every other kind projects as an empty frame.
	const textFrame = shape.textFrame
	const paragraphs = textFrame ? textFrame.paragraphs.map((p) => ({ runs: p.runs.map(toRun) })) : []
	const textRuns = paragraphs.flatMap((paragraph) => paragraph.runs)
	// A graphic frame has no text body of its own; its text comes from the structure
	// it hosts, and arrives as a plain string with no runs behind it.
	const runText = textRuns.map((run) => run.text).join('')
	const raw = shape.shapeType === 'graphicFrame' ? graphicFrameText(shape) : runText
	// Deliberately not `TextFrame.text`: this is one whitespace-collapsed line for
	// matching and word-counting, not the frame's text with its line structure.
	const text = raw.replace(/\s+/g, ' ').trim()
	const kind: PptxSlideElementKind =
		shape.shapeType === 'group'
			? 'group'
			: shape.shapeType === 'graphicFrame'
				? 'graphicFrame'
				: text
					? 'text'
					: shape.shapeType === 'picture'
						? 'image'
						: 'shape'

	return {
		id: shape.id ?? zIndex + 1,
		name: shape.name || `${kind} ${zIndex + 1}`,
		kind,
		graphicKind: shape.shapeType === 'graphicFrame' ? graphicKindOf(shape) : null,
		zIndex,
		box: {
			x: emuToInches(frame.left),
			y: emuToInches(frame.top),
			w: emuToInches(frame.width),
			h: emuToInches(frame.height),
		},
		rotation: frame.rotation,
		flipH: frame.flipH,
		flipV: frame.flipV,
		parentZIndex: entry.parentZIndex,
		childZIndices: entry.childZIndices,
		text,
		textWrap: textFrame?.bodyProperties?.wrap ?? null,
		autofit: textFrame?.autofit ?? null,
		autofitFontScale: textFrame?.autofitFontScale ?? null,
		bodyInsets: readBodyInsets(textFrame),
		textRuns,
		paragraphs,
		fontSizes: [...new Set(textRuns.map((run) => run.fontSizePt).filter((size): size is number => size !== null))],
		colors: [...new Set(textRuns.map((run) => run.color).filter((color): color is string => Boolean(color)))],
		fill: shape.fillColor,
		line: shape.lineColor,
		shapeType: shape.presetGeometry,
	}
}

/**
 * What a graphic frame hosts, from the same `a:graphicData/@uri` the read model's
 * `has*` predicates compare against. A frame that answers none of them is `'other'`
 * rather than dropped: a box the deck's author placed is on the slide whether or not
 * this library models its payload, and an overlap or coverage check that skipped it
 * would report a gap where a 3D model or an OLE object sits.
 */
function graphicKindOf(frame: GraphicFrame): PptxGraphicKind {
	if (frame.hasTable) return 'table'
	if (frame.hasChart) return 'chart'
	if (frame.hasChartEx) return 'chartEx'
	if (frame.hasDiagram) return 'diagram'
	return 'other'
}

/**
 * The text a graphic frame puts on the slide: table cells in row order, or SmartArt
 * node text. Charts contribute nothing, matching `Slide.text` on the read model — a
 * chart's strings are data labels and axis titles that PowerPoint itself does not
 * treat as slide body text.
 *
 * Separators are single spaces rather than the read model's `	`/`
`, because the
 * caller collapses whitespace anyway; the point here is only that cells do not run
 * together into one word.
 */
function graphicFrameText(frame: GraphicFrame): string {
	const table = frame.table
	if (table) return table.rows.map((row) => row.cells.map((cell) => cell.text).join(' ')).join(' ')
	return frame.diagram?.text ?? ''
}

/** Project one read-model {@link Run} onto the flat run shape. */
function toRun(run: Run): PptxTextRun {
	return {
		text: run.text,
		fontSizePt: run.fontSizePt,
		color: run.color,
		fontFace: run.fontName,
		// The read model distinguishes "explicitly not bold" from "unset, inherits";
		// this surface reports what renders, so an unset flag is `false`.
		bold: run.bold ?? false,
		italic: run.italic ?? false,
		strike: run.strike,
		highlight: run.highlight?.hex ?? null,
		charSpacingPt: run.charSpacingPt,
	}
}

// PowerPoint body-inset defaults (ECMA-376 §21.1.2.1.1 prose; the XSD leaves
// lIns/tIns/rIns/bIns optional with no schema default): 0.1in left/right,
// 0.05in top/bottom. The read model reports only the explicitly-set sides.
const DEFAULT_INSET_LR_IN = 0.1
const DEFAULT_INSET_TB_IN = 0.05
const POINTS_PER_INCH = 72

function readBodyInsets(textFrame: TextFrame | null): PptxBodyInsets | null {
	const insets = textFrame?.bodyProperties?.insetsPt
	if (!insets) return null
	const inches = (pt: number | undefined, fallback: number): number =>
		pt === undefined ? fallback : pt / POINTS_PER_INCH
	return {
		left: inches(insets.left, DEFAULT_INSET_LR_IN),
		top: inches(insets.top, DEFAULT_INSET_TB_IN),
		right: inches(insets.right, DEFAULT_INSET_LR_IN),
		bottom: inches(insets.bottom, DEFAULT_INSET_TB_IN),
	}
}

/* ────────────────────────────────────────────────────────────────────────────
 * Unresolvable positions
 *
 * `absoluteFrame` reports `null` for three different situations, and only two of
 * them are worth telling the caller about. The distinction is not recoverable
 * from the `null`, so it is re-derived here from the same ancestry the read model
 * walks.
 * ──────────────────────────────────────────────────────────────────────────── */

/** Warn about a shape whose slide position cannot be resolved, unless it never claimed one. */
function reportUnresolvable(shape: AnyShape, zIndex: number, slidePath: string): void {
	const element = shape.element_
	// No own `a:xfrm` at all: the shape inherits its geometry from a layout
	// placeholder, which this surface deliberately does not resolve. Nothing is
	// wrong with the deck, so nothing is reported.
	if (!hasBox(ownTransform(element), 'a:off', 'a:ext')) return

	const who = shape.name ? `"${shape.name}"` : `the element at zIndex ${zIndex}`
	if (enclosingGroupLacksTransform(element)) {
		warn(
			'inspect/group-transform-missing',
			`inspect: skipped ${who} on ${slidePath} — an enclosing group has no usable a:xfrm, so its slide position cannot be resolved.`
		)
		return
	}
	warn(
		'inspect/group-transform-degenerate',
		`inspect: skipped ${who} on ${slidePath} — an enclosing group has a degenerate transform (zero a:chExt), so its slide position cannot be resolved.`
	)
}

/** The element's own `a:xfrm` — on `p:grpSpPr` for a group, `p:spPr` for every other kind. */
function ownTransform(element: Element): Element | null {
	const props = firstChild(element, element.localName === 'grpSp' ? 'p:grpSpPr' : 'p:spPr')
	return props && firstChild(props, 'a:xfrm')
}

/** Whether any enclosing `p:grpSp` fails to state a complete child-space mapping. */
function enclosingGroupLacksTransform(element: Element): boolean {
	for (let node = element.parentNode; node && node.nodeType === ELEMENT_NODE; node = node.parentNode) {
		const parent = node as Element
		if (parent.namespaceURI !== OOXML_NS.p || parent.localName !== 'grpSp') return false // reached the shape tree
		const xfrm = ownTransform(parent)
		if (!hasBox(xfrm, 'a:off', 'a:ext') || !hasBox(xfrm, 'a:chOff', 'a:chExt')) return true
	}
	return false
}

/** Whether a transform states a complete position + extent pair. */
function hasBox(xfrm: Element | null, offName: string, extName: string): boolean {
	if (!xfrm) return false
	const off = firstChild(xfrm, offName)
	const ext = firstChild(xfrm, extName)
	if (!off || !ext) return false
	return (
		intValue(attr(off, 'x')) !== null &&
		intValue(attr(off, 'y')) !== null &&
		intValue(attr(ext, 'cx')) !== null &&
		intValue(attr(ext, 'cy')) !== null
	)
}

function countWords(text: string): number {
	return text.split(/\s+/).filter(Boolean).length
}

function round(value: number, digits: number): number {
	const factor = 10 ** digits
	return Math.round(value * factor) / factor
}

// Error taxonomy — see `entry-errors.ts`. Re-exported from every entry so `instanceof`
// works whichever subpath a consumer imports.
export * from './entry-errors.js'
