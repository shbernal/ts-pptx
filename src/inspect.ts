import { XMLParser } from 'fast-xml-parser'
import { composeGroupFrame, type GroupTransform, type TransformBox } from './group-transform.js'
import { warn } from './log.js'
import { readZip } from './zip.js'
import { STANDARD_LAYOUTS, emuToInches } from './units.js'

/**
 * Input to the inspect surface. A `string` is a **filesystem path** (Node) read
 * from disk; pass `Uint8Array`/`ArrayBuffer`/`Blob`/`number[]` for an in-memory
 * archive. (A string is a path, not JSZip's latin1 binary content — see
 * {@link ZipInput}.)
 */
type PptxInspectInputValue = string | number[] | Uint8Array | ArrayBuffer | Blob

export type PptxInspectInput = PptxInspectInputValue | Promise<PptxInspectInputValue>

export interface PptxPackageFile {
	async(type: 'string'): Promise<string>
	async(type: 'uint8array'): Promise<Uint8Array>
}

export interface PptxPackage {
	files: Record<string, unknown>
	file(path: string): PptxPackageFile | null
}

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
	/** Highlight colour hex (`a:rPr > a:highlight > a:srgbClr@val`), or null when unset (theme tokens are not resolved here). */
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
 * in {@link PptxSlideElement.parentZIndex}.
 */
export type PptxSlideElementKind = 'text' | 'image' | 'shape' | 'group'

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

// `preserveOrder` keeps sibling order across *different* tag names, which the default
// tag-keyed output discards ({a:[…], b:{…}} cannot say whether `b` came first). zIndex
// is document order, so that ordering is load-bearing. It costs a more awkward parse
// shape, which `toElements` normalises away immediately.
const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: '',
	allowBooleanAttributes: true,
	parseAttributeValue: true,
	parseTagValue: false,
	preserveOrder: true,
})

const textDecoder = new TextDecoder('utf-8')

export async function loadPptxPackage(input: PptxInspectInput): Promise<PptxPackage> {
	const entries = await readZip(input)
	const files: Record<string, unknown> = {}
	for (const path of entries.keys()) files[path] = true
	return {
		files,
		file(path: string): PptxPackageFile | null {
			const bytes = entries.get(path)
			if (!bytes) return null
			const read = (async (type: 'string' | 'uint8array') =>
				type === 'uint8array' ? bytes : textDecoder.decode(bytes)) as PptxPackageFile['async']
			return { async: read }
		},
	}
}

export function listPptxParts(pptxPackage: PptxPackage): string[] {
	return Object.keys(pptxPackage.files).sort()
}

export async function readPptxTextPart(pptxPackage: PptxPackage, path: string): Promise<string | null> {
	const entry = pptxPackage.file(path)
	return entry ? entry.async('string') : null
}

/**
 * Read a package part as raw bytes — the binary sibling of {@link readPptxTextPart}
 * for embedded media (SVG/PNG/EMF blobs, fonts, …) that must not be UTF-8 decoded.
 * Returns `null` when the part is absent. The `Uint8Array` is browser-isomorphic;
 * Node consumers can wrap it with `Buffer.from(...)` if they need Buffer methods.
 */
export async function readPptxBinaryPart(pptxPackage: PptxPackage, path: string): Promise<Uint8Array | null> {
	const entry = pptxPackage.file(path)
	return entry ? entry.async('uint8array') : null
}

export async function inspectPptx(input: PptxInspectInput): Promise<PptxInspection> {
	const pptxPackage = await loadPptxPackage(input)
	const slideSize = await readPresentationSize(pptxPackage)
	const slides = await extractSlides(pptxPackage, slideSize)
	return { slideSize, slides }
}

export async function readPresentationSize(
	pptxPackage: PptxPackage,
	fallback: PptxSlideSize = DEFAULT_INSPECT_SLIDE_SIZE
): Promise<PptxSlideSize> {
	const presentationXml = await readPptxTextPart(pptxPackage, 'ppt/presentation.xml')
	if (!presentationXml) return fallback

	const presentation = firstChild(toElements(parser.parse(presentationXml)), 'p:presentation')
	const size = child(presentation, 'p:sldSz')
	const cx = numericValue(attr(size, 'cx'))
	const cy = numericValue(attr(size, 'cy'))
	if (cx === null || cy === null) return fallback

	return {
		widthIn: round(emuToInches(cx), 3),
		heightIn: round(emuToInches(cy), 3),
	}
}

export async function extractSlides(pptxPackage: PptxPackage, size?: PptxSlideSize): Promise<PptxSlideInspection[]> {
	const slideSize = size || (await readPresentationSize(pptxPackage))
	const slidePaths = listPptxParts(pptxPackage)
		.filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
		.sort((a, b) => slideNumberFromPath(a) - slideNumberFromPath(b))

	const slides: PptxSlideInspection[] = []
	for (const [index, path] of slidePaths.entries()) {
		const xml = await readPptxTextPart(pptxPackage, path)
		if (!xml) continue

		const root = firstChild(toElements(parser.parse(xml)), 'p:sld')
		const cSld = child(root, 'p:cSld')
		const elements = normalizeElements(collectElements(child(cSld, 'p:spTree')), path)
		const text = elements
			.map((el) => el.text)
			.filter(Boolean)
			.join(' ')

		slides.push({
			index,
			name: stringValue(attr(cSld, 'name')) || `Slide ${index + 1}`,
			path,
			size: slideSize,
			elements,
			text,
			wordCount: countWords(text),
		})
	}

	return slides
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

function slideNumberFromPath(path: string): number {
	return Number(path.match(/slide(\d+)\.xml$/)?.[1] || 0)
}

/* ────────────────────────────────────────────────────────────────────────────
 * Shape-tree harvest
 * ──────────────────────────────────────────────────────────────────────────── */

const SHAPE_TAGS = new Set(['p:sp', 'p:pic', 'p:cxnSp'])

/**
 * One harvested shape-tree entry, before it is turned into a {@link PptxSlideElement}.
 * `groups` is the enclosing `p:grpSp` chain **innermost first**, as
 * {@link composeGroupFrame} wants it, or `null` when some enclosing group has no
 * usable transform and the element's slide position is therefore unresolvable.
 */
interface HarvestedElement {
	element: XmlElement
	isGroup: boolean
	groups: GroupTransform[] | null
	parentZIndex: number | null
	childZIndices: number[]
}

/**
 * Flatten `p:spTree` into a depth-first, document-order list: each element is
 * followed by its descendants, which is both paint order and the order the
 * {@link PptxSlideElement.zIndex} contract promises. Every entry's index in the
 * returned array is its zIndex.
 */
function collectElements(spTree: XmlElement | null): HarvestedElement[] {
	const harvested: HarvestedElement[] = []
	harvest(spTree, [], null, harvested) // slide level: no enclosing groups, not an unresolvable chain
	return harvested
}

/** Harvest `parent`'s shape children, returning their zIndices in document order. */
function harvest(
	parent: XmlElement | null,
	groups: GroupTransform[] | null,
	parentZIndex: number | null,
	out: HarvestedElement[]
): number[] {
	const zIndices: number[] = []
	if (!parent) return zIndices

	for (const element of shapeTreeChildren(parent)) {
		if (element.tag === 'p:grpSp') {
			const entry: HarvestedElement = { element, isGroup: true, groups, parentZIndex, childZIndices: [] }
			const zIndex = out.push(entry) - 1
			zIndices.push(zIndex)
			// The group's own transform maps its children; if it is missing or incomplete,
			// nothing below it can be placed, so the whole subtree inherits `null`.
			const own = groups && readGroupTransform(element)
			entry.childZIndices = harvest(element, own ? [own, ...groups] : null, zIndex, out)
			continue
		}
		if (!SHAPE_TAGS.has(element.tag)) continue // p:nvGrpSpPr, p:grpSpPr, p:graphicFrame, p:extLst, …
		zIndices.push(out.push({ element, isGroup: false, groups, parentZIndex, childZIndices: [] }) - 1)
	}
	return zIndices
}

/**
 * `parent`'s children in document order, descending transparently through
 * `mc:AlternateContent`. PowerPoint wraps a version-gated shape in one, where
 * `mc:Choice` and `mc:Fallback` are two renderings of the *same* shape — take a
 * single branch, or every such shape is reported twice.
 */
function shapeTreeChildren(parent: XmlElement): XmlElement[] {
	const children: XmlElement[] = []
	for (const element of parent.children) {
		if (element.tag === 'mc:AlternateContent') {
			const branch = child(element, 'mc:Choice') || child(element, 'mc:Fallback')
			if (branch) children.push(...shapeTreeChildren(branch))
			continue
		}
		children.push(element)
	}
	return children
}

/** A `p:grpSp`'s child-space mapping (`p:grpSpPr > a:xfrm`), or null when it has no usable transform. */
function readGroupTransform(group: XmlElement): GroupTransform | null {
	const xfrm = child(child(group, 'p:grpSpPr'), 'a:xfrm')
	if (!xfrm) return null
	const outer = readTransformBox(xfrm, 'a:off', 'a:ext')
	const childSpace = readTransformBox(xfrm, 'a:chOff', 'a:chExt')
	if (!outer || !childSpace) return null
	return { outer, child: childSpace, ...orientationOf(xfrm) }
}

function normalizeElements(harvested: HarvestedElement[], slidePath: string): PptxSlideElement[] {
	return harvested
		.map((entry, zIndex) => normalizeElement(entry, zIndex, slidePath))
		.filter((element): element is PptxSlideElement => Boolean(element))
}

function normalizeElement(entry: HarvestedElement, zIndex: number, slidePath: string): PptxSlideElement | null {
	const { element, isGroup } = entry
	// A group's own frame lives on p:grpSpPr; every other kind carries p:spPr.
	const props = child(element, isGroup ? 'p:grpSpPr' : 'p:spPr')
	const xfrm = child(props, 'a:xfrm')
	const own = xfrm && readTransformBox(xfrm, 'a:off', 'a:ext')
	if (!xfrm || !own) return null // no own transform (e.g. geometry inherited from a layout placeholder)

	const name = stringValue(attr(cNvPr(element), 'name'))
	if (!entry.groups) {
		warn(
			`inspect: skipped ${describe(name, zIndex)} on ${slidePath} — an enclosing group has no usable a:xfrm, so its slide position cannot be resolved.`
		)
		return null
	}
	const frame = composeGroupFrame({ box: own, ...orientationOf(xfrm) }, entry.groups)
	if (!frame) {
		warn(
			`inspect: skipped ${describe(name, zIndex)} on ${slidePath} — an enclosing group has a degenerate transform (zero a:chExt), so its slide position cannot be resolved.`
		)
		return null
	}

	const textBody = child(element, 'p:txBody')
	const textRuns = extractTextRuns(textBody)
	const text = textRuns
		.map((run) => run.text)
		.join('')
		.replace(/\s+/g, ' ')
		.trim()
	const kind: PptxSlideElementKind = isGroup
		? 'group'
		: text
			? 'text'
			: child(element, 'p:blipFill')
				? 'image'
				: 'shape'

	return {
		id: stringOrNumberValue(attr(cNvPr(element), 'id')) ?? zIndex + 1,
		name: name || `${kind} ${zIndex + 1}`,
		kind,
		zIndex,
		box: {
			x: emuToInches(frame.box.x),
			y: emuToInches(frame.box.y),
			w: emuToInches(frame.box.cx),
			h: emuToInches(frame.box.cy),
		},
		rotation: frame.rotation,
		flipH: frame.flipH,
		flipV: frame.flipV,
		parentZIndex: entry.parentZIndex,
		childZIndices: entry.childZIndices,
		text,
		textWrap: readTextWrap(textBody),
		autofit: readAutofit(textBody),
		autofitFontScale: readAutofitFontScale(textBody),
		bodyInsets: readBodyInsets(textBody),
		textRuns,
		paragraphs: extractParagraphs(textBody),
		fontSizes: [...new Set(textRuns.map((run) => run.fontSizePt).filter((size): size is number => size !== null))],
		colors: [...new Set(textRuns.map((run) => run.color).filter((color): color is string => Boolean(color)))],
		fill: readFill(props),
		line: readLine(props),
		shapeType: stringValue(attr(child(props, 'a:prstGeom'), 'prst')),
	}
}

/** Identify an element in a warning: its authored name when it has one, else its z position. */
function describe(name: string | null, zIndex: number): string {
	return name ? `"${name}"` : `the element at zIndex ${zIndex}`
}

const NON_VISUAL_PROPS = ['p:nvSpPr', 'p:nvPicPr', 'p:nvCxnSpPr', 'p:nvGrpSpPr']

/** The `p:cNvPr` identity element, whichever non-visual-properties wrapper this kind uses. */
function cNvPr(element: XmlElement): XmlElement | null {
	for (const tag of NON_VISUAL_PROPS) {
		const found = child(child(element, tag), 'p:cNvPr')
		if (found) return found
	}
	return null
}

/** A position + extent pair off a transform, in EMU; null when either is incomplete. */
function readTransformBox(xfrm: XmlElement, offName: string, extName: string): TransformBox | null {
	const off = child(xfrm, offName)
	const ext = child(xfrm, extName)
	const x = numericValue(attr(off, 'x'))
	const y = numericValue(attr(off, 'y'))
	const cx = numericValue(attr(ext, 'cx'))
	const cy = numericValue(attr(ext, 'cy'))
	if (x === null || y === null || cx === null || cy === null) return null
	return { x, y, cx, cy }
}

/** A transform's orientation: `@rot` in degrees (authored in 60000ths) plus the flip flags. */
function orientationOf(xfrm: XmlElement): { rotation: number; flipH: boolean; flipV: boolean } {
	return {
		rotation: (numericValue(attr(xfrm, 'rot')) ?? 0) / 60000,
		flipH: readXmlBool(attr(xfrm, 'flipH')),
		flipV: readXmlBool(attr(xfrm, 'flipV')),
	}
}

/** Read one `<a:r>` run's text + character properties, or null if it has no text node. */
function readTextRun(run: XmlElement): PptxTextRun | null {
	const textNode = child(run, 'a:t')
	if (!textNode) return null
	const props = child(run, 'a:rPr')
	const spc = numericValue(attr(props, 'spc'))
	const size = numericValue(attr(props, 'sz'))
	return {
		text: textNode.text,
		fontSizePt: size === null ? null : size / 100,
		color: readTextColor(props),
		fontFace: stringValue(attr(child(props, 'a:latin'), 'typeface')),
		bold: readXmlBool(attr(props, 'b')),
		italic: readXmlBool(attr(props, 'i')),
		strike: stringValue(attr(props, 'strike')),
		highlight: stringValue(attr(child(child(props, 'a:highlight'), 'a:srgbClr'), 'val')),
		charSpacingPt: spc === null ? null : spc / 100,
	}
}

function extractTextRuns(textBody: XmlElement | null): PptxTextRun[] {
	if (!textBody) return []
	const runs: PptxTextRun[] = []
	walk(textBody, (node) => {
		if (node.tag !== 'a:r') return
		const built = readTextRun(node)
		if (built) runs.push(built)
	})
	return runs
}

/** Runs grouped by their source paragraph (`a:p`), preserving line boundaries. */
function extractParagraphs(textBody: XmlElement | null): PptxParagraph[] {
	if (!textBody) return []
	return children(textBody, 'a:p').map((paragraph) => ({
		runs: children(paragraph, 'a:r')
			.map((run) => readTextRun(run))
			.filter((run): run is PptxTextRun => Boolean(run)),
	}))
}

/** OOXML boolean attributes arrive as 1/0 or "true"/"false" depending on the writer. */
function readXmlBool(value: unknown): boolean {
	return value === 1 || value === '1' || value === true || value === 'true'
}

function readTextColor(props: XmlElement | null): string | null {
	return stringValue(attr(child(child(props, 'a:solidFill'), 'a:srgbClr'), 'val'))
}

function readFill(props: XmlElement | null): string | null {
	return stringValue(attr(child(child(props, 'a:solidFill'), 'a:srgbClr'), 'val'))
}

function readLine(props: XmlElement | null): string | null {
	return stringValue(attr(child(child(child(props, 'a:ln'), 'a:solidFill'), 'a:srgbClr'), 'val'))
}

function readTextWrap(textBody: XmlElement | null): string | null {
	return stringValue(attr(child(textBody, 'a:bodyPr'), 'wrap'))
}

// PowerPoint body-inset defaults (ECMA-376 §21.1.2.1.1 prose; the XSD leaves
// lIns/tIns/rIns/bIns optional with no schema default): 0.1in left/right, 0.05in top/bottom.
const DEFAULT_INSET_LR_EMU = 91440
const DEFAULT_INSET_TB_EMU = 45720

function readAutofit(textBody: XmlElement | null): PptxAutofitMode | null {
	const bodyPr = child(textBody, 'a:bodyPr')
	if (!bodyPr) return null
	if (child(bodyPr, 'a:spAutoFit')) return 'spAutoFit'
	if (child(bodyPr, 'a:normAutofit')) return 'normAutofit'
	return 'none'
}

/** Baked `<a:normAutofit@fontScale>` as a percent (62.5 = 62.5%), or null when unset. */
function readAutofitFontScale(textBody: XmlElement | null): number | null {
	const norm = child(child(textBody, 'a:bodyPr'), 'a:normAutofit')
	if (!norm) return null
	// OOXML stores fontScale in 1000ths of a percent (62500 = 62.5%). Return a percent.
	const raw = numericValue(attr(norm, 'fontScale'))
	return raw === null ? null : raw / 1000
}

function readBodyInsets(textBody: XmlElement | null): PptxBodyInsets | null {
	const bodyPr = child(textBody, 'a:bodyPr')
	if (!bodyPr) return null
	return {
		left: emuToInches(numericValue(attr(bodyPr, 'lIns')) ?? DEFAULT_INSET_LR_EMU),
		top: emuToInches(numericValue(attr(bodyPr, 'tIns')) ?? DEFAULT_INSET_TB_EMU),
		right: emuToInches(numericValue(attr(bodyPr, 'rIns')) ?? DEFAULT_INSET_LR_EMU),
		bottom: emuToInches(numericValue(attr(bodyPr, 'bIns')) ?? DEFAULT_INSET_TB_EMU),
	}
}

/* ────────────────────────────────────────────────────────────────────────────
 * XML tree model
 *
 * `preserveOrder` hands back entries shaped like `{ 'p:sp': [...children], ':@': {attrs} }`
 * with text as `{ '#text': '…' }` siblings. `toElements` converts that once into a
 * plain element tree so the readers above stay about OOXML rather than about the parser.
 * ──────────────────────────────────────────────────────────────────────────── */

interface XmlElement {
	/** Qualified tag name as authored, e.g. `p:sp`. */
	tag: string
	attrs: Record<string, unknown>
	/** Child *elements* in document order; text nodes are folded into {@link text}. */
	children: XmlElement[]
	/** Concatenated direct text content (`<a:t>` bodies); `''` for elements with none. */
	text: string
}

function toElements(entries: unknown): XmlElement[] {
	const elements: XmlElement[] = []
	for (const item of asArray(entries)) {
		const entry = asNode(item)
		if (!entry) continue
		for (const [tag, value] of Object.entries(entry)) {
			if (tag === ':@' || tag === '#text') continue // attribute bag / text sibling, not an element
			elements.push({
				tag,
				attrs: asNode(entry[':@']) || {},
				children: toElements(value),
				text: directText(value),
			})
		}
	}
	return elements
}

function directText(value: unknown): string {
	let text = ''
	for (const item of asArray(value)) {
		const entry = asNode(item)
		if (entry && '#text' in entry) text += String(entry['#text'])
	}
	return text
}

function firstChild(elements: XmlElement[], tag: string): XmlElement | null {
	return elements.find((element) => element.tag === tag) || null
}

/** First child element of `parent` named `tag`, or null. */
function child(parent: XmlElement | null | undefined, tag: string): XmlElement | null {
	return parent ? firstChild(parent.children, tag) : null
}

/** All direct children of `parent` named `tag`, in document order. */
function children(parent: XmlElement, tag: string): XmlElement[] {
	return parent.children.filter((element) => element.tag === tag)
}

function attr(element: XmlElement | null | undefined, name: string): unknown {
	return element?.attrs[name]
}

/** Visit `element` and every descendant, in document order. */
function walk(element: XmlElement, visitor: (node: XmlElement) => void): void {
	visitor(element)
	for (const node of element.children) walk(node, visitor)
}

function asArray(value: unknown): unknown[] {
	if (value === undefined || value === null) return []
	return Array.isArray(value) ? value : [value]
}

function asNode(value: unknown): Record<string, unknown> | null {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: null
}

function numericValue(value: unknown): number | null {
	if (value === undefined || value === null || value === '') return null
	const number = Number(value)
	return Number.isFinite(number) ? number : null
}

function stringValue(value: unknown): string | null {
	return value === undefined || value === null ? null : String(value)
}

function stringOrNumberValue(value: unknown): string | number | null {
	if (typeof value === 'number' || typeof value === 'string') return value
	return stringValue(value)
}

function countWords(text: string): number {
	return text.split(/\s+/).filter(Boolean).length
}

function round(value: number, digits: number): number {
	const factor = 10 ** digits
	return Math.round(value * factor) / factor
}
