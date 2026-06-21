import { XMLParser } from 'fast-xml-parser'
import { readZip } from './zip.js'
import { STANDARD_LAYOUTS, emuToInches } from './units.js'

type XmlNode = Record<string, unknown>

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
}

export type PptxSlideElementKind = 'text' | 'image' | 'shape'

/**
 * Vertical-autofit mode of a text frame, read from the `a:bodyPr` child element:
 * - `'none'`: no autofit (`a:noAutofit`, or no autofit child at all). The box has a
 *   fixed height the text must fit inside — a genuine overflow candidate.
 * - `'normAutofit'`: shrink text to fit (`a:normAutofit`, PptxGenJS `fit: 'shrink'`).
 *   Text is downscaled rather than overflowing.
 * - `'spAutoFit'`: resize shape to fit text (`a:spAutoFit`, PptxGenJS `fit: 'resize'`).
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
	zIndex: number
	box: PptxBox
	text: string
	textWrap: string | null
	autofit: PptxAutofitMode | null
	bodyInsets: PptxBodyInsets | null
	textRuns: PptxTextRun[]
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

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: '',
	allowBooleanAttributes: true,
	parseAttributeValue: true,
	parseTagValue: false,
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

	const parsed = asNode(parser.parse(presentationXml))
	const presentation = nodeChild(parsed, 'p:presentation')
	const size = nodeChild(presentation, 'p:sldSz')
	const cx = numericValue(size?.cx)
	const cy = numericValue(size?.cy)
	if (cx === null || cy === null) return fallback

	return {
		widthIn: round(emuToInches(cx), 3),
		heightIn: round(emuToInches(cy), 3),
	}
}

export async function extractSlides(
	pptxPackage: PptxPackage,
	size?: PptxSlideSize
): Promise<PptxSlideInspection[]> {
	const slideSize = size || (await readPresentationSize(pptxPackage))
	const slidePaths = listPptxParts(pptxPackage)
		.filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path))
		.sort((a, b) => slideNumberFromPath(a) - slideNumberFromPath(b))

	const slides: PptxSlideInspection[] = []
	for (const [index, path] of slidePaths.entries()) {
		const xml = await readPptxTextPart(pptxPackage, path)
		if (!xml) continue

		const parsed = asNode(parser.parse(xml))
		const root = nodeChild(parsed, 'p:sld')
		const cSld = nodeChild(root, 'p:cSld')
		const elements = collectElements(nodeChild(cSld, 'p:spTree'))
			.map((el, zIndex) => normalizeElement(el, zIndex))
			.filter((element): element is PptxSlideElement => Boolean(element))
		const text = elements.map(el => el.text).filter(Boolean).join(' ')

		slides.push({
			index,
			name: stringValue(cSld?.name) || `Slide ${index + 1}`,
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

function collectElements(node: XmlNode | null): XmlNode[] {
	const elements: XmlNode[] = []
	walk(node, candidate => {
		for (const key of ['p:sp', 'p:pic', 'p:cxnSp']) {
			for (const child of asArray(candidate[key])) {
				const element = asNode(child)
				if (element) elements.push(element)
			}
		}
	})
	return elements
}

function normalizeElement(node: XmlNode, zIndex: number): PptxSlideElement | null {
	const spPr = nodeChild(node, 'p:spPr')
	const box = readBox(nodeChild(spPr, 'a:xfrm'))
	if (!box) return null

	const textBody = nodeChild(node, 'p:txBody')
	const textRuns = extractTextRuns(textBody)
	const text = textRuns.map(run => run.text).join('').replace(/\s+/g, ' ').trim()
	const kind: PptxSlideElementKind = text ? 'text' : nodeChild(node, 'p:blipFill') ? 'image' : 'shape'

	return {
		id:
			stringOrNumberValue(cNvPr(node, 'p:nvSpPr')?.id) ||
			stringOrNumberValue(cNvPr(node, 'p:nvPicPr')?.id) ||
			stringOrNumberValue(cNvPr(node, 'p:nvCxnSpPr')?.id) ||
			zIndex + 1,
		name:
			stringValue(cNvPr(node, 'p:nvSpPr')?.name) ||
			stringValue(cNvPr(node, 'p:nvPicPr')?.name) ||
			stringValue(cNvPr(node, 'p:nvCxnSpPr')?.name) ||
			`${kind} ${zIndex + 1}`,
		kind,
		zIndex,
		box,
		text,
		textWrap: readTextWrap(textBody),
		autofit: readAutofit(textBody),
		bodyInsets: readBodyInsets(textBody),
		textRuns,
		fontSizes: [...new Set(textRuns.map(run => run.fontSizePt).filter((size): size is number => size !== null))],
		colors: [...new Set(textRuns.map(run => run.color).filter((color): color is string => Boolean(color)))],
		fill: readFill(spPr),
		line: readLine(spPr),
		shapeType: stringValue(nodeChild(spPr, 'a:prstGeom')?.prst),
	}
}

function cNvPr(node: XmlNode, parentKey: string): XmlNode | null {
	return nodeChild(nodeChild(node, parentKey), 'p:cNvPr')
}

function readBox(xfrm: XmlNode | null): PptxBox | null {
	const off = nodeChild(xfrm, 'a:off')
	const ext = nodeChild(xfrm, 'a:ext')
	if (!off || !ext) return null
	return {
		x: emuToInches(numericValue(off.x) || 0),
		y: emuToInches(numericValue(off.y) || 0),
		w: emuToInches(numericValue(ext.cx) || 0),
		h: emuToInches(numericValue(ext.cy) || 0),
	}
}

function extractTextRuns(textBody: XmlNode | null): PptxTextRun[] {
	if (!textBody) return []
	const runs: PptxTextRun[] = []
	walk(textBody, node => {
		for (const item of asArray(node['a:r'])) {
			const run = asNode(item)
			if (!run) continue
			const text = stringValue(run['a:t'])
			if (text === null) continue
			const props = nodeChild(run, 'a:rPr')
			runs.push({
				text,
				fontSizePt: props?.sz ? Number(props.sz) / 100 : null,
				color: readTextColor(props),
			})
		}
	})
	return runs
}

function readTextColor(props: XmlNode | null): string | null {
	return stringValue(nodeChild(nodeChild(props, 'a:solidFill'), 'a:srgbClr')?.val)
}

function readFill(spPr: XmlNode | null): string | null {
	return stringValue(nodeChild(nodeChild(spPr, 'a:solidFill'), 'a:srgbClr')?.val)
}

function readLine(spPr: XmlNode | null): string | null {
	return stringValue(nodeChild(nodeChild(nodeChild(spPr, 'a:ln'), 'a:solidFill'), 'a:srgbClr')?.val)
}

function readTextWrap(textBody: XmlNode | null): string | null {
	return stringValue(nodeChild(textBody, 'a:bodyPr')?.wrap)
}

// PowerPoint body-inset defaults (ECMA-376 §21.1.2.1.1 prose; the XSD leaves
// lIns/tIns/rIns/bIns optional with no schema default): 0.1in left/right, 0.05in top/bottom.
const DEFAULT_INSET_LR_EMU = 91440
const DEFAULT_INSET_TB_EMU = 45720

function readAutofit(textBody: XmlNode | null): PptxAutofitMode | null {
	const bodyPr = nodeChild(textBody, 'a:bodyPr')
	if (!bodyPr) return null
	if ('a:spAutoFit' in bodyPr) return 'spAutoFit'
	if ('a:normAutofit' in bodyPr) return 'normAutofit'
	return 'none'
}

function readBodyInsets(textBody: XmlNode | null): PptxBodyInsets | null {
	const bodyPr = nodeChild(textBody, 'a:bodyPr')
	if (!bodyPr) return null
	return {
		left: emuToInches(numericValue(bodyPr.lIns) ?? DEFAULT_INSET_LR_EMU),
		top: emuToInches(numericValue(bodyPr.tIns) ?? DEFAULT_INSET_TB_EMU),
		right: emuToInches(numericValue(bodyPr.rIns) ?? DEFAULT_INSET_LR_EMU),
		bottom: emuToInches(numericValue(bodyPr.bIns) ?? DEFAULT_INSET_TB_EMU),
	}
}

function walk(value: unknown, visitor: (node: XmlNode) => void): void {
	if (Array.isArray(value)) {
		for (const child of value) walk(child, visitor)
		return
	}
	const node = asNode(value)
	if (!node) return
	visitor(node)
	for (const child of Object.values(node)) walk(child, visitor)
}

function asArray(value: unknown): unknown[] {
	if (value === undefined || value === null) return []
	return Array.isArray(value) ? value : [value]
}

function nodeChild(node: XmlNode | null | undefined, key: string): XmlNode | null {
	return asNode(node?.[key])
}

function asNode(value: unknown): XmlNode | null {
	return typeof value === 'object' && value !== null && !Array.isArray(value) ? value as XmlNode : null
}

function numericValue(value: unknown): number | null {
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
