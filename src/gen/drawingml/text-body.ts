/**
 * ts-pptx: DrawingML text body
 *
 * The container layer over `text-run.ts`: body properties (`<a:bodyPr>`), the
 * full `<p:txBody>`/`<a:txBody>` builder (`genXmlTextBody`), the native-equation
 * predicate, and the placeholder `<p:ph>` element.
 */

import { PlaceholderType, SlideObjectType } from '../../enums.js'
import { CRLF } from '../../constants-internal.js'
import { warn } from '../../diagnostics.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import { TEXT_SHAPE_TYPES, TEXT_VERTICAL } from '../../ooxml/st-enums.js'
import type { ObjectOptions, TextProps, TextPropsOptions } from '../../types/index.js'
import type { SlideObject, TableCellInternal } from '../../types/internal.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { TEXT_ANCHOR_BY_VALIGN, type TextAnchorToken } from '../../ooxml/text-anchor.js'
import {
	genXmlNormAutofit,
	genXmlParagraphProperties,
	groupRunsIntoLines,
	renderTextParagraphsXml,
	type RunProps,
} from './text-run.js'

const PLACEHOLDER_TYPE_MAP = PlaceholderType as Record<string, string>

/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {SlideObject | TableCellInternal} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject: SlideObject | TableCellInternal): string {
	// A table cell always emits bare body properties, whatever else is configured. (This used to be
	// a ternary on the return, which built the full element first and then threw it away.)
	if (slideObject._type === SlideObjectType.tablecell) return voidEl('a:bodyPr')

	// Placeholders (incl. master/layout placeholders) carry their margin/valign in `_bodyProp` just
	// like text boxes, so they must emit the same configured `<a:bodyPr>` — otherwise a placeholder
	// authored with insets or a vertical anchor silently degrades to the default.
	// `_bodyProp`/`options` are optional on the type but present on text/placeholder objects that reach
	// this branch; bind them once so the body reads a narrowed, non-undefined value.
	if (
		!slideObject ||
		(slideObject._type !== SlideObjectType.text && slideObject._type !== SlideObjectType.placeholder)
	) {
		// DEFAULT: paired, not self-closing — `<a:bodyPr .../>` would be a byte change.
		return el('a:bodyPr', { wrap: 'square', rtlCol: '0' })
	}
	// A text/placeholder object with no `_bodyProp` still takes the full builder below, and its
	// defaults. It reached that arm by way of the text serializer creating an empty bag on the
	// authored object, which is the kind of write a serializer does not get to make; the empty
	// bag is spelled here instead.
	//
	// `addShape` is who arrives here that way — it builds a `_type === text` object and never
	// writes `_bodyProp`, where `addTextDefinition` always does. So every shape's `wrap` is
	// decided by what an *absent* key means, which is why it is stated rather than read off a
	// truthiness test below.
	const options = slideObject.options
	const bodyProp = options?._bodyProp ?? {}

	// PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>
	// NOTE: attribute ORDER is byte-significant; `rtlCol` sits after the margins and columns but
	// before the anchor points, which is why this is one ordered literal rather than grouped writes.
	const attrs: XmlAttrs = {
		// A: Enable or disable textwrapping none or square. Only an explicit `false` turns wrapping
		// off: `square` is PowerPoint's own default and what `addTextDefinition` writes whenever it
		// runs, so an object that carries no `wrap` at all is one nobody made a decision about —
		// and the un-decided case is the default, not the opposite of it.
		wrap: bodyProp.wrap === false ? 'none' : 'square',
		// B: Textbox margins [padding] — an explicit zero is meaningful, so test for it separately
		lIns: bodyProp.lIns || bodyProp.lIns === 0 ? bodyProp.lIns : null,
		tIns: bodyProp.tIns || bodyProp.tIns === 0 ? bodyProp.tIns : null,
		rIns: bodyProp.rIns || bodyProp.rIns === 0 ? bodyProp.rIns : null,
		bIns: bodyProp.bIns || bodyProp.bIns === 0 ? bodyProp.bIns : null,
		// C.1: Text columns (numCol/spcCol). Spacing is only meaningful when there is more than one column.
		numCol: bodyProp.numCol ? bodyProp.numCol : null,
		spcCol: bodyProp.spcCol ? bodyProp.spcCol : null,
		// C: Add rtl after margins
		rtlCol: '0',
		// D: Add anchorPoints
		anchor: bodyProp.anchor ? bodyProp.anchor : null, // VALS: [t,ctr,b]
		vert: checkEnumOrWarn(bodyProp.vert, TEXT_VERTICAL, 'text/invalid-vertical', 'text: vert'),
	}

	const children: string[] = []

	// E.1: Preset text warp (`<a:prstTxWarp>`). Per CT_TextBodyProperties this child
	// comes before the autofit group, so emit it immediately after the attributes.
	// NOTE: this `<a:avLst/>` has NO space before the slash, unlike the one `custGeom` writes.
	const prstTxWarp = checkEnumOrWarn(bodyProp.prstTxWarp, TEXT_SHAPE_TYPES, 'text/invalid-warp', 'text: textWarp')
	if (prstTxWarp) {
		children.push(el('a:prstTxWarp', { prst: prstTxWarp }, raw(voidEl('a:avLst'))))
	}

	/**
	 * F: Text Fit/AutoFit/Shrink option
	 * @see: http://officeopenxml.com/drwSp-text-bodyPr-fit.php
	 * @see: http://www.datypic.com/sc/ooxml/g-a_EG_TextAutofit.html
	 */
	if (options?.fit) {
		const fit = options.fit
		// NOTE: Use of '<a:noAutofit/>' instead of '' causes issues in PPT-2013! ('none' emits nothing.)
		// NOTE: Bare shrink does not work automatically - PowerPoint calculates fontScale/lnSpcReduction dynamically upon edit/resize.
		// The object form bakes explicit values into the file (MS-PPT > Format shape > Text Options: "Shrink text on overflow").
		if (fit === 'shrink') children.push(voidEl('a:normAutofit'))
		else if (fit === 'resize') children.push(voidEl('a:spAutoFit'))
		else if (typeof fit === 'object' && fit.type === 'shrink') children.push(genXmlNormAutofit(fit))
	}

	return el('a:bodyPr', attrs, children.map(raw))
}

/**
 * Resolve a caller's `valign` to the `a:bodyPr/@anchor` token, or `null` when they stated
 * nothing usable.
 *
 * `null` means "no answer here", not "omit the attribute": each call site then applies its own
 * default — the text definer's is `ctr`, while a table cell and the slide-number placeholder
 * leave the attribute off and inherit.
 *
 * `ST_TextAnchoringType` is an enumeration, so a string outside it makes PowerPoint report the
 * package as needing repair. Typed callers are unaffected — `VAlign` is
 * `'top' | 'middle' | 'bottom'` — and a JavaScript caller passing something else now gets a
 * diagnostic instead of an attribute the schema rejects.
 * @param valign - the caller's `valign`, in any of the spellings above
 */
export function resolveTextAnchor(valign: string | null | undefined): TextAnchorToken | null {
	if (valign === null || valign === undefined || valign === '') return null
	const anchor = TEXT_ANCHOR_BY_VALIGN[String(valign).trim().toLowerCase()]
	if (anchor) return anchor
	warn(
		'text/invalid-valign',
		`valign "${String(valign)}" is not one of top/middle/bottom; leaving the text anchor to inherit.`
	)
	return null
}

export function objectHasMath(slideObj: SlideObject): boolean {
	const text = slideObj.text as TextProps | TextProps[] | string | number | undefined
	if (Array.isArray(text)) return text.some((item) => item && typeof item === 'object' && !!item.math)
	if (text && typeof text === 'object') return !!text.math
	return false
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {SlideObject|TableCellInternal} slideObj - slideObj or tableCell
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragragh-level formatting device
 * @template
 *    <p:txBody>
 *        <a:bodyPr wrap="square" rtlCol="0">
 *            <a:spAutoFit/>
 *        </a:bodyPr>
 *        <a:lstStyle/>
 *        <a:p>
 *            <a:pPr algn="ctr"/>
 *            <a:r>
 *                <a:rPr lang="en-US" dirty="0" err="1"/>
 *                <a:t>textbox text</a:t>
 *            </a:r>
 *            <a:endParaRPr lang="en-US" dirty="0"/>
 *        </a:p>
 *    </p:txBody>
 * @returns XML containing the param object's text and formatting
 */
export function genXmlTextBody(slideObj: SlideObject | TableCellInternal): string {
	const opts: ObjectOptions = slideObj.options || {}
	let tmpTextObjects: TextProps[] = []
	const arrTextObjects: RunProps[] = []

	// FIRST: Shapes without text reach this point with `slideObj.text` null/undefined.
	// We MUST still emit a `<p:txBody>` with at least an empty `<a:p>` paragraph;
	// the empty-txBody fallback below appends `<a:p><a:endParaRPr/></a:p>` when no
	// `<a:p>` was produced. Returning early here would emit `<p:sp>` without
	// `<p:txBody>`, which PowerPoint reports as a needs-repair error.

	// STEP 1: Accumulate the body's children; the `<p:txBody>`/`<a:txBody>` wrapper closes over
	// them in STEP 7.
	let strSlideXml = ''

	// STEP 2: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj)

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// KNOWN LIMITATION: horizontal align on a LINE does not work — text is always left-aligned inside the line.
		if (opts.h === 0 && opts.line && opts.align)
			strSlideXml += el('a:lstStyle', null, raw(voidEl('a:lvl1pPr', { algn: 'l' })))
		else if (slideObj._type === SlideObjectType.placeholder)
			strSlideXml += el('a:lstStyle', null, raw(genXmlParagraphProperties(slideObj, true)))
		else strSlideXml += voidEl('a:lstStyle')
	}

	/* STEP 3: Modify slideObj.text to array
		CASES:
		addText( 'string' ) // string
		addText( 'line1\n line2' ) // string with lineBreak
		addText( {text:'word1'} ) // TextProps object
		addText( ['barry','allen'] ) // array of strings
		addText( [{text:'word1'}, {text:'word2'}] ) // TextProps object array
		addText( [{text:'line1\n line2'}, {text:'end word'}] ) // TextProps object array with lineBreak
	*/
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		// Handle cases 1,2
		tmpTextObjects.push({ text: slideObj.text.toString(), options: opts || {} })
	} else if (
		slideObj.text &&
		!Array.isArray(slideObj.text) &&
		typeof slideObj.text === 'object' &&
		Object.keys(slideObj.text).includes('text')
	) {
		// Handle case 3
		// `math`/`inline` are copied only when stated: this is a model the steps below spread
		// (`{ ...itext.options }`, `{ ...itext }`), so a key holding `undefined` is not the same as
		// an absent one.
		const authored = slideObj.text as TextProps
		const textObject: TextProps = { text: slideObj.text || '', options: slideObj.options || {} }
		if (authored.math !== undefined) textObject.math = authored.math
		if (authored.inline !== undefined) textObject.inline = authored.inline
		tmpTextObjects.push(textObject)
	} else if (Array.isArray(slideObj.text)) {
		// Handle cases 4,5,6
		// NOTE: use cast as text is TextProps[]|TableCellInternal[] and their `options` dont overlap (they share the same TextBaseProps though)
		// `math` carries raw OMML for native equation paragraphs — preserved here so STEP 5/6 can isolate it.
		tmpTextObjects = (slideObj.text as TextProps[]).map((item) => {
			// Projected key by key rather than spread: the array may really be `TableCellInternal[]` (see
			// the cast note above), and only these four belong on a `TextProps`. A key the item
			// does not have stays off the projection rather than arriving as an `undefined`.
			const projected: TextProps = {}
			if (item.text !== undefined) projected.text = item.text
			if (item.options !== undefined) projected.options = item.options
			if (item.math !== undefined) projected.math = item.math
			if (item.inline !== undefined) projected.inline = item.inline
			return projected
		})
	}

	// STEP 4: Iterate over text objects, set text/options, break into pieces if '\n'/breakLine found
	tmpTextObjects.forEach((itext, idx) => {
		if (!itext.text) itext.text = ''

		// A: Set options
		itext.options = itext.options || opts || {}
		if (idx === 0 && itext.options && !itext.options.bullet && opts.bullet) itext.options.bullet = opts.bullet

		// B: Cast to text-object and fix line-breaks (if needed)
		if (typeof itext.text === 'string' || typeof itext.text === 'number') {
			// 1: Convert "\n" or any variation into CRLF
			itext.text = itext.text.toString().replace(/\r*\n/g, CRLF)
		}

		// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
		// NOTE: Filter for trailing lineBreak prevents the creation of an empty textObj as the last item
		if (itext.text.includes(CRLF) && itext.text.match(/\n$/g) === null) {
			const lines = itext.text.split(CRLF)
			lines.forEach((line, lineIdx) => {
				const isLast = lineIdx === lines.length - 1
				// Non-last pieces need a paragraph break after them (the \n implies it).
				// The last piece inherits the caller's breakLine intent, *including* none — writing the
				// key as `undefined` would make an unstated intent look like a stated one to the next
				// spread. Copied rather than mutated, so the caller's own options are untouched.
				const lineOptions: TextPropsOptions = { ...itext.options }
				if (!isLast) lineOptions.breakLine = true
				arrTextObjects.push({ text: line, options: lineOptions })
			})
		} else {
			arrTextObjects.push({ ...itext, options: itext.options ?? {} })
		}
	})

	// STEP 5: Group textObj into lines by checking for lineBreak, bullets, alignment change, etc.
	const arrLines = groupRunsIntoLines(arrTextObjects, opts)

	// STEP 6: Loop over each line and create paragraph props, text run, etc.
	strSlideXml += renderTextParagraphsXml(arrLines, slideObj, opts)

	// IMPORTANT: An empty txBody will cause "needs repair" error! Add <p> content if missing.
	// This fixes an issue with table auto-paging where some cells would be empty on subsequent pages.
	/*
		<a:txBody>
			<a:bodyPr/>
			<a:lstStyle/>
		</a:txBody>
	*/
	// NOTE: this scans the accumulated CHILDREN, which no longer include the opening `<p:txBody>`.
	// That does not change the result: neither `<a:bodyPr>` nor `<a:lstStyle>` can contain the
	// exact substring `<a:p>` (`<a:pPr`/`<a:lvl1pPr` do not match), so only real paragraphs do.
	if (!strSlideXml.includes('<a:p>')) {
		strSlideXml += el('a:p', null, raw(voidEl('a:endParaRPr')))
	}

	// STEP 7: Close the textBody
	// LAST: Return XML
	return el(slideObj._type === SlideObjectType.tablecell ? 'a:txBody' : 'p:txBody', null, raw(strSlideXml))
}

/**
 * Generate an XML Placeholder
 * @param {SlideObject} placeholderObj
 * @returns XML
 */
export function genXmlPlaceholder(placeholderObj: SlideObject | null): string {
	if (!placeholderObj) return ''

	const placeholderIdx = placeholderObj.options?._placeholderIdx ? placeholderObj.options._placeholderIdx : ''
	const placeholderTyp = placeholderObj.options?._placeholderType ? placeholderObj.options._placeholderType : ''
	// Normalize to the OOXML ST_PlaceholderType value, accepting either a friendly PlaceholderType
	// key ('image', 'table') or the mapped value ('pic', 'tbl') - the latter is what `PLACEHOLDER_TYPE`
	// actually declares. Unknown strings emit no type rather than an invalid attribute.
	const placeholderType = PLACEHOLDER_TYPE_MAP[placeholderTyp]
		? PLACEHOLDER_TYPE_MAP[placeholderTyp].toString()
		: (Object.values(PlaceholderType) as string[]).includes(placeholderTyp)
			? placeholderTyp
			: ''

	// `hasCustomPrompt` flags a placeholder *definition* (layout/master) that carries custom
	// prompt text; it must not be set on a populated slide-level text shape promoted to a
	// placeholder, or PowerPoint would treat the visible text as prompt text.
	const isPlaceholderDef = placeholderObj._type === SlideObjectType.placeholder

	// NOTE: `placeholderType` is already the mapped OOXML value (e.g. 'pic', 'tbl') validated on
	// the line above; do NOT re-look it up in PLACEHOLDER_TYPE_MAP (its keys are the input names,
	// not the mapped values), or the type attribute is silently dropped for image/table placeholders.
	//
	// NOT built with the element builder, deliberately. This template's own source indentation
	// reaches the file: attributes are separated by a newline + two tabs, each carries an extra
	// leading space, and an ABSENT attribute still emits its separator — so the whitespace depends
	// on which attributes are missing. `voidEl` joins attributes with exactly one space and omits
	// absent ones entirely, so this layout is not expressible. 105 baseline parts carry `<p:ph`,
	// so tidying it is a visible byte change and belongs in its own fixture-gated commit.
	return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx.toString() + '"' : ''}
		${placeholderType ? ` type="${placeholderType}"` : ''}
		${isPlaceholderDef && placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}
		/>`
}
