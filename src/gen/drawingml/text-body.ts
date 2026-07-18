/**
 * PptxGenJS: DrawingML text body
 *
 * The container layer over `text-run.ts`: body properties (`<a:bodyPr>`), the
 * full `<p:txBody>`/`<a:txBody>` builder (`genXmlTextBody`), the native-equation
 * predicate, and the placeholder `<p:ph>` element.
 */

import { CRLF, PlaceholderType, SlideObjectType } from '../../core-enums.js'
import type { ObjectOptions, SlideObject, TableCell, TextProps } from '../../core-interfaces.js'
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
 * @param {SlideObject | TableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject: SlideObject | TableCell): string {
	let bodyProperties = '<a:bodyPr'

	// Placeholders (incl. master/layout placeholders) carry their margin/valign in `_bodyProp` just
	// like text boxes, so they must emit the same configured `<a:bodyPr>` — otherwise a placeholder
	// authored with insets or a vertical anchor silently degrades to the default.
	// `_bodyProp`/`options` are optional on the type but present on text/placeholder objects that reach
	// this branch; bind them once so the body reads a narrowed, non-undefined value.
	const options = (slideObject as SlideObject).options
	const bodyProp = options?._bodyProp
	if (
		slideObject &&
		(slideObject._type === SlideObjectType.text || slideObject._type === SlideObjectType.placeholder) &&
		bodyProp
	) {
		// PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>

		// A: Enable or disable textwrapping none or square
		bodyProperties += bodyProp.wrap ? ' wrap="square"' : ' wrap="none"'

		// B: Textbox margins [padding]
		if (bodyProp.lIns || bodyProp.lIns === 0) bodyProperties += ` lIns="${bodyProp.lIns}"`
		if (bodyProp.tIns || bodyProp.tIns === 0) bodyProperties += ` tIns="${bodyProp.tIns}"`
		if (bodyProp.rIns || bodyProp.rIns === 0) bodyProperties += ` rIns="${bodyProp.rIns}"`
		if (bodyProp.bIns || bodyProp.bIns === 0) bodyProperties += ` bIns="${bodyProp.bIns}"`

		// C.1: Text columns (numCol/spcCol). Spacing is only meaningful when there is more than one column.
		if (bodyProp.numCol) bodyProperties += ` numCol="${bodyProp.numCol}"`
		if (bodyProp.spcCol) bodyProperties += ` spcCol="${bodyProp.spcCol}"`

		// C: Add rtl after margins
		bodyProperties += ' rtlCol="0"'

		// D: Add anchorPoints
		if (bodyProp.anchor) bodyProperties += ' anchor="' + bodyProp.anchor + '"' // VALS: [t,ctr,b]
		if (bodyProp.vert) bodyProperties += ' vert="' + bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// E: Close <a:bodyPr element
		bodyProperties += '>'

		// E.1: Preset text warp (`<a:prstTxWarp>`). Per CT_TextBodyProperties this child
		// comes before the autofit group, so emit it immediately after the attributes.
		if (bodyProp.prstTxWarp) {
			bodyProperties += `<a:prstTxWarp prst="${bodyProp.prstTxWarp}"><a:avLst/></a:prstTxWarp>`
		}

		/**
		 * F: Text Fit/AutoFit/Shrink option
		 * @see: http://officeopenxml.com/drwSp-text-bodyPr-fit.php
		 * @see: http://www.datypic.com/sc/ooxml/g-a_EG_TextAutofit.html
		 */
		if (options?.fit) {
			const fit = options.fit
			// NOTE: Use of '<a:noAutofit/>' instead of '' causes issues in PPT-2013!
			if (fit === 'none') bodyProperties += ''
			// NOTE: Bare shrink does not work automatically - PowerPoint calculates fontScale/lnSpcReduction dynamically upon edit/resize.
			// The object form bakes explicit values into the file (MS-PPT > Format shape > Text Options: "Shrink text on overflow").
			else if (fit === 'shrink') bodyProperties += '<a:normAutofit/>'
			else if (fit === 'resize') bodyProperties += '<a:spAutoFit/>'
			else if (typeof fit === 'object' && fit.type === 'shrink') bodyProperties += genXmlNormAutofit(fit)
		}

		// LAST: Close _bodyProp
		bodyProperties += '</a:bodyPr>'
	} else {
		// DEFAULT:
		bodyProperties += ' wrap="square" rtlCol="0">'
		bodyProperties += '</a:bodyPr>'
	}

	// LAST: Return Close _bodyProp
	return slideObject._type === SlideObjectType.tablecell ? '<a:bodyPr/>' : bodyProperties
}

/** Whether a slide object carries a native equation (`math` raw OMML) on any of its text items. */
export function objectHasMath(slideObj: SlideObject): boolean {
	const text = slideObj.text as TextProps | TextProps[] | string | number | undefined
	if (Array.isArray(text)) return text.some((item) => item && typeof item === 'object' && !!item.math)
	if (text && typeof text === 'object') return !!text.math
	return false
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {SlideObject|TableCell} slideObj - slideObj or tableCell
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
export function genXmlTextBody(slideObj: SlideObject | TableCell): string {
	const opts: ObjectOptions = slideObj.options || {}
	let tmpTextObjects: TextProps[] = []
	const arrTextObjects: RunProps[] = []

	// FIRST: Shapes without text reach this point with `slideObj.text` null/undefined.
	// We MUST still emit a `<p:txBody>` with at least an empty `<a:p>` paragraph;
	// the empty-txBody fallback below appends `<a:p><a:endParaRPr/></a:p>` when no
	// `<a:p>` was produced. Returning early here would emit `<p:sp>` without
	// `<p:txBody>`, which PowerPoint reports as a needs-repair error.

	// STEP 1: Start textBody
	let strSlideXml = slideObj._type === SlideObjectType.tablecell ? '<a:txBody>' : '<p:txBody>'

	// STEP 2: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj)

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// KNOWN LIMITATION: horizontal align on a LINE does not work — text is always left-aligned inside the line.
		if (opts.h === 0 && opts.line && opts.align) strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		else if (slideObj._type === SlideObjectType.placeholder)
			strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(slideObj, true)}</a:lstStyle>`
		else strSlideXml += '<a:lstStyle/>'
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
		tmpTextObjects.push({
			text: slideObj.text || '',
			options: slideObj.options || {},
			math: (slideObj.text as TextProps).math,
			inline: (slideObj.text as TextProps).inline,
		})
	} else if (Array.isArray(slideObj.text)) {
		// Handle cases 4,5,6
		// NOTE: use cast as text is TextProps[]|TableCell[] and their `options` dont overlap (they share the same TextBaseProps though)
		// `math` carries raw OMML for native equation paragraphs — preserved here so STEP 5/6 can isolate it.
		tmpTextObjects = (slideObj.text as TextProps[]).map((item) => ({
			text: item.text,
			options: item.options,
			math: item.math,
			inline: item.inline,
		}))
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
				// The last piece inherits the caller's breakLine intent — do not mutate the original options object.
				arrTextObjects.push({
					text: line,
					options: { ...itext.options, breakLine: isLast ? itext.options?.breakLine : true },
				})
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
	if (!strSlideXml.includes('<a:p>')) {
		strSlideXml += '<a:p><a:endParaRPr/></a:p>'
	}

	// STEP 7: Close the textBody
	strSlideXml += slideObj._type === SlideObjectType.tablecell ? '</a:txBody>' : '</p:txBody>'

	// LAST: Return XML
	return strSlideXml
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
	return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx.toString() + '"' : ''}
		${placeholderType ? ` type="${placeholderType}"` : ''}
		${isPlaceholderDef && placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}
		/>`
}
