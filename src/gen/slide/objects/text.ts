/**
 * ts-pptx: text / placeholder slide-object serialization
 *
 * Emits a `text` or `placeholder` slide object as a `<p:sp>`: body insets from the caller's
 * margin, the shape geometry (preset or custom), fill, outline and shadow, then the text body.
 * An equation-bearing shape is wrapped in the `a14` markup-compatibility envelope.
 */

import { SlideObjectType } from '../../../enums.js'
import { DEF_TEXT_SHADOW } from '../../../constants-internal.js'
import { createShadowEffectLst } from '../../drawingml/effect.js'
import { genXmlColorSelection } from '../../drawingml/fill.js'
import { genXmlCustGeom, genXmlPresetGeom } from '../../drawingml/geometry.js'
import { genXmlObjectLock, SHAPE_LOCK_ATTRS } from '../../drawingml/locks.js'
import { genXmlPlaceholder, genXmlTextBody, objectHasMath } from '../../drawingml/text-body.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { OOXML_NS } from '../../../ooxml/namespaces.js'
import { cNvPrHyperlink, cNvPrOpen, genXmlShapeLine, type RenderContext, xfrmEl } from './shared.js'
import { xsdBoolIfTrue } from '../../../ooxml/xsd-boolean.js'

/**
 * Render a `text` / `placeholder` slide object to its `<p:sp>` XML.
 */
export function renderTextObject(ctx: RenderContext): string {
	const {
		obj: slideItemObj,
		shapeId,
		slide,
		frame: { x, y, cx, cy },
		placeholder: placeholderObj,
		locationAttrs,
		itemOpts,
	} = ctx
	let strSlideXml = ''
	// `itemOpts` is the caller's already-normalized `itemOpts` (see the dispatch in
	// `slideObjectToXml`). Read it rather than re-narrowing the field: this function has exactly
	// one call site, and a contract stated there beats a defensive re-assignment here.
	//
	// The zero-height rescue that used to sit here is gone: it read `!itemOpts.line`, and both
	// definers write a `line` object onto every text object unconditionally, so it had never
	// once fired. The default it was reaching for now lives in `addTextDefinition`, where an
	// omitted height is still distinguishable from a stated `h: 0`.

	// A: Start SHAPE =======================================================
	strSlideXml += '<p:sp>'

	// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
	const txtOpts = itemOpts
	strSlideXml +=
		'<p:nvSpPr>' +
		cNvPrOpen(shapeId, txtOpts.objectName, txtOpts.altText || '') +
		'>' +
		cNvPrHyperlink(txtOpts.hyperlink) +
		'</p:cNvPr>'
	{
		const spLockXml = genXmlObjectLock('a:spLocks', SHAPE_LOCK_ATTRS, txtOpts.objectLock, txtOpts.objectName)
		// NOTE: paired only when there are locks to carry; otherwise self-closing. That is an arity
		// difference, so it cannot be expressed as one `el()` call.
		const cNvSpPrAttrs: XmlAttrs = { txBox: xsdBoolIfTrue(txtOpts?.isTextBox) }
		strSlideXml += spLockXml ? el('p:cNvSpPr', cNvSpPrAttrs, raw(spLockXml)) : voidEl('p:cNvSpPr', cNvSpPrAttrs)
	}
	// Prefer the resolved slide-layout placeholder; otherwise fall back to the shape's own
	// placeholder type so a standalone title/body text box still emits a real <p:ph>.
	strSlideXml += el(
		'p:nvPr',
		null,
		raw(
			genXmlPlaceholder(
				slideItemObj._type === SlideObjectType.placeholder || (placeholderObj == null && itemOpts?._placeholderType)
					? slideItemObj
					: placeholderObj
			)
		)
	)
	strSlideXml += '</p:nvSpPr><p:spPr>'
	strSlideXml += xfrmEl('a:xfrm', { x, y, cx, cy }, locationAttrs)

	if (slideItemObj.shape === 'custGeom') {
		strSlideXml += genXmlCustGeom(itemOpts, cx, cy, slide._presLayout)
	} else {
		strSlideXml += genXmlPresetGeom(slideItemObj.shape ?? '', itemOpts, cx, cy)
	}

	// Option: FILL
	// A missing `fill` defaults to `<a:noFill/>` here — an unfilled box is what a text box
	// should be, and every deck ever authored against this writer depends on it. That default
	// is also why omission cannot mean *inherit* on this path: `fill: { type: 'inherit' }` is
	// the spelling that emits no fill child and lets `p:style/a:fillRef` or the placeholder
	// paint the interior.
	strSlideXml += itemOpts.fill ? genXmlColorSelection(itemOpts.fill) : '<a:noFill/>'

	// shape Type: LINE: line color
	if (itemOpts.line) strSlideXml += genXmlShapeLine(itemOpts.line)

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (itemOpts.shadow && itemOpts.shadow.type !== 'none') {
		strSlideXml += createShadowEffectLst(itemOpts.shadow, DEF_TEXT_SHADOW)
	}

	// B: Close shape Properties
	strSlideXml += '</p:spPr>'

	// C: Add formatted text (text body "bodyPr")
	strSlideXml += genXmlTextBody(slideItemObj)

	// LAST: Close SHAPE =======================================================
	strSlideXml += '</p:sp>'

	// A native equation uses the `a14` (drawing-2010) markup-compatibility extension. PowerPoint
	// wraps the whole shape in <mc:AlternateContent><mc:Choice Requires="a14"> so non-a14
	// consumers (and schema validators) treat the a14:m subtree as a known extension.
	if (objectHasMath(slideItemObj)) {
		return el(
			'mc:AlternateContent',
			{ 'xmlns:mc': OOXML_NS.mc },
			raw(el('mc:Choice', { 'xmlns:a14': OOXML_NS.a14, Requires: 'a14' }, raw(strSlideXml)))
		)
	}
	return strSlideXml
}
