/**
 * ts-pptx: text / placeholder slide-object serialization
 *
 * Emits a `text` or `placeholder` slide object as a `<p:sp>`: body insets from the caller's
 * margin, the shape geometry (preset or custom), fill, outline and shadow, then the text body.
 * An equation-bearing shape is wrapped in the `a14` markup-compatibility envelope.
 */

import { SlideObjectType } from '../../../enums.js'
import { DEF_TEXT_SHADOW } from '../../../constants-internal.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideObject } from '../../../types/internal.js'
import { createShadowEffectLst } from '../../drawingml/effect.js'
import { genXmlColorSelection } from '../../drawingml/fill.js'
import { genXmlCustGeom, genXmlPresetGeom } from '../../drawingml/geometry.js'
import { genXmlObjectLock, SHAPE_LOCK_ATTRS } from '../../drawingml/locks.js'
import { genXmlPlaceholder, genXmlTextBody, objectHasMath } from '../../drawingml/text-body.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { marginToEmu } from '../../../units-internal.js'
import { EMU_PER_INCH } from '../../../units.js'
import { cNvPrHyperlink, cNvPrOpen, genXmlShapeLine } from './shared.js'

/**
 * Render a `text` / `placeholder` slide object to its `<p:sp>` XML.
 */
export function renderTextObject(
	slideItemObj: SlideObject,
	idx: number,
	slide: PresSlideInternal | SlideLayoutInternal,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null,
	locationAttrs: XmlAttrs
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// Lines can have zero cy, but text should not
	if (!slideItemObj.options.line && cy === 0) cy = EMU_PER_INCH * 0.3

	// Margin/Padding/Inset for textboxes
	if (!slideItemObj.options._bodyProp) slideItemObj.options._bodyProp = {}
	if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
		// Margin arrays are documented as [Top, Right, Bottom, Left] (CSS order) and table cells /
		// slide numbers already map them that way. Keep textboxes consistent: index 0=Top, 3=Left.
		// Margins are inches (see `marginToEmu`), matching cell margins and the PowerPoint dialog.
		slideItemObj.options._bodyProp.tIns = marginToEmu(slideItemObj.options.margin[0] || 0)
		slideItemObj.options._bodyProp.rIns = marginToEmu(slideItemObj.options.margin[1] || 0)
		slideItemObj.options._bodyProp.bIns = marginToEmu(slideItemObj.options.margin[2] || 0)
		slideItemObj.options._bodyProp.lIns = marginToEmu(slideItemObj.options.margin[3] || 0)
	} else if (typeof slideItemObj.options.margin === 'number') {
		slideItemObj.options._bodyProp.lIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.rIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.bIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.tIns = marginToEmu(slideItemObj.options.margin)
	}

	// A: Start SHAPE =======================================================
	// A native equation uses the `a14` (drawing-2010) markup-compatibility extension.
	// PowerPoint wraps the whole shape in <mc:AlternateContent><mc:Choice Requires="a14"> so
	// non-a14 consumers (and schema validators) treat the a14:m subtree as a known extension.
	if (objectHasMath(slideItemObj)) {
		strSlideXml += '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
		strSlideXml += '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
	}
	strSlideXml += '<p:sp>'

	// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
	const txtOpts = slideItemObj.options
	strSlideXml +=
		'<p:nvSpPr>' +
		cNvPrOpen(idx + 2, txtOpts.objectName, txtOpts.altText || '') +
		'>' +
		cNvPrHyperlink(txtOpts.hyperlink) +
		'</p:cNvPr>'
	{
		const spLockXml = genXmlObjectLock('a:spLocks', SHAPE_LOCK_ATTRS, txtOpts.objectLock, txtOpts.objectName)
		// NOTE: paired only when there are locks to carry; otherwise self-closing. That is an arity
		// difference, so it cannot be expressed as one `el()` call.
		const cNvSpPrAttrs: XmlAttrs = { txBox: txtOpts?.isTextBox ? '1' : null }
		strSlideXml += spLockXml ? el('p:cNvSpPr', cNvSpPrAttrs, raw(spLockXml)) : voidEl('p:cNvSpPr', cNvSpPrAttrs)
	}
	// Prefer the resolved slide-layout placeholder; otherwise fall back to the shape's own
	// placeholder type so a standalone title/body text box still emits a real <p:ph>.
	strSlideXml += el(
		'p:nvPr',
		null,
		raw(
			genXmlPlaceholder(
				slideItemObj._type === SlideObjectType.placeholder ||
					(placeholderObj == null && slideItemObj.options?._placeholderType)
					? slideItemObj
					: placeholderObj
			)
		)
	)
	strSlideXml += '</p:nvSpPr><p:spPr>'
	strSlideXml += el('a:xfrm', locationAttrs, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])

	if (slideItemObj.shape === 'custGeom') {
		strSlideXml += genXmlCustGeom(slideItemObj.options, cx, cy, slide._presLayout)
	} else {
		strSlideXml += genXmlPresetGeom(slideItemObj.shape ?? '', slideItemObj.options, cx, cy)
	}

	// Option: FILL
	// A missing `fill` defaults to `<a:noFill/>` here — an unfilled box is what a text box
	// should be, and every deck ever authored against this writer depends on it. That default
	// is also why omission cannot mean *inherit* on this path: `fill: { type: 'inherit' }` is
	// the spelling that emits no fill child and lets `p:style/a:fillRef` or the placeholder
	// paint the interior.
	strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

	// shape Type: LINE: line color
	if (slideItemObj.options.line) strSlideXml += genXmlShapeLine(slideItemObj.options.line)

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (slideItemObj.options.shadow && slideItemObj.options.shadow.type !== 'none') {
		strSlideXml += createShadowEffectLst(slideItemObj.options.shadow, DEF_TEXT_SHADOW)
	}

	// B: Close shape Properties
	strSlideXml += '</p:spPr>'

	// C: Add formatted text (text body "bodyPr")
	strSlideXml += genXmlTextBody(slideItemObj)

	// LAST: Close SHAPE =======================================================
	strSlideXml += '</p:sp>'

	// Close the a14 markup-compatibility envelope for an equation-bearing shape.
	if (objectHasMath(slideItemObj)) strSlideXml += '</mc:Choice></mc:AlternateContent>'
	return strSlideXml
}
