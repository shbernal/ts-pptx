/**
 * ts-pptx: slide part + slide/layout rels
 *
 * Emit a slide (`ppt/slides/slideN.xml`) and the relationship files for slides
 * and layouts (`slideN.xml.rels`, `slideLayoutN.xml.rels`).
 */

import { CRLF, XML_DECL } from '../../constants-internal.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { slideTimingToXml } from '../anim/timing.js'
import { slideTransitionToXml } from '../anim/transition.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { slideObjectRelationsToXml, slideObjectToXml } from './object.js'
import { InternalError } from '../../errors.js'
import { PML_ROOT_NS } from '../../ooxml/namespaces.js'
import { NOTES_SLIDE_REL, OFFICE_REL, SLIDE_LAYOUT_REL, SLIDE_MASTER_REL } from '../../ooxml/rel-types.js'

/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlideInternal} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlSlide(slide: PresSlideInternal): string {
	return (
		XML_DECL +
		CRLF +
		el(
			'p:sld',
			{
				...PML_ROOT_NS,
				show: slide?.hidden ? '0' : null,
			},
			[
				raw(slideObjectToXml(slide)),
				raw(el('p:clrMapOvr', null, raw(voidEl('a:masterClrMapping')))),
				raw(slideTransitionToXml(slide)),
				raw(slideTimingToXml(slide)),
			]
		)
	)
}

/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: SlideLayoutInternal[]): string {
	const slideLayout = slideLayouts[layoutNumber - 1]
	if (!slideLayout)
		throw new InternalError(
			'slide/rel-index-out-of-range',
			`makeXmlSlideLayoutRel: no slide layout at index ${layoutNumber - 1}`
		)
	return slideObjectRelationsToXml(slideLayout, [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: SLIDE_MASTER_REL,
		},
	])
}

/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlideInternal[]} slides
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlSlideRel(
	slides: PresSlideInternal[],
	slideLayouts: SlideLayoutInternal[],
	slideNumber: number
): string {
	const slide = slides[slideNumber - 1]
	if (!slide)
		throw new InternalError('slide/rel-index-out-of-range', `makeXmlSlideRel: no slide at index ${slideNumber - 1}`)
	const defaultRels = [
		{
			target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
			type: SLIDE_LAYOUT_REL,
		},
		{
			target: `../notesSlides/notesSlide${slideNumber}.xml`,
			type: NOTES_SLIDE_REL,
		},
	]
	// Only emit the comments rel for slides that actually carry comments (the comment part
	// is likewise only written for those slides); the rId is assigned after slideLayout/notesSlide.
	if ((slide._comments || []).length > 0) {
		defaultRels.push({
			target: `../comments/comment${slideNumber}.xml`,
			// Only this module emits a comments rel, so it is built here rather than hoisted.
			type: OFFICE_REL + 'comments',
		})
	}
	return slideObjectRelationsToXml(slide, defaultRels)
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {PresSlideInternal[]} slides - srray of slides
 * @param {SlideLayoutInternal[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide(
	slides: PresSlideInternal[],
	slideLayouts: SlideLayoutInternal[],
	slideNumber: number
): number {
	for (let i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i]?._name === slides[slideNumber - 1]?._slideLayout?._name) {
			return i + 1
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1
}
