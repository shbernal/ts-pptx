/**
 * PptxGenJS: slide part + slide/layout rels
 *
 * Emit a slide (`ppt/slides/slideN.xml`) and the relationship files for slides
 * and layouts (`slideN.xml.rels`, `slideLayoutN.xml.rels`).
 */

import { CRLF, XML_DECL } from '../../core-enums-internal.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { slideTimingToXml } from '../anim/timing.js'
import { slideTransitionToXml } from '../anim/transition.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { slideObjectRelationsToXml, slideObjectToXml } from './object.js'

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
				'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
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
	if (!slideLayout) throw new Error(`makeXmlSlideLayoutRel: no slide layout at index ${layoutNumber - 1}`)
	return slideObjectRelationsToXml(slideLayout, [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
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
	if (!slide) throw new Error(`makeXmlSlideRel: no slide at index ${slideNumber - 1}`)
	const defaultRels = [
		{
			target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		},
		{
			target: `../notesSlides/notesSlide${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
		},
	]
	// Only emit the comments rel for slides that actually carry comments (the comment part
	// is likewise only written for those slides); the rId is assigned after slideLayout/notesSlide.
	if ((slide._comments || []).length > 0) {
		defaultRels.push({
			target: `../comments/comment${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
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
