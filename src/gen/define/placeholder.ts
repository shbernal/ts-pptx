/**
 * PptxGenJS: Placeholder Definition
 *
 * `addPlaceholdersToSlideLayouts` seeds a slide with any layout placeholders it has not already
 * populated, so every inherited placeholder is present as an (empty) text object.
 */
import { SlideObjectType } from '../../core-enums.js'
import type { PresSlideInternal } from '../../core-interfaces.js'
import { addTextDefinition } from './text.js'

/**
 * Adds placeholder objects to slide
 * @param {PresSlideInternal} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: PresSlideInternal): void {
	if (!slide._slideLayout) return // Add all placeholders on this Slide that dont already exist
	;(slide._slideLayout._slideObjects || []).forEach((slideLayoutObj) => {
		if (slideLayoutObj._type === SlideObjectType.placeholder) {
			const slideLayoutOptions = slideLayoutObj.options || {}
			// A: Search for this placeholder on Slide before we add
			// NOTE: Check to ensure a placeholder does not already exist on the Slide
			// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
			if (
				!slide._slideObjects.some(
					(slideObj) => slideObj.options && slideObj.options.placeholder === slideLayoutOptions.placeholder
				)
			) {
				addTextDefinition(slide, [{ text: '' }], slideLayoutOptions, true)
			}
		}
	})
}
