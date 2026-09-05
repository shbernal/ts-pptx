/**
 * ts-pptx: the zoom construct family
 *
 * Slide, Section and Summary Zoom (PowerPoint's Insert > Zoom) -- clickable tiles that zoom to a
 * slide or to the start of a section. The two section forms resolve their targets against the live
 * section list the slide carries.
 */

import { addSectionZoomDefinition, addSlideZoomDefinition, addSummaryZoomDefinition } from '../gen/define/zoom.js'
import type { ConstructFamily } from './shared.js'

export const zoomFamily: ConstructFamily = {
	name: 'zoom',
	authors: {
		addSlideZoom(slide, options) {
			addSlideZoomDefinition(slide, options)
		},
		addSectionZoom(slide, options) {
			addSectionZoomDefinition(slide, options, slide.getSections())
		},
		addSummaryZoom(slide, options) {
			addSummaryZoomDefinition(slide, options, slide.getSections())
		},
	},
}
