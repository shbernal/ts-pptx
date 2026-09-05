/**
 * ts-pptx: the table construct family
 *
 * Tables and their auto-paging, from rows a caller already holds. Reproducing a rendered HTML
 * `<table>` belongs to this family too, but only where there is a DOM to read it from, so it is
 * composed separately by the browser entry.
 */

import { SlideObjectType } from '../enums.js'
import { renderTableObject } from '../gen/slide/objects/table.js'
import { addTableDefinition } from '../gen/define/table.js'
import type { ConstructFamily } from './shared.js'

export const tableFamily = {
	name: 'table',
	authors: {
		addTable(slide, tableRows, options) {
			// Appended, not assigned. Two `addTable` calls on one slide used to leave the accessor
			// reporting only the second table's continuations; the first table's were in the deck and
			// invisible through the one API that names them.
			//
			// Appended by identity, though. A table pages onto the slides after this one whether it
			// created them or found them already there, so a second table on the same slide usually
			// lands on the first table's continuations -- the same slide, spilled onto twice. The
			// accessor names slides so a caller can address them; naming one twice is noise, not
			// information.
			const paged = addTableDefinition(
				slide,
				tableRows,
				options || {},
				slide._slideLayout,
				slide._presLayout,
				slide.addSlide,
				slide.getSlide
			)
			for (const made of paged) if (!slide._newAutoPagedSlides.includes(made)) slide._newAutoPagedSlides.push(made)
		},
	},
	renderers: {
		[SlideObjectType.table]: renderTableObject,
	},
} satisfies ConstructFamily
