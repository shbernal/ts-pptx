/**
 * ts-pptx: finding the layout placeholder a slide object names
 *
 * A slide object that names a placeholder (`addText('x', { placeholder: 'body' })`) takes options
 * from the layout's placeholder when it is defined, and its frame and `<p:ph>` from it when it is
 * written. Those lookups were written out separately, and the write-time one matched any layout
 * object carrying a `placeholder` option while the definition-time ones matched only placeholders.
 * So a slide object could take its options from one layout object and its frame and `<p:ph>` from
 * another, and bind a `<p:ph>` to a master's plain text box. This is the one lookup. It has no
 * dependencies of its own, so the render path can use it without linking the definers.
 */

import { SlideObjectType } from '../../enums.js'
import type { SlideLayoutInternal, SlideObject } from '../../types/internal.js'

/**
 * The layout placeholder named `name`: the first `placeholder` object on the layout whose
 * `placeholder` option is `name`.
 *
 * Only placeholder objects qualify. A layout text box can carry a `placeholder` option too, but it is
 * not a placeholder a slide can bind a `<p:ph>` to or inherit from.
 * @param layout - the slide's layout, if it has one
 * @param name - the placeholder name the slide object states, if any
 * @returns the layout placeholder, or `null` when there is none
 */
export function findLayoutPlaceholder(
	layout: SlideLayoutInternal | null | undefined,
	name: string | undefined
): SlideObject | null {
	if (!layout?._slideObjects || !name) return null
	return (
		layout._slideObjects.find(
			(obj) => obj._type === SlideObjectType.placeholder && obj.options?.placeholder === name
		) ?? null
	)
}
