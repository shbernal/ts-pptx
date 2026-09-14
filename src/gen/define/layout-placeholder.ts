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
import type { AuthoredFrame } from './frame.js'

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

/**
 * The frame the layout placeholder named `name` states, for an object that fills the placeholder's
 * geometry to take any axis it leaves unstated, as `given[axis] ?? placeholder[axis]`.
 *
 * An image and a table each copied the four axes themselves, under different rules: the image
 * treated a `null` axis as unstated and the table did not, so `x: null` put an image at the
 * placeholder's left edge and a table at the default half inch.
 * @param layout - the slide's layout, if it has one
 * @param name - the placeholder name the slide object states, if any
 * @returns the placeholder's axes, each absent when it states none or there is no such placeholder
 */
export function placeholderFrame(
	layout: SlideLayoutInternal | null | undefined,
	name: string | undefined
): AuthoredFrame {
	const options = findLayoutPlaceholder(layout, name)?.options
	return options ? { x: options.x, y: options.y, w: options.w, h: options.h } : {}
}
