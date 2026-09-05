/**
 * ts-pptx: the shape construct family
 *
 * Preset geometry (`addShape`) and the four shape descriptors a slide master or a group is written
 * with. A shape is authored as a text object carrying a preset geometry, so the family has no
 * renderer of its own -- what it owns is the preset table and the geometry emitter under it.
 */

import { ShapeType } from '../enums.js'
import { addShapeDefinition } from '../gen/define/shape.js'
import type { ConstructFamily } from './shared.js'

export const shapeFamily: ConstructFamily = {
	name: 'shape',
	authors: {
		addShape(slide, shapeName, options) {
			// `shapeName` is a plain string preset name (e.g. `ShapeType.rect` === "rect").
			addShapeDefinition(slide, shapeName, options || {})
		},
	},
	children: {
		line(target, child) {
			addShapeDefinition(target, ShapeType.line, child)
		},
		rect(target, child) {
			addShapeDefinition(target, ShapeType.rect, child)
		},
		roundRect(target, child) {
			addShapeDefinition(target, ShapeType.roundRect, child)
		},
		shape(target, child) {
			addShapeDefinition(target, child.type, child.options || {})
		},
	},
}
