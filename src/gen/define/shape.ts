/**
 * ts-pptx: Shape Definition
 *
 * `addShapeDefinition` normalizes an `addShape()` preset (mapping friendly aliases, rejecting
 * presets PowerPoint can't parse), applies line defaults, registers hyperlink + image-fill rels,
 * and pushes a `text`-type shape object.
 */
import { type SHAPE_NAME, ShapeType, SlideObjectType } from '../../enums.js'
import { DEF_SHAPE_LINE_COLOR } from '../../constants-internal.js'
import type { ShapeLineProps, ShapeProps } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, validateObjectName } from '../utils.js'
import { correctShadowOptions } from '../drawingml/effect.js'
import { resolveFillKind } from '../drawingml/fill.js'
import { assertKnownPreset } from '../drawingml/geometry.js'
import { nextObjectNameIdx } from './object-name.js'
import { createHyperlinkRels } from './hyperlinks.js'
import { registerImageFillMedia } from './image.js'
import { InvalidOptionError } from '../../errors.js'

/**
 * Map of common friendly shape names users pass as bare strings to their
 * valid OOXML preset values. PowerPoint can't parse the friendly spellings
 * and removes the shape during repair .
 */
const SHAPE_NAME_ALIASES: { [key: string]: SHAPE_NAME } = {
	oval: 'ellipse',
	rectangle: 'rect',
	roundedRectangle: 'roundRect',
	roundedrectangle: 'roundRect',
}

// ===== Shapes & connectors =====

/**
 * Adds a shape object to a slide definition.
 * @param {PresSlideInternal} target slide object that the shape should be added to
 * @param {SHAPE_NAME} shapeName shape name
 * @param {ShapeProps} opts shape options
 */
export function addShapeDefinition(target: PresSlideInternal, shapeName: SHAPE_NAME, opts: ShapeProps): void {
	// Take ownership of the options before touching them, the same way `addTextDefinition` does.
	// Everything below writes normalization back onto whatever object it is handed — the `line`
	// defaults and the assigned `objectName` above all — so without the copy a style literal reused
	// across shapes carries one shape's settings to the next:
	//
	//   slide.addShape('rect', STYLE)            // STYLE now holds `objectName: 'Shape 0'`
	//   slide.addShape('ellipse', { ...STYLE })  // …spread onto the second, and the third
	//
	// which emits three shapes named `Shape 0` and a duplicate-`objectName` warning.
	//
	// Nested caller objects (`fill`, `shadow`) stay shared by reference, as in `addTextDefinition`:
	// the image fill's rel id is registered through that reference and read back at emit time.
	// `shadow` follows text rather than `copyChartOptions`, which does copy it. `correctShadowOptions`
	// normalizes in place, so the caller's shadow object comes back with a rounded `angle` and the
	// derived `_alpha` on it — but those are idempotent normalizations of the caller's own values,
	// not one shape's identity carried to the next, and `test/regression/shape/shared-shadow.test.js`
	// pins that `_alpha` precisely because a shadow object shared across shapes has to keep emitting
	// the same `<a:effectLst>`.
	const options: ShapeProps = typeof opts === 'object' ? { ...opts } : {}
	options.line = options.line || { type: 'none' }
	options.shadow = correctShadowOptions(options.shadow)
	// Normalize friendly shape names (e.g. "oval" -> "ellipse") to their valid
	// OOXML preset spellings before storing on the slide object.
	const resolvedShapeName: SHAPE_NAME =
		typeof shapeName === 'string' && SHAPE_NAME_ALIASES[shapeName] ? SHAPE_NAME_ALIASES[shapeName] : shapeName
	const newObject: SlideObject = {
		_type: SlideObjectType.text,
		shape: resolvedShapeName || ShapeType.rect,
		options,
	}

	// Reality check
	if (!shapeName)
		throw new InvalidOptionError(
			'shape/missing-type',
			'Missing/Invalid shape parameter! Example: `addShape(ShapeType.line, {x:1, y:1, w:1, h:1});`'
		)

	// Reject presets PowerPoint can't parse (a typo, or an unmapped friendly name)
	// here at the API boundary as well as at serialization time, so the message
	// names the call the caller actually made.
	assertKnownPreset(resolvedShapeName)

	// 1: ShapeLineProps defaults
	// A stroke can carry a non-solid paint just like a fill, so the kind comes from
	// `resolveFillKind` and is stamped on before the emitter sees it. Only a solid stroke
	// gets the default line color; every other kind takes its paint from its sub-object.
	// This block used to infer `gradient` alone, which is how `line: { pattern }` came out
	// a default-black solid: normalization had already written `type: 'solid'` over it.
	//
	// Spread first, then override only the keys this block actually defaults. Listing the
	// carried keys instead is what broke `pattern`, `image` and `cap`: each was added to
	// `ShapeLineProps` without being added here, so the emitter read a key normalization
	// had already dropped — `line: { type: 'pattern' }` reached `genXmlPatternFill` with no
	// pattern object and threw. A spread cannot fall out of sync with the type.
	const lineType = resolveFillKind(options.line)
	const newLineOpts: ShapeLineProps = {
		...options.line,
		type: lineType,
		color: lineType === 'solid' ? options.line.color || DEF_SHAPE_LINE_COLOR : options.line.color,
		transparency: options.line.transparency || 0,
		width: options.line.width || 1,
		dashType: options.line.dashType || 'solid',
	}
	if (typeof options.line === 'object' && options.line.type !== 'none') options.line = newLineOpts

	// 2: Set options defaults
	options.x = options.x || (options.x === 0 ? 0 : 1)
	options.y = options.y || (options.y === 0 ? 0 : 1)
	options.w = options.w || (options.w === 0 ? 0 : 1)
	options.h = options.h || (options.h === 0 ? 0 : 1)
	// Shapes are `_type === text` objects, so they share the text-box name bucket (`Shape 0`,
	// `Text 1`, …) — which is what stops a shape and a text box colliding on `0`.
	const shapeNameIdx = nextObjectNameIdx(target, SlideObjectType.text)
	options.objectName = options.objectName
		? encodeXmlAttrValue(validateObjectName(options.objectName, 'shape'))
		: `Shape ${shapeNameIdx}`

	// 3: Create hyperlink rels
	createHyperlinkRels(target, newObject)

	// 5: Register an image fill (if any) as a media relationship for serialize-time blipFill
	if (typeof options.fill === 'object' && resolveFillKind(options.fill) === 'image') {
		registerImageFillMedia(target, options.fill)
	}

	// LAST: Add object to slide
	target._slideObjects.push(newObject)
}
