/**
 * PptxGenJS: DrawingML geometry
 *
 * Emit `<a:prstGeom>` (preset shapes with their adjustment guides) and
 * `<a:custGeom>` (freeform paths built from the `points` DSL). Shared by the
 * shape and image code paths so path/preset emission stays in one place.
 */

import { VALID_SHAPE_PRESETS } from '../../core-enums.js'
import type { Coord, GeometryPoint, ObjectOptions, PresLayout } from '../../core-interfaces.js'
import { convertArcAngle, convertRotationDegrees, getSmartParseNumber } from '../../units-internal.js'
import { EMU_PER_INCH, PERCENT_SCALE } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { warn } from '../../log.js'

/**
 * Several `<a:custGeom>` children are emitted as `<a:avLst />` — with a space before the
 * slash — where the rest of the tree writes `<a:avLst/>`. Byte-significant, so it stays.
 */
const SPACE_BEFORE_SLASH = { closePrefix: ' ' }

/**
 * The path nodes inside `<a:cubicBezTo>`/`<a:quadBezTo>` are newline-and-tab separated,
 * an artifact of the source indentation these were originally written with. It reaches
 * the emitted bytes, so it is described here rather than left to leak from a literal.
 */
const BEZ_INDENT = { childPrefix: '\n\t\t\t\t\t', closePrefix: '\n\t\t\t\t\t' }

/**
 * Emit an `<a:prstGeom>` for a preset shape, including any adjust values (`<a:avLst>`).
 * Shared by the shape and image code paths so that geometry + adjust handling stays in one place.
 * @param {string} shapeName - preset geometry name (e.g. `rect`, `ellipse`, `roundRect`, `hexagon`)
 * @param {ObjectOptions} options - object options carrying optional `rectRadius`/`angleRange`/`arcThicknessRatio`
 * @param {number} cx - shape width (EMU), used to scale `rectRadius`
 * @param {number} cy - shape height (EMU), used to scale `rectRadius`
 * @return {string} `<a:prstGeom>` XML
 */
// Shapes whose corner-radius adjust value is named adj1 (+ adj2) instead of adj.
// Sourced from ECMA-376 Annex D electronic addenda (presetShapeDefinitions.xml).
export const RECT_RADIUS_ADJ1_SHAPES = new Set(['round2SameRect', 'round2DiagRect'])

export function genXmlPresetGeom(shapeName: string, options: ObjectOptions, cx: number, cy: number): string {
	// Safety net for every prstGeom emitter (addShape, addText/addImage `shape`):
	// an unknown preset becomes an invalid `prst` value that makes PowerPoint show
	// the "needs repair" dialog and drop the shape. Fail loudly instead.
	if (!VALID_SHAPE_PRESETS.has(shapeName)) {
		throw new Error(
			`Invalid shape "${String(shapeName)}"! Use a value from \`pptxgen.ShapeType.*\` (e.g. \`pptxgen.ShapeType.rect\`). PowerPoint can't render unknown preset geometries and will drop the shape during repair.`
		)
	}
	// Collect adjustment guides; track names so the generic `shapeAdjust` passthrough
	// never emits a duplicate `<a:gd>` for a handle a friendly shortcut already set.
	let avLst = ''
	const emittedAdjNames = new Set<string>()
	const emitGuide = (name: string, fmlaVal: number): void => {
		// `voidEl` escapes the guide name, which the `shapeAdjust` passthrough takes from the
		// caller. It is validated as a non-empty string but not charset-checked, so escaping is
		// a real (if narrow) hardening; every in-tree name is a plain `adj*` so bytes don't move.
		avLst += voidEl('a:gd', { name, fmla: `val ${fmlaVal}` })
		emittedAdjNames.add(name)
	}
	if (options.rectRadius) {
		const adjVal = Math.round((options.rectRadius * EMU_PER_INCH * PERCENT_SCALE) / Math.min(cx, cy))
		if (RECT_RADIUS_ADJ1_SHAPES.has(shapeName)) {
			emitGuide('adj1', adjVal)
			emitGuide('adj2', 0)
		} else {
			emitGuide('adj', adjVal)
		}
	} else if (options.angleRange) {
		for (let i = 0; i < 2; i++) {
			const angle = options.angleRange[i] ?? 0
			emitGuide(`adj${i + 1}`, convertRotationDegrees(angle))
		}

		if (options.arcThicknessRatio) {
			emitGuide('adj3', Math.round(options.arcThicknessRatio * (PERCENT_SCALE / 2)))
		}
	}
	// Generic adjustment handles (`shapeAdjust`) for any preset shape.
	if (options.shapeAdjust) {
		const adjusts = Array.isArray(options.shapeAdjust) ? options.shapeAdjust : [options.shapeAdjust]
		adjusts.forEach((adj) => {
			// Silent coercion of a bad guide produces a shape PowerPoint silently drops or repairs,
			// so warn and skip instead of emitting a degenerate `<a:gd>`.
			if (
				!adj ||
				typeof adj.name !== 'string' ||
				adj.name.length === 0 ||
				typeof adj.value !== 'number' ||
				!isFinite(adj.value)
			) {
				warn(
					`shapeAdjust entry ${JSON.stringify(adj)} is invalid (needs { name:string, value:number }) and was ignored.`
				)
				return
			}
			if (emittedAdjNames.has(adj.name)) {
				warn(`shapeAdjust "${adj.name}" was ignored because rectRadius/angleRange already set that handle.`)
				return
			}
			// `value` is a 0.0-1.0 fraction of the handle range, emitted as a percentage guide (1/100000 units).
			emitGuide(adj.name, Math.round(adj.value * PERCENT_SCALE))
		})
	}
	return el('a:prstGeom', { prst: shapeName }, raw(el('a:avLst', null, raw(avLst))))
}

/**
 * Narrow a freeform path node to an arc segment. An arc is the one curve node with no end
 * point, so it must be split off before the remaining curve nodes can be read for x/y.
 */
function isArcPoint(point: GeometryPoint): point is Extract<GeometryPoint, { curve: { type: 'arc' } }> {
	return 'curve' in point && point.curve.type === 'arc'
}

/**
 * Emit an `<a:custGeom>` for a freeform path built from `points`.
 * Shared by the shape and image code paths so that path emission stays in one place.
 * Points are authored in the object's own inch/EMU space (0..cx, 0..cy) — not slide-relative and not normalized.
 * @param {ObjectOptions['points']} points - freeform path DSL (`moveTo`/`lnTo`/`cubicBezTo`/`quadBezTo`/`arcTo`/`close`)
 * @param {number} cx - object width (EMU), used as the path viewport width
 * @param {number} cy - object height (EMU), used as the path viewport height
 * @param {PresLayout} layout - presentation layout used to resolve point coordinates to EMU
 * @return {string} `<a:custGeom>` XML
 */
export function genXmlCustGeom(points: ObjectOptions['points'], cx: number, cy: number, layout: PresLayout): string {
	/** `<a:pt>` in the object's own EMU space. */
	const pt = (x: Coord | undefined, y: Coord | undefined): string =>
		voidEl(
			'a:pt',
			{ x: getSmartParseNumber(x, 'X', layout), y: getSmartParseNumber(y, 'Y', layout) },
			SPACE_BEFORE_SLASH
		)

	const nodes: string[] = []
	points?.forEach((point, i) => {
		if (isArcPoint(point)) {
			// An `<a:arcTo>` has no end point: it is derived from the pen position, radii and sweep.
			// An authored x/y is silently unused, so say so rather than let it read as meaningful.
			// (A union excess-property check does not reject one, so this is the only signal.)
			if ('x' in point || 'y' in point)
				warn('freeform arc node: x/y are ignored — an arcTo end point is computed from stAng/swAng and the radii.')
			nodes.push(
				voidEl(
					'a:arcTo',
					{
						hR: getSmartParseNumber(point.curve.hR, 'Y', layout),
						wR: getSmartParseNumber(point.curve.wR, 'X', layout),
						stAng: convertArcAngle(point.curve.stAng, 'stAng'),
						swAng: convertArcAngle(point.curve.swAng, 'swAng'),
					},
					SPACE_BEFORE_SLASH
				)
			)
		} else if ('curve' in point) {
			switch (point.curve.type) {
				case 'cubic':
					nodes.push(
						el(
							'a:cubicBezTo',
							null,
							[pt(point.curve.x1, point.curve.y1), pt(point.curve.x2, point.curve.y2), pt(point.x, point.y)].map(raw),
							BEZ_INDENT
						)
					)
					break
				case 'quadratic':
					nodes.push(
						el('a:quadBezTo', null, [pt(point.curve.x1, point.curve.y1), pt(point.x, point.y)].map(raw), BEZ_INDENT)
					)
					break
				default:
					break
			}
		} else if ('close' in point) {
			nodes.push(voidEl('a:close', null, SPACE_BEFORE_SLASH))
		} else if (point.moveTo || i === 0) {
			nodes.push(el('a:moveTo', null, raw(pt(point.x, point.y))))
		} else {
			nodes.push(el('a:lnTo', null, raw(pt(point.x, point.y))))
		}
	})

	// custGeom preamble — the sub-lists OOXML requires before `<a:pathLst>`: adjust values
	// (avLst), guide formulas (gdLst), adjust handles (ahLst), connection sites (cxnLst), and the
	// text rectangle (rect). PptxGenJS drives geometry entirely from the path, so all stay empty.
	return el(
		'a:custGeom',
		null,
		[
			voidEl('a:avLst', null, SPACE_BEFORE_SLASH),
			el('a:gdLst'),
			voidEl('a:ahLst', null, SPACE_BEFORE_SLASH),
			el('a:cxnLst'),
			voidEl('a:rect', { l: 'l', t: 't', r: 'r', b: 'b' }, SPACE_BEFORE_SLASH),
			el('a:pathLst', null, raw(el('a:path', { w: cx, h: cy }, nodes.map(raw)))),
		].map(raw)
	)
}
