/**
 * ts-pptx: DrawingML geometry
 *
 * Emit `<a:prstGeom>` (preset shapes with their adjustment guides) and
 * `<a:custGeom>` (freeform paths built from the `points` DSL). Shared by the
 * shape and image code paths so path/preset emission stays in one place.
 */

import { VALID_SHAPE_PRESETS } from '../../enums.js'
import type { Coord, GeometryPoint, ObjectOptions, PresLayout, ShapeAdjustHandleXY } from '../../types/index.js'
import { convertAngleUnits, convertArcAngle, getSmartParseNumber } from '../../units-internal.js'
import { EMU_PER_INCH, PERCENT_SCALE } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { warn } from '../../diagnostics.js'
import { InvalidOptionError } from '../../errors.js'

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
 * The 17 guide-formula operations ECMA-376 Part 1 §20.1.9.11 (`a:gd/@fmla`) defines.
 * The set is closed — the spec enumerates every operation with its arity — so an
 * unrecognized leading token is always a caller mistake, not a newer dialect.
 *
 * Only the operation is checked. The operands are deliberately left uninterpreted:
 * they may be literals, adjust names, or other guide names resolved in declaration
 * order, and validating them would mean re-implementing the geometry engine — which
 * is exactly what this passthrough exists to avoid.
 */
const GEOM_GUIDE_OPS = new Set([
	'*/',
	'+-',
	'+/',
	'?:',
	'abs',
	'at2',
	'cat2',
	'cos',
	'max',
	'min',
	'mod',
	'pin',
	'sat2',
	'sin',
	'sqrt',
	'tan',
	'val',
])

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
const RECT_RADIUS_ADJ1_SHAPES = new Set(['round2SameRect', 'round2DiagRect'])

/**
 * Reject a preset PowerPoint can't parse.
 *
 * An unknown `prst` value corrupts the package: PowerPoint shows the "needs
 * repair" dialog and drops the shape. So this is the safety net for every
 * `prstGeom` emitter — `addShape`, and the `shape` option on `addText`/`addImage`
 * — and it is also called during `addShape` normalization, so the caller sees
 * their own typo at the API boundary rather than at serialization time.
 * @param {string} shapeName - the resolved preset geometry name
 * @throws {InvalidOptionError} `shape/unknown-preset` when the name is not a preset
 */
export function assertKnownPreset(shapeName: string): void {
	if (VALID_SHAPE_PRESETS.has(shapeName)) return
	throw new InvalidOptionError(
		'shape/unknown-preset',
		`Invalid shape "${String(shapeName)}"! Use a value from \`ShapeType.*\` (e.g. \`ShapeType.rect\`). PowerPoint can't render unknown preset geometries and will drop the shape during repair.`
	)
}

export function genXmlPresetGeom(shapeName: string, options: ObjectOptions, cx: number, cy: number): string {
	assertKnownPreset(shapeName)
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
	// Zero is a deliberate value for every shortcut here — radius 0 is a sharp corner,
	// `[0, 0]` is a closed arc, thickness ratio 0 is a zero-thickness band — so each gate
	// asks whether the option was supplied (`!== undefined`), not whether it is truthy.
	// A truthy check silently dropped rectRadius: 0's guide entirely (PowerPoint then fell
	// back to the preset default rounding) and arcThicknessRatio: 0's adj3. (#24)
	if (options.rectRadius !== undefined) {
		const adjVal = Math.round((options.rectRadius * EMU_PER_INCH * PERCENT_SCALE) / Math.min(cx, cy))
		if (RECT_RADIUS_ADJ1_SHAPES.has(shapeName)) {
			emitGuide('adj1', adjVal)
			emitGuide('adj2', 0)
		} else {
			emitGuide('adj', adjVal)
		}
	} else if (options.angleRange !== undefined) {
		for (let i = 0; i < 2; i++) {
			const angle = options.angleRange[i] ?? 0
			emitGuide(`adj${i + 1}`, convertAngleUnits(angle, `angleRange[${i}]`))
		}

		if (options.arcThicknessRatio !== undefined) {
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
					'geometry/invalid-shape-adjust',
					`shapeAdjust entry ${JSON.stringify(adj)} is invalid (needs { name:string, value:number }) and was ignored.`
				)
				return
			}
			if (emittedAdjNames.has(adj.name)) {
				warn(
					'geometry/shape-adjust-overridden',
					`shapeAdjust "${adj.name}" was ignored because rectRadius/angleRange already set that handle.`
				)
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
 * Emit an `<a:custGeom>` for a freeform path built from `points`, plus any caller-supplied
 * connection sites (`connectionSites` → `<a:cxnLst>`), adjust handles (`adjustHandles` →
 * `<a:ahLst>`) and guide formulas (`guides` → `<a:gdLst>`) that make the shape connectable and
 * editable. Shared by the shape and image code paths so that path emission stays in one place.
 * Points and connection-site/handle positions are authored in the object's own inch/EMU space
 * (0..cx, 0..cy) — not slide-relative and not normalized; a bare non-`Coord` string is passed
 * through verbatim as a guide-name reference (matching `ST_AdjCoordinate`'s number-or-name union).
 * @param {ObjectOptions} options - object options carrying `points`/`guides`/`connectionSites`/`adjustHandles`
 * @param {number} cx - object width (EMU), used as the path viewport width
 * @param {number} cy - object height (EMU), used as the path viewport height
 * @param {PresLayout} layout - presentation layout used to resolve point coordinates to EMU
 * @return {string} `<a:custGeom>` XML
 */
export function genXmlCustGeom(options: ObjectOptions, cx: number, cy: number, layout: PresLayout): string {
	const points = options.points

	/**
	 * Resolve an `ST_AdjCoordinate` value: a number / unit-suffixed `Coord` → EMU (like `points`),
	 * but a token that starts with a letter/underscore is a guide-name reference emitted verbatim.
	 * `undefined` → `undefined`, so `el`/`voidEl` drop the (optional) attribute entirely.
	 */
	const adjCoord = (v: Coord | string | undefined, axis: 'X' | 'Y'): string | undefined => {
		if (v === undefined || v === null) return undefined
		if (typeof v === 'string' && /^[A-Za-z_]/.test(v.trim())) return v
		return String(getSmartParseNumber(v as Coord, axis, layout))
	}

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
				warn(
					'geometry/arc-node-point-ignored',
					'freeform arc node: x/y are ignored — an arcTo end point is computed from stAng/swAng and the radii.'
				)
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

	// custGeom preamble — the sub-lists OOXML requires before `<a:pathLst>` in this exact order:
	// adjust values (avLst), guide formulas (gdLst), adjust handles (ahLst), connection sites
	// (cxnLst), and the text rectangle (rect). `avLst` and `rect` stay as ts-pptx drives them
	// from the path; the other three are populated from caller options when supplied, and MUST
	// fall back to today's exact empty-case bytes when absent (byte-identity contract).

	// `<a:gdLst>`: named construction formulas, emitted verbatim (advanced escape hatch).
	let gdLst = el('a:gdLst')
	if (options.guides?.length) {
		const guideEls: string[] = []
		options.guides.forEach((g) => {
			// A degenerate `<a:gd>` (empty name or formula) is one PowerPoint silently repairs, so
			// warn and skip rather than emit it (mirrors the `shapeAdjust` guard in genXmlPresetGeom).
			if (
				!g ||
				typeof g.name !== 'string' ||
				g.name.length === 0 ||
				typeof g.formula !== 'string' ||
				g.formula.length === 0
			) {
				warn(
					'geometry/invalid-guide',
					`guide entry ${JSON.stringify(g)} is invalid (needs { name:string, formula:string }) and was ignored.`
				)
				return
			}
			// The formula is passed through uninterpreted, but its leading operation is a
			// closed set: an unknown one emits schema-shaped, semantically dead geometry
			// whose first feedback is a PowerPoint repair prompt. Warn and skip instead.
			const op = g.formula.trimStart().split(/\s+/)[0] ?? ''
			if (!GEOM_GUIDE_OPS.has(op)) {
				warn(
					'geometry/unknown-guide-operation',
					`guide "${g.name}" formula ${JSON.stringify(g.formula)} starts with an unknown operation "${op}" (expected one of ${[...GEOM_GUIDE_OPS].join(' ')}) and was ignored.`
				)
				return
			}
			guideEls.push(voidEl('a:gd', { name: g.name, fmla: g.formula }))
		})
		if (guideEls.length) gdLst = el('a:gdLst', null, guideEls.map(raw))
	}

	// `<a:cxnLst>`: connection sites a connector can attach to (indexed by startShapeIdx/endShapeIdx).
	let cxnLst = el('a:cxnLst')
	if (options.connectionSites?.length) {
		const cxnEls: string[] = []
		options.connectionSites.forEach((c) => {
			if (!c || typeof c.ang !== 'number' || !isFinite(c.ang)) {
				warn(
					'geometry/invalid-connection-site',
					`connectionSite entry ${JSON.stringify(c)} is invalid (needs a finite \`ang\` in degrees) and was ignored.`
				)
				return
			}
			const pos = voidEl('a:pos', { x: adjCoord(c.x, 'X'), y: adjCoord(c.y, 'Y') })
			cxnEls.push(el('a:cxn', { ang: convertAngleUnits(c.ang, 'connectionSite ang') }, raw(pos)))
		})
		if (cxnEls.length) cxnLst = el('a:cxnLst', null, cxnEls.map(raw))
	}

	// `<a:ahLst>`: draggable adjust handles — XY or polar, discriminated by any polar-only key.
	let ahLst = voidEl('a:ahLst', null, SPACE_BEFORE_SLASH)
	if (options.adjustHandles?.length) {
		const ahEls: string[] = options.adjustHandles.map((h) => {
			const pos = raw(voidEl('a:pos', { x: adjCoord(h.x, 'X'), y: adjCoord(h.y, 'Y') }))
			const isPolar = 'gdRefR' in h || 'minR' in h || 'maxR' in h || 'gdRefAng' in h || 'minAng' in h || 'maxAng' in h
			// `isPolar` is an aliased condition, so TS narrows `h` to `ShapeAdjustHandlePolar` in the
			// positive branch. The negation of a multi-`in` OR does not narrow the other way, so the
			// XY branch keeps an explicit cast.
			if (isPolar) {
				return el(
					'a:ahPolar',
					{
						gdRefR: h.gdRefR,
						minR: adjCoord(h.minR, 'X'),
						maxR: adjCoord(h.maxR, 'X'),
						gdRefAng: h.gdRefAng,
						minAng: h.minAng == null ? undefined : convertAngleUnits(h.minAng, 'adjustHandle minAng'),
						maxAng: h.maxAng == null ? undefined : convertAngleUnits(h.maxAng, 'adjustHandle maxAng'),
					},
					pos
				)
			}
			const xy = h as ShapeAdjustHandleXY
			return el(
				'a:ahXY',
				{
					gdRefX: xy.gdRefX,
					minX: adjCoord(xy.minX, 'X'),
					maxX: adjCoord(xy.maxX, 'X'),
					gdRefY: xy.gdRefY,
					minY: adjCoord(xy.minY, 'Y'),
					maxY: adjCoord(xy.maxY, 'Y'),
				},
				pos
			)
		})
		ahLst = el('a:ahLst', null, ahEls.map(raw))
	}

	return el(
		'a:custGeom',
		null,
		[
			voidEl('a:avLst', null, SPACE_BEFORE_SLASH),
			gdLst,
			ahLst,
			cxnLst,
			voidEl('a:rect', { l: 'l', t: 't', r: 'r', b: 'b' }, SPACE_BEFORE_SLASH),
			el('a:pathLst', null, raw(el('a:path', { w: cx, h: cy }, nodes.map(raw)))),
		].map(raw)
	)
}
