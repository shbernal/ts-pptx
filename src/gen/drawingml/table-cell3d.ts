/**
 * ts-pptx: DrawingML table-cell 3-D bevel
 *
 * Emit the `<a:cell3D>` child of a table cell's `<a:tcPr>`. `CT_Cell3D` is a **required**
 * `a:bevel` plus an optional `a:lightRig`, which is why an empty `cell3D: {}` still produces
 * a bevel element (at the schema's own defaults) rather than an empty, invalid `a:cell3D`.
 *
 * The surface is deliberately minimal — preset, size, material and light rig. PowerPoint's
 * table UI exposes no control for any of it, so a cell3D reaches a deck from a theme or
 * another producer; this exists to author and reproduce one, not to model DrawingML 3-D.
 */

import type { TableCell3DProps } from '../../types/index.js'
import { warnOnce } from '../../diagnostics.js'
import { valToPts } from '../../units-internal.js'
import { el, raw, voidEl } from '../oxml/el.js'

/** `ST_BevelPresetType`. */
const BEVEL_PRESETS: readonly string[] = [
	'relaxedInset',
	'circle',
	'slope',
	'cross',
	'angle',
	'softRound',
	'convex',
	'coolSlant',
	'divot',
	'riblet',
	'hardEdge',
	'artDeco',
]

/** `ST_PresetMaterialType`. */
const MATERIALS: readonly string[] = [
	'legacyMatte',
	'legacyPlastic',
	'legacyMetal',
	'legacyWireframe',
	'matte',
	'plastic',
	'metal',
	'warmMatte',
	'translucentPowder',
	'powder',
	'dkEdge',
	'softEdge',
	'clear',
	'flat',
	'softmetal',
]

/** `ST_LightRigType`. */
const LIGHT_RIGS: readonly string[] = [
	'legacyFlat1',
	'legacyFlat2',
	'legacyFlat3',
	'legacyFlat4',
	'legacyNormal1',
	'legacyNormal2',
	'legacyNormal3',
	'legacyNormal4',
	'legacyHarsh1',
	'legacyHarsh2',
	'legacyHarsh3',
	'legacyHarsh4',
	'threePt',
	'balanced',
	'soft',
	'harsh',
	'flood',
	'contrasting',
	'morning',
	'sunrise',
	'sunset',
	'chilly',
	'freezing',
	'flat',
	'twoPt',
	'glow',
	'brightRoom',
]

/** `ST_LightRigDirection`. */
const LIGHT_DIRECTIONS: readonly string[] = ['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br']

/**
 * Vet one enumerated value before it reaches the XML. Anything outside its `ST_` union would
 * make the slide part schema-invalid, and PowerPoint reports that as a corrupt file rather
 * than as a mis-set option — so an unrecognized value is reported and the attribute dropped,
 * leaving the renderer on the schema default.
 */
function resolveEnum(value: string | undefined, valid: readonly string[], field: string): string | null {
	if (value === undefined || value === null) return null
	if (valid.includes(value)) return value
	warnOnce(
		'table/invalid-cell3d',
		`table cell: cell3D \`${field}\` value \`${String(value)}\` is not valid and is ignored — ` +
			`use one of ${valid.join(', ')}.`,
		{ received: value, field, valid }
	)
	return null
}

/**
 * A bevel dimension in EMU (`ST_PositiveCoordinate`), from a width in points.
 * A negative or non-finite value cannot be written, so it falls back to the attribute's
 * own default of 76200 EMU (6pt) rather than emitting something PowerPoint would reject.
 */
function bevelSizeEmu(points: number | undefined, field: string): number | null {
	if (points === undefined || points === null) return null
	if (!Number.isFinite(points) || points < 0) {
		warnOnce(
			'table/invalid-cell3d',
			`table cell: cell3D \`${field}\` must be a non-negative number of points; \`${String(points)}\` is ignored.`,
			{ received: points, field }
		)
		return null
	}
	return valToPts(points)
}

/**
 * Emit a cell's `<a:cell3D>`, or `''` when it has none.
 * @param {TableCell3DProps} [cell3D] - the cell's 3-D options
 * @return {string} the `a:cell3D` element XML
 */
export function genTableCell3DXml(cell3D?: TableCell3DProps): string {
	if (!cell3D || typeof cell3D !== 'object') return ''
	const bevel = voidEl('a:bevel', {
		w: bevelSizeEmu(cell3D.width, 'width'),
		h: bevelSizeEmu(cell3D.height, 'height'),
		prst: resolveEnum(cell3D.preset, BEVEL_PRESETS, 'preset'),
	})
	// Both `rig` and `dir` are required on CT_LightRig, so a half-specified rig is dropped
	// whole — emitting one attribute would produce a part PowerPoint calls corrupt.
	const rig = resolveEnum(cell3D.lightRig?.rig, LIGHT_RIGS, 'lightRig.rig')
	const dir = resolveEnum(cell3D.lightRig?.dir, LIGHT_DIRECTIONS, 'lightRig.dir')
	const lightRig = rig && dir ? voidEl('a:lightRig', { rig, dir }) : ''
	if (cell3D.lightRig && !lightRig) {
		warnOnce(
			'table/invalid-cell3d',
			'table cell: cell3D `lightRig` needs both `rig` and `dir` (both are required by CT_LightRig); ' +
				'the light rig is dropped and the bevel takes the renderer default.',
			{ received: cell3D.lightRig }
		)
	}
	return el('a:cell3D', { prstMaterial: resolveEnum(cell3D.material, MATERIALS, 'material') }, [
		raw(bevel),
		lightRig ? raw(lightRig) : null,
	])
}
