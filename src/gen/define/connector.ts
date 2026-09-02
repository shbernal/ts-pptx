/**
 * ts-pptx: Connector Definition
 *
 * `addConnectorDefinition` converts two endpoints to a flip-oriented bounding box, resolves the
 * connector preset + adjust guides, and optionally binds start / end shapes. Emitted later as
 * `<p:cxnSp>`.
 */
import { connectorPresetFor, SlideObjectType } from '../../enums.js'
import { DEF_SHAPE_LINE_COLOR } from '../../constants-internal.js'
import { warn } from '../../diagnostics.js'
import type { ConnectorProps } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { EMU_PER_INCH, FIXED_PCT_PER_PERCENT } from '../../units.js'
import { getSmartParseNumber } from '../../units-internal.js'
import { resolveObjectName } from './object-name.js'
import { InvalidOptionError } from '../../errors.js'
import { pickDefined } from '../../options-internal.js'

/**
 * Adds a connector object to a slide definition.
 * A connector is a line between two points emitted as a PowerPoint connector (`<p:cxnSp>`).
 * Endpoints are converted to a bounding box (`x/y/w/h`) plus `flipH`/`flipV` so the box can be
 * oriented from any corner; the connector preset geometry is derived from `type`.
 * @param {PresSlideInternal} target - slide the connector is added to
 * @param {ConnectorProps} opts - connector options (endpoints + line styling)
 */
export function addConnectorDefinition(target: PresSlideInternal, opts: ConnectorProps): void {
	if (!opts || [opts.x1, opts.y1, opts.x2, opts.y2].some((v) => typeof v === 'undefined')) {
		throw new InvalidOptionError(
			'connector/missing-endpoints',
			'addConnector requires { x1, y1, x2, y2 }. Example: `slide.addConnector({ x1:1, y1:1, x2:4, y2:3 })`'
		)
	}

	const type = opts.type || 'straight'
	if (type !== 'straight' && type !== 'elbow' && type !== 'curved') {
		throw new InvalidOptionError(
			'connector/invalid-type',
			`Invalid connector type "${String(type)}". Use 'straight', 'elbow', or 'curved'.`
		)
	}

	// Resolve the preset variant + adjust guides. `bentConnector{3,4,5}` / `curvedConnector{3,4,5}`
	// each expose `bends` adjustable jogs as `<a:gd name="adjN" fmla="val …"/>` (1000ths-of-a-percent,
	// so 50% → 50000; values verified against PowerPoint-authored decks). `straightConnector1` has none.
	const adjInput = opts.adj === undefined ? [] : Array.isArray(opts.adj) ? opts.adj : [opts.adj]
	const bends = opts.bends ?? (adjInput.length || 1)
	let connectorAdj: number[] = []
	if (type === 'straight') {
		if (opts.bends !== undefined || opts.adj !== undefined) {
			warn(
				'connector/bends-ignored-for-straight',
				'addConnector `bends`/`adj` are ignored for type "straight" (a straight connector has no bends).'
			)
		}
	} else {
		if (bends !== 1 && bends !== 2 && bends !== 3) {
			throw new InvalidOptionError(
				'connector/invalid-bends',
				`addConnector \`bends\` must be 1, 2, or 3 (got ${String(bends)}).`
			)
		}
		if (opts.adj !== undefined && adjInput.length !== bends) {
			throw new InvalidOptionError(
				'connector/adj-count-mismatch',
				`addConnector \`adj\` must supply ${bends} value(s) to match \`bends\`=${bends} (got ${adjInput.length}).`
			)
		}
		// Convert each percent to OOXML 1000ths-of-a-percent. Fail loud on non-finite input
		// (silent coercion would emit a degenerate guide PowerPoint repairs); warn but allow
		// out-of-range, which legitimately places a jog beyond the endpoint box.
		connectorAdj = adjInput.map((pct, i) => {
			if (typeof pct !== 'number' || !Number.isFinite(pct)) {
				throw new InvalidOptionError(
					'connector/adj-non-finite',
					`addConnector \`adj\` value #${i + 1} must be a finite number (percent 0–100); got ${String(pct)}.`
				)
			}
			if (pct < 0 || pct > 100) {
				warn(
					'connector/adj-out-of-range',
					`addConnector \`adj\` value ${pct} is outside 0–100; the bend will sit beyond the endpoint box.`
				)
			}
			return Math.round(pct * FIXED_PCT_PER_PERCENT)
		})
	}
	const preset = connectorPresetFor(type, bends)

	// Optional shape binding (<a:stCxn>/<a:endCxn>). The target id is resolved at serialize time
	// (it equals the shape's slide-object index + 2); here we just capture the name + site index.
	// The site index must be a non-negative integer — a bad idx makes PowerPoint repair the connector.
	const resolveCxn = (
		shapeName: string | undefined,
		idx: number | undefined,
		end: 'startShape' | 'endShape'
	): { name: string; idx: number } | undefined => {
		if (shapeName === undefined) return undefined
		if (typeof shapeName !== 'string' || shapeName.trim().length === 0) {
			throw new InvalidOptionError(
				'connector/invalid-binding-name',
				`addConnector \`${end}\` must be a non-empty shape objectName.`
			)
		}
		const site = idx ?? 0
		if (!Number.isInteger(site) || site < 0) {
			throw new InvalidOptionError(
				'connector/invalid-connection-site',
				`addConnector \`${end}Idx\` must be a non-negative integer (got ${String(site)}).`
			)
		}
		// Stored as the caller spelled it: `resolveObjectNameToId` escapes its own lookup key at
		// serialize time (see its docblock), so escaping here too would look up `Q&amp;amp;A`.
		return { name: shapeName, idx: site }
	}
	const startCxn = resolveCxn(opts.startShape, opts.startShapeIdx, 'startShape')
	const endCxn = resolveCxn(opts.endShape, opts.endShapeIdx, 'endShape')
	const objectName = resolveObjectName(target, SlideObjectType.connector, {
		label: 'Connector',
		base: 0,
		kind: 'connector',
		supplied: opts.objectName,
	})

	// Resolve all four endpoints to inches up front (handles every `Coord` form: number,
	// '50%', '2in', etc.). The connector box uses the min corner as its origin and flips
	// horizontally/vertically when the end point is left of / above the start point.
	const x1 = getSmartParseNumber(opts.x1, 'X', target._presLayout) / EMU_PER_INCH
	const y1 = getSmartParseNumber(opts.y1, 'Y', target._presLayout) / EMU_PER_INCH
	const x2 = getSmartParseNumber(opts.x2, 'X', target._presLayout) / EMU_PER_INCH
	const y2 = getSmartParseNumber(opts.y2, 'Y', target._presLayout) / EMU_PER_INCH

	const newObject: SlideObject = {
		_type: SlideObjectType.connector,
		// store the connector preset on `shape`; the serializer emits it as the prstGeom `prst`
		shape: preset,
		options: {
			x: Math.min(x1, x2),
			y: Math.min(y1, y2),
			w: Math.abs(x2 - x1),
			h: Math.abs(y2 - y1),
			flipH: x2 < x1,
			flipV: y2 < y1,
			// A straight connector has no adjust values and an unattached end no `cxn`; each is one
			// absent key rather than a key holding `undefined`, which is what the emitters test for.
			...(connectorAdj.length ? { _connectorAdj: connectorAdj } : {}),
			...pickDefined({ _startCxn: startCxn, _endCxn: endCxn }, ['_startCxn', '_endCxn']),
			line: {
				type: 'solid',
				color: opts.color || DEF_SHAPE_LINE_COLOR,
				width: typeof opts.width === 'number' ? opts.width : 1,
				dashType: opts.dashType || 'solid',
				...pickDefined(opts, ['beginArrowType', 'endArrowType']),
			},
			...pickDefined(opts, ['altText']),
			objectName,
		},
	}

	target._slideObjects.push(newObject)
}
