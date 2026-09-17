/**
 * `AnyShape` → one {@link CallIr}.
 *
 * The dispatch is on what the shape *is*, not on what it looks like: an `AutoShape` with a
 * text frame and no geometry is an `addText`, one with geometry is an `addShape` that may
 * also carry text, a `Picture` is an `addImage`, and a `GraphicFrame` delegates to the
 * table or chart mapper. A `GroupShape` recurses.
 *
 * The single most important thing this module does is record what it cannot carry.
 * Measurement against the fixture corpus found the binding constraint is the *read* side,
 * not the write side — several constructs the write API can express have no accessor at
 * all — so a converter that mapped only what it could see would produce clean-looking
 * output with silent visual losses. The clearest is theme-referenced line styling: a shape
 * whose outline comes from `p:style/a:lnRef` reports a colour through `resolvedLine` but
 * reports `null` for width and dash, because there is a resolved-colour path and no
 * resolved-width path. In the corpus that was every such shape, so a themed 2pt
 * border silently becomes a 1pt default. Each one is noted here rather than quietly thinned.
 *
 * The same dispatch serves the deck's *chrome*: {@link masterObject} re-tags a call as an
 * entry in a `defineSlideMaster({ objects })` array, so a rectangle on a layout is transcribed
 * by the code that transcribes one on a slide. Sharing it is the point — a second mapper would
 * be a second set of decisions about the same OOXML, drifting from this one silently.
 */
import {
	isAutoShape,
	isConnector,
	isGraphicFrame,
	isGroupShape,
	isPicture,
	type AnyShape,
	type AutoShape,
	type Connector,
	type CustomGeometry,
	type GraphicFrame,
	type GroupShape,
	type Picture,
} from '../../read/api/shapes.js'
import { UNWRITABLE_FRAME_CONSTRUCTS, type NoteScope, type UnwritableFramePayload } from '../fidelity.js'
import { isAssetRef, type CallIr, type IrValue } from '../ir.js'
import type { ConnectorProps, ShapeLineProps } from '../../types/style.js'
import {
	compact,
	compactRequired,
	emu,
	frameOf,
	identityOptions,
	nameOf,
	positionOptions,
	type ShapeBox,
} from './values.js'
import { hasIdentityChildSpace, isAudioVideo, isTextBox } from './detect.js'
import { noteUnreadText, textFrameOptions, textRuns } from './text.js'
import type { TextFrame } from '../../read/api/text.js'
import { tableCall } from './table.js'
import { chartCall } from './chart.js'
import { forShape, type MapContext } from './context.js'
import { cropOption, isRectSet } from './picture-fill.js'
export type { AssetResolver } from './context.js'
import { PERCENT_SCALE } from '../../units.js'

import { glowOption, lineOption, noteHidden, shadowOption, styleOptions, transformOptions } from './shape-paint.js'

/**
 * Map one shape. Returns `null` when the shape produces no call at all — a note is always
 * recorded in that case, so a dropped shape is never silent.
 */
export function shapeCall(shape: AnyShape, ctx: MapContext): CallIr | null {
	const scoped = forShape(ctx, shape)

	if (shape.hidden) {
		noteHidden(scoped.notes)
		return null
	}

	if (isGroupShape(shape)) return groupCall(shape, scoped)
	if (isPicture(shape)) return pictureCall(shape, scoped)
	if (isConnector(shape)) return connectorCall(shape, scoped.notes)
	if (isGraphicFrame(shape)) return graphicFrameCall(shape, scoped)
	if (isAutoShape(shape)) return autoShapeCall(shape, scoped)

	// Unreachable: the five guards above exhaust `AnyShape`, which is why `shape` narrows to
	// `never` here. Left as a floor so a sixth shape kind fails to compile at the call sites
	// that assume a call, rather than silently falling through to a silent drop.
	return null
}

/**
 * The shape's authorable text, and the notes for what it holds that cannot be authored.
 *
 * "Has text" means text this converter can author, which is *runs*. `TextFrame.text` also
 * counts `a:fld` field text -- a slide number, date or footer placeholder -- and there is no
 * run behind it, so measuring by `text` made every field placeholder emit an `addText`
 * carrying a single empty string: an invisible empty box where a number used to be.
 *
 * Extracted from `autoShapeCall`, which inlined this alongside the placeholder note, the
 * common-option assembly and three mutually exclusive emission arms.
 * @returns the frame when it holds authorable runs, else `null`
 */
function autoShapeText(shape: AutoShape, notes: NoteScope): TextFrame | null {
	const frame = shape.hasTextFrame ? shape.textFrame : null
	if (frame === null) return null
	return noteUnreadText(frame, shape.element_, notes, 'shape').length > 0 ? frame : null
}

function autoShapeCall(shape: AutoShape, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const box = frameOf(shape, notes)
	if (!box) return null
	const preset = shape.presetGeometry
	const custom = shape.customGeometry
	const frame = autoShapeText(shape, notes)

	if (shape.placeholder) {
		notes.note(
			'shape.placeholder',
			'flattened',
			'unsupported',
			'a placeholder is emitted as a plain shape with its inherited geometry and styling baked in, so it no longer follows its layout'
		)
	}

	const common = {
		...positionOptions(box),
		...transformOptions(shape),
		...styleOptions(shape, ctx),
		...identityOptions(shape),
	}

	if (custom) {
		noteCustGeomGuides(notes)
		const points = customGeometryPoints(box, custom)
		// A freeform holding text is an `addText` with the same points: this arm used to return
		// `addShape` before the text arm below, so a freeform's text was discarded.
		if (frame !== null) {
			const options = compact({
				...common,
				...textFrameOptions(frame, notes),
				...(isTextBox(shape.element_) ? { isTextBox: true } : {}),
				shape: 'custGeom',
				points,
			})
			return { method: 'addText', args: [textRuns(frame, ctx), options ?? {}], ...nameOf(shape) }
		}
		const options = compact({ ...common, points })
		return { method: 'addShape', args: ['custGeom', options ?? {}], ...nameOf(shape) }
	}

	if (frame !== null) {
		// `isTextBox` and `shape` are independent, not alternatives: a PowerPoint text box
		// carries `txBox="1"` *and* an explicit `prstGeom rect`, so reading text-box-ness off
		// the absence of geometry misclassified every one of them as an auto shape.
		const options = compact({
			...common,
			...textFrameOptions(frame, notes),
			...(isTextBox(shape.element_) ? { isTextBox: true } : {}),
			...(preset === null ? {} : { shape: preset, ...adjustOptions(shape) }),
		})
		return { method: 'addText', args: [textRuns(frame, ctx), options ?? {}], ...nameOf(shape) }
	}

	// No text and no geometry of its own: an empty placeholder, whose shape comes from the
	// layout it inherits. `addShape` needs a preset name and `addText` would author an
	// invisible empty box, so the shape is dropped — which matches how it renders in a
	// slideshow, where an unfilled placeholder shows nothing. It is still a loss in the
	// editor, so it is declared rather than assumed harmless.
	if (preset === null) {
		notes.note(
			'shape.empty',
			'dropped',
			'unsupported',
			'this shape carries no text and no geometry of its own (an unfilled placeholder drawing its outline from the layout), so there is no addShape preset to name and it is omitted'
		)
		return null
	}
	const options = compact({ ...common, ...adjustOptions(shape) })
	return { method: 'addShape', args: [preset, options ?? {}], ...nameOf(shape) }
}

/** The note for what a `custGeom` holds beside its path, shared by a freeform shape and a picture clipped to one. */
function noteCustGeomGuides(notes: NoteScope): void {
	notes.note(
		'shape.custGeom.guides',
		'dropped',
		'unread',
		'custGeom guides, adjust handles and connection sites (a:gdLst / a:ahLst / a:cxnLst) have no accessor, so only the path outline carries'
	)
}

/**
 * A picture's clip as `addImage` spells it: `points` for a freeform, `shape` and `shapeAdjust` for
 * a preset.
 *
 * PowerPoint writes `rect` for a picture that has not been cropped to a shape, and `addImage` writes
 * `rect` when it is given no `shape`, so a `rect` is left off rather than printed on every picture.
 */
function pictureClipOptions(shape: Picture, box: ShapeBox, notes: NoteScope): Record<string, IrValue | undefined> {
	const custom = shape.customGeometry
	if (custom) {
		noteCustGeomGuides(notes)
		return { points: customGeometryPoints(box, custom) }
	}
	const preset = shape.presetGeometry
	return preset === null || preset === 'rect' ? {} : { shape: preset, ...adjustOptions(shape) }
}

/**
 * Preset-geometry adjust handles (`a:avLst`) as `shapeAdjust`.
 *
 * The read model reports each guide's raw formula string; a preset's adjust values are
 * `val <n>` where `<n>` is in 1000ths of a percent, and the write API takes the same handle
 * as a `0.0–1.0` fraction. Anything that is not a plain `val` is a computed guide, which
 * has no write-API expression, so it is skipped rather than mis-scaled.
 */
function adjustOptions(shape: AutoShape | Picture): Record<string, IrValue | undefined> {
	const values = shape.adjustValues
	const guides: IrValue[] = []
	for (const name of Object.keys(values).sort()) {
		const match = /^val\s+(-?\d+)$/.exec((values[name] ?? '').trim())
		if (!match?.[1]) continue
		guides.push({ name, value: Number(match[1]) / PERCENT_SCALE })
	}
	return guides.length ? { shapeAdjust: guides } : {}
}

/**
 * `custGeom` path commands as the write API's `points` array.
 *
 * Path coordinates are integers in the path's own `0..w`/`0..h` viewport, not EMU, so each
 * is scaled onto the box the call is placed at before being printed as EMU. A path that
 * declares no viewport (`@w`/`@h` of 0) is already in the shape's own space and passes through.
 *
 * The box is the one {@link frameOf} placed the call at, not `absoluteFrame`: in a group whose
 * transform cannot be composed that is `null`, the shape is placed at its resolved box, and
 * scaling by 1 would print its points in raw path units.
 */
function customGeometryPoints(box: ShapeBox, custom: CustomGeometry): IrValue {
	const points: IrValue[] = []

	for (const path of custom.paths) {
		const sx = path.w > 0 ? box.width / path.w : 1
		const sy = path.h > 0 ? box.height / path.h : 1
		const px = (v: number): string => emu(v * sx)
		const py = (v: number): string => emu(v * sy)

		for (const command of path.commands) {
			switch (command.cmd) {
				case 'moveTo':
					points.push({ x: px(command.x), y: py(command.y), moveTo: true })
					break
				case 'lnTo':
					points.push({ x: px(command.x), y: py(command.y) })
					break
				case 'cubicBezTo':
					points.push({
						x: px(command.x),
						y: py(command.y),
						curve: { type: 'cubic', x1: px(command.x1), y1: py(command.y1), x2: px(command.x2), y2: py(command.y2) },
					})
					break
				case 'quadBezTo':
					points.push({
						x: px(command.x),
						y: py(command.y),
						curve: { type: 'quadratic', x1: px(command.x1), y1: py(command.y1) },
					})
					break
				case 'arcTo':
					// An arcTo carries no end point on either side — the renderer derives it from
					// the pen position, radii and swept angle — so this maps one-to-one.
					points.push({
						curve: { type: 'arc', wR: px(command.wR), hR: py(command.hR), stAng: command.stAng, swAng: command.swAng },
					})
					break
				case 'close':
					points.push({ close: true })
					break
			}
		}
	}
	return points
}

/**
 * A `Picture` becomes `addImage`, with its bytes carried as an asset.
 *
 * PowerPoint authors embedded audio and video as a picture too — the poster frame is the
 * blip, and the media itself hangs off `p:nvPr` as `a:videoFile`/`a:audioFile`. The read
 * model surfaces only the poster (`mediaKind` separates raster from SVG, not still from
 * moving), so a video would otherwise convert to a static image with nothing to show it had
 * ever moved. The write API has `addMedia`, so this is a read-side gap, and it is declared
 * rather than silently flattened.
 */
function pictureCall(shape: Picture, ctx: MapContext): CallIr | null {
	const { notes, assets } = ctx
	const box = frameOf(shape, notes)
	if (!box) return null
	if (isAudioVideo(shape.element_)) {
		notes.note(
			'media.audioVideo',
			'flattened',
			'unread',
			'this is embedded audio or video: no accessor reports the media part or its kind, only the poster frame, so it is emitted as a still image even though addMedia could author it'
		)
	}

	// The vector part wins when there is one. A PowerPoint SVG picture is an `asvg:svgBlip`
	// extension over a raster fallback, and the fallback is the lossy copy — taking it would
	// throw away the vector for no reason, since `addImage` accepts SVG bytes directly and
	// regenerates a fallback of its own.
	const partName = shape.svgPartName ?? shape.imagePartName
	if (!partName) {
		notes.note(
			'image.data',
			'dropped',
			'unsupported',
			'this picture references no embedded image part (an external or linked image), so there are no bytes to carry'
		)
		return null
	}
	const asset = assets.assetFor(partName)
	if (!asset) {
		notes.note('image.data', 'dropped', 'unsupported', `image part ${partName} is not present in the package`)
		return null
	}

	if (shape.svgPartName && shape.imagePartName) {
		notes.note(
			'image.svg',
			'approximated',
			'unsupported',
			"the vector part of this SVG picture is carried and the source's own raster fallback is not; the write path generates a fresh fallback, so a renderer that cannot draw SVG shows a different image from the one the source deck shipped"
		)
	}
	if (shape.recolor) {
		notes.note(
			'image.recolor',
			'dropped',
			'unwritable',
			'a picture recolour (a:duotone / a:clrChange / a:grayscl) is read but has no write-API option'
		)
	}

	// The same guard a picture fill's crop takes. A negative inset (PowerPoint writes one for a fit
	// crop) or opposite insets summing to 100% or more have no `crop` spelling, and passing them
	// through put a crop in the IR that made the printed script throw at `addImage`.
	const crop = cropOption(shape.crop)
	if (crop === undefined && isRectSet(shape.crop)) {
		notes.note(
			'image.crop',
			'approximated',
			'unwritable',
			"this picture's source crop (a:srcRect) has an inset outside the 0–100% `crop` accepts, or opposite insets that leave no image, so the picture is emitted uncropped"
		)
	}
	const options = compact({
		...positionOptions(box),
		...transformOptions(shape),
		...identityOptions(shape),
		...pictureClipOptions(shape, box, notes),
		// `ImageBaseProps.shadow` takes the same `ShadowProps` every other shape does; a picture's
		// shadow was read and never mapped.
		shadow: shadowOption(shape, notes),
		// Nor was its border, though `ImageBaseProps.line` takes the same `ShapeLineProps` and the
		// read side has always reported a picture's own `p:spPr/a:ln`. It goes through the shared
		// `lineOption`, so a picture's outline takes the same colour ladder and the same loss notes
		// as every other shape's.
		line: lineOption(shape, notes),
		// `crop`, not `sizing`. Both exist and they are not interchangeable: `crop` is
		// `a:srcRect` emitted verbatim as percentage edge insets — the exact model the read
		// model reports — while `sizing: 'crop'` cuts a window in *displayed inches* against
		// the image's measured natural size. Feeding fractions to `sizing` read them as
		// inches, which shrank every cropped picture to a fraction of its box.
		crop,
	})
	return { method: 'addImage', args: [{ ...options, data: asset }], ...nameOf(shape) }
}

/**
 * `ConnectorProps` keys that {@link lineOption} can produce. A connector styles its stroke
 * with *flat* options rather than the nested `line` object every other shape takes, and it
 * accepts a strict subset: no `transparency`, no gradient stroke, and its own `type` names
 * the routing (`straight`/`elbow`/`curved`), not a fill kind — so spreading a line option
 * object into it wholesale would both drop silently and collide.
 */
const CONNECTOR_LINE_KEYS = [
	'color',
	'width',
	'dashType',
	'beginArrowType',
	'endArrowType',
] as const satisfies readonly (keyof ConnectorProps & keyof ShapeLineProps)[]

/**
 * Every other `ShapeLineProps` key: what a connector's flat stroke options cannot express, which
 * the `connector.line` note declares.
 *
 * Typed as the complement of {@link CONNECTOR_LINE_KEYS}, so a stroke option the write API gains
 * does not compile here until it is placed, carried or noted. Unplaced, it would reach `lineOption`'s
 * output and the copy loop would drop it without a note. `type` is the fill kind, and gets this far
 * only as `'gradient'`: a `'none'` outline has already become a `line` shape.
 */
const CONNECTOR_FLATTENED_LINE_KEYS: Record<
	Exclude<keyof ShapeLineProps, (typeof CONNECTOR_LINE_KEYS)[number]>,
	true
> = {
	type: true,
	gradient: true,
	pattern: true,
	transparency: true,
	cap: true,
}

/**
 * A `Connector` becomes `addConnector`.
 *
 * The endpoints are the one real translation here. OOXML gives a connector a bounding box
 * plus flip flags; `addConnector` takes two points. `a:flipH`/`a:flipV` are what say which
 * diagonal of the box the line runs along, so dropping them would silently mirror every
 * connector that runs up or leftward.
 *
 * The *bindings* (`a:stCxn`/`a:endCxn`) reference source shape ids that have no counterpart
 * in the output, so a connector lands unbound and stops following its shapes when they move.
 */
function connectorCall(shape: Connector, notes: NoteScope): CallIr | null {
	const drawn = connectorLine(
		shape,
		notes,
		'connector endpoint bindings reference source shape ids with no counterpart in the output, so the connector lands unbound and no longer follows its shapes'
	)
	if (!drawn) return null
	// A connector with no outline has no `addConnector` spelling: its stroke options take no `type`,
	// and one given none is drawn with the write path's default line. The `line` shape it paints as
	// states `line: { type: 'none' }` and stays invisible, the way a layout's connectors already do.
	if (drawn.line?.['type'] === 'none') {
		return { method: 'addShape', args: ['line', drawn.options], ...nameOf(shape) }
	}
	const frame = drawn.box

	if (shape.rotation) {
		notes.note(
			'connector.rotation',
			'dropped',
			'unwritable',
			'addConnector takes two endpoints and no rotation, so a rotated connector lands along the unrotated diagonal of its box'
		)
	}

	const line = drawn.line ?? {}
	const stroke: Record<string, IrValue> = {}
	for (const key of CONNECTOR_LINE_KEYS) {
		const value = line[key]
		if (value !== undefined) stroke[key] = value
	}
	if (Object.keys(CONNECTOR_FLATTENED_LINE_KEYS).some((key) => line[key] !== undefined)) {
		notes.note(
			'connector.line',
			'flattened',
			'unwritable',
			'a connector styles its stroke with flat colour, width, dash and arrowhead options, which cannot express a gradient stroke, a stroke transparency or a line cap, so a gradient or a transparency falls back to a plain line and a cap to the default'
		)
	}

	const right = frame.left + frame.width
	const bottom = frame.top + frame.height
	return {
		method: 'addConnector',
		args: [
			compact({
				// Routing is not readable — `presetGeometry` names the connector preset but the
				// bend count it implies is not exposed — so every connector emits straight.
				type: 'straight',
				x1: emu(shape.flipH ? right : frame.left),
				y1: emu(shape.flipV ? bottom : frame.top),
				x2: emu(shape.flipH ? frame.left : right),
				y2: emu(shape.flipV ? frame.top : bottom),
				...identityOptions(shape),
				...stroke,
			}) ?? {},
		],
		...nameOf(shape),
	}
}

/**
 * Why a graphic frame cannot be authored from a script, or `null` when it can.
 *
 * The single enumeration of the payloads with no write-API emitter, and the reason it is a
 * function rather than a condition spelled at each site: `graphicFrameCall` needs it to pick
 * a fidelity note, and `hasUnwritableContent` in `from-read/deck.ts` needs it to decide
 * whether the printer must copy the slide instead of transcribing it. Those were two
 * hand-maintained lists, they disagreed — the deck-level one named extended charts only — and
 * a SmartArt slide was therefore transcribed with a hole where the graphic had been. The two
 * consequences are the same for every member, so the membership has to be too.
 */
export function unwritableFramePayload(shape: GraphicFrame): UnwritableFramePayload | null {
	if (shape.table || shape.chart) return null
	if (shape.hasChartEx) return 'chartEx'
	if (shape.hasDiagram) return 'diagram'
	return 'unknown'
}

/** A `GraphicFrame` hosts a table or a chart; both have their own mapper. */
function graphicFrameCall(shape: GraphicFrame, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const table = shape.table
	if (table) return tableCall(shape, table, ctx)

	const chart = shape.chart
	if (chart) return chartCall(shape, chart, notes)

	// Everything below emits no call. Reached even when the deck-level walk already marked the
	// slide `carried`, because the standalone tier has no source package to copy from and needs
	// the per-shape note to say what it lost.
	switch (unwritableFramePayload(shape)) {
		case 'chartEx':
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.chartEx,
				'dropped',
				'unwritable',
				'an extended chart (waterfall, funnel, box-and-whisker, …) has a full reader but no write-API counterpart'
			)
			return null
		case 'diagram':
			// Read but unwritable, so it takes the `chartEx.all` shape rather than the
			// `graphicFrame.unknown` one below: the loss is the *write* side's, and saying
			// "not decoded" of a construct whose text the read model now returns would be false.
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.diagram,
				'dropped',
				'unwritable',
				'a SmartArt diagram has a full reader and its text can be edited in place through it (DiagramPoint.text), but no write API authors one from a script, so a converted script cannot describe this frame'
			)
			return null
		default:
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.unknown,
				'dropped',
				'unread',
				'this graphic frame hosts none of a table, a chart or a SmartArt diagram (an OLE object or an ink annotation), which the read model does not decode'
			)
			return null
	}
}

/**
 * A `GroupShape` becomes `addGroup`, with each child mapped and re-tagged into the
 * key-tagged `GroupChildProps` shape the write API expects.
 *
 * Children keep slide-absolute coordinates, which is not a compromise here: `addGroup`
 * holds an identity child coordinate space (`chOff/chExt == off/ext`) at every depth
 * precisely so they can. Nested groups therefore nest rather than flatten. What does not
 * survive is a source group with a *non*-identity child space — that group scales its
 * contents when resized and the emitted one will not — so that case, and only that case,
 * is noted.
 *
 * Charts, tables and placeholders are excluded from groups by the write API, so a group holding
 * one loses that child. A connector is kept as the `line` shape it paints as, which a group takes.
 */
function groupCall(shape: GroupShape, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const box = frameOf(shape, notes)
	if (!box) return null
	const children: IrValue[] = []
	for (const child of shape.shapes) {
		// `addConnector` has no group child variant, and a connector child used to be dropped.
		if (isConnector(child)) {
			const drawn = connectorLine(
				child,
				forShape(ctx, child).notes,
				'this connector is bound to shapes; it is re-authored as a line shape inside its group, which paints the same stroke but no longer follows them'
			)
			if (drawn) children.push(lineShape(drawn.options))
			continue
		}
		const call = shapeCall(child, ctx)
		if (!call) continue
		const tagged = asGroupChild(call)
		if (!tagged) {
			notes.note(
				'group.child',
				'dropped',
				'unsupported',
				`addGroup accepts no ${kindLabel(call.method)} child (a group takes text, shapes, images and other groups), so this child is omitted from the group`
			)
			continue
		}
		children.push(tagged)
	}
	if (children.length === 0) {
		notes.note(
			'group.empty',
			'dropped',
			'unsupported',
			'no child of this group produced a call, and addGroup rejects an empty child list, so the group itself is omitted too'
		)
		return null
	}

	if (!hasIdentityChildSpace(shape.element_)) {
		notes.note(
			'group.childSpace',
			'flattened',
			'unsupported',
			"this group scales its contents (its a:chOff/a:chExt differ from its a:off/a:ext); addGroup's child space is always identity, so the children are emitted pre-scaled and resizing the group no longer rescales them"
		)
	}

	// The group's own rotation and flips are deliberately NOT emitted here. `absoluteFrame`
	// composes them into every child, and the children above were built from it, so the
	// transform is already baked in at slide-absolute coordinates. Emitting it on the group
	// as well would apply it a second time — measured: a group rotated 30° came back at 60°
	// with its children walked off their positions, and a flipped group came back unflipped
	// because the double flip cancelled. The group's frame is a bounding box and nothing else.
	if (shape.rotation !== 0 || shape.flipH || shape.flipV) {
		notes.note(
			'group.transform',
			'flattened',
			'unsupported',
			"the group's own rotation and flips are baked into its children's coordinates rather than kept on the group, because addGroup would otherwise apply them twice; the shapes render identically but rotating or resetting the group no longer moves them together"
		)
	}

	return {
		method: 'addGroup',
		args: [children, compact({ ...positionOptions(box), ...identityOptions(shape) }) ?? {}],
		...nameOf(shape),
	}
}

/** A call's object kind as a note's prose names it: `addImage` is `image`. */
function kindLabel(method: string): string {
	return method.replace(/^add/, '').toLowerCase()
}

/**
 * Re-tag a call as a `GroupChildProps` entry, or `null` when the content type is one
 * `addGroup` excludes. The tagged form is a single-key object, not a `{ type, args }` pair.
 */
function asGroupChild(call: CallIr): IrValue | null {
	const shared = commonShapeVariant(call)
	if (shared) return shared
	const [first, second] = call.args
	switch (call.method) {
		case 'addGroup':
			return { group: compact({ children: first, options: second }) ?? {} }
		default:
			// Charts and tables have no GroupChildProps variant. A connector never arrives here:
			// `groupCall` keeps it as the line shape it paints as.
			return null
	}
}

/**
 * The three key-tagged variants both union-shaped call sites emit identically: a text box, an
 * auto shape, and an image. `null` when `call.method` is not one of them, which is the signal
 * for the caller to try its own extra arms.
 *
 * Shared rather than duplicated because this is exactly where a fourth common variant would be
 * added to one site and forgotten in the other. What is deliberately *not* shared is either
 * caller's `default`: a group child silently declines a chart, while a layout object reports it
 * as a fidelity loss, and that difference is the point of having two functions.
 */
function commonShapeVariant(call: CallIr): IrValue | null {
	const [first, second] = call.args
	switch (call.method) {
		case 'addText':
			return { text: compact({ text: first, options: second }) ?? {} }
		case 'addShape':
			return { shape: compact({ type: first, options: second }) ?? {} }
		case 'addImage':
			return { image: first ?? {} }
		default:
			return null
	}
}

/* ===== a layout's own shapes ===== */

/**
 * One shape from a **layout's** shape tree as a `SlideMasterObject` — an entry in the
 * `objects` array of a `defineSlideMaster` call.
 *
 * Everything about the shape itself is decided by {@link shapeCall}, so a rectangle on a
 * layout is transcribed by the same code that transcribes one on a slide and cannot drift
 * from it. Only the *envelope* differs: `objects` is a key-tagged union rather than a method
 * name plus arguments — the same shape `GroupChildProps` takes — which is why this reads as a
 * sibling of {@link asGroupChild}.
 *
 * `notes` must already be a `layoutShapeScope`, so the slide vocabulary this borrows lands
 * under `layout.` and cannot be mistaken for a loss on a slide.
 *
 * Two methods have no variant in the union. `addGroup` never reaches here — the layout walk
 * flattens a group into its children first, since children already carry slide-absolute
 * coordinates. `addTable` genuinely has nowhere to go, and is the one kind this reports as
 * lost.
 */
export function masterObject(shape: AnyShape, ctx: MapContext): IrValue | null {
	// A connector reaches a layout only as a line shape, so it bypasses `shapeCall`.
	if (isConnector(shape)) {
		const drawn = connectorLine(
			shape,
			forShape(ctx, shape).notes,
			'this connector is bound to shapes on the layout; it is re-authored as a line shape, which paints the same stroke but no longer follows them'
		)
		return drawn && lineShape(drawn.options)
	}

	const call = shapeCall(shape, ctx)
	if (!call) return null

	const shared = commonShapeVariant(call)
	if (shared) return shared

	const [first, second] = call.args
	switch (call.method) {
		case 'addChart': {
			// `addChart(data, options)` carries the chart type inside its options, while a master
			// object names it separately. Lifted back out rather than moved: `addChildDefinition`
			// passes the whole options object on to `addChartDefinition` as well, which is where a
			// chart's `type` is read from in the first place.
			const options = second !== null && typeof second === 'object' && !Array.isArray(second) && !isAssetRef(second) ? second : null // prettier-ignore
			const type = options?.['type']
			// Unreachable today - `chartCall` returns `null` for a chart with no type or no data,
			// so this arm never sees one - but `collectObjects` states that a `null` from here is
			// always already noted, and a silent `return null` is the one thing that could break
			// that invariant if `chartCall` ever loosens.
			if (type === undefined || first === undefined) {
				forShape(ctx, shape).notes.note(
					'decoration',
					'dropped',
					'unwritable',
					'a chart on a slide layout that names no type or carries no series has no defineSlideMaster({ objects }) variant, so it is dropped from the layout the output rebuilds'
				)
				return null
			}
			return { chart: compact({ type, data: first, opts: options }) ?? {} }
		}
		default:
			forShape(ctx, shape).notes.note(
				'decoration',
				'dropped',
				'unwritable',
				`a ${kindLabel(call.method)} on a slide layout has no defineSlideMaster({ objects }) variant — the union covers a chart, an image, a shape and a text box — so it is dropped from the layout the output rebuilds`
			)
			return null
	}
}

/** What {@link connectorLine} reads off a connector. */
interface ConnectorLine {
	/** Where the connector's box sits. */
	box: ShapeBox
	/** Its stroke, as {@link lineOption} reads it. */
	line: Record<string, IrValue> | undefined
	/** The options of the `line` shape the connector paints as. */
	options: Record<string, IrValue>
}

/**
 * A `Connector` as the `line` shape it paints as, or `null` when it produces nothing: it is hidden,
 * or nothing places it. Either way the note has been recorded.
 *
 * A layout reaches a connector only this way. The alternative is dropping every connector a layout
 * carries, and in the fixture corpus that is 18 of the 45 shapes every layout actually draws,
 * because PowerPoint's line tool authors a `p:cxnSp`, so a plain rule under a title is usually a
 * connector rather than a shape. Drawing it as a `line` preset paints the identical stroke, and a
 * shape carries `rotate`, which `addConnector` does not. What that costs is the endpoint binding,
 * which on a layout costs nothing that was working: `a:stCxn`/`a:endCxn` reference shapes in the
 * layout's own tree.
 *
 * `addConnector` reads its endpoints and stroke from the same place, so the frame, the binding note
 * and the stroke are decided once for every way a connector is emitted. The binding note's prose
 * is the caller's, since what the loss means depends on where the connector lands.
 * @param shape - the connector
 * @param notes - the scope bound to this connector
 * @param bindingDetail - the endpoint-binding note's detail, for where this connector lands
 */
function connectorLine(shape: Connector, notes: NoteScope, bindingDetail: string): ConnectorLine | null {
	if (shape.hidden) {
		noteHidden(notes)
		return null
	}
	const box = frameOf(shape, notes)
	if (!box) return null

	if (shape.startConnection || shape.endConnection) {
		notes.note('connector.binding', 'dropped', 'unsupported', bindingDetail)
	}

	const line = lineOption(shape, notes)
	// No `fill`: a connector's `p:spPr` has no fill that a line geometry could show, and asking
	// `fillOption` for one would resolve the `p:style/a:fillRef` a `p:cxnSp` always carries into
	// a colour that paints nothing on the source and a filled box on the output.
	const options = compactRequired({
		...positionOptions(box),
		...transformOptions(shape),
		line,
		shadow: shadowOption(shape, notes),
		glow: glowOption(shape, notes),
		...identityOptions(shape),
	})
	return { box, line, options }
}

/**
 * A connector's `line` shape as a key-tagged descriptor, the form a layout object and a group child
 * both take. Through `compact` like the other arms, so every emitted descriptor spells its keys in
 * the same order and the printed script does not read as two different mappers.
 */
function lineShape(options: Record<string, IrValue>): IrValue {
	return { shape: compact({ type: 'line', options }) ?? {} }
}
