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
 * resolved-width path. In the corpus that was every such shape, 25 of 25, so a themed 2pt
 * border silently becomes a 1pt default. Each one is noted here rather than quietly thinned.
 */
import type {
	AnyShape,
	AutoShape,
	Connector,
	CustomGeometry,
	GraphicFrame,
	GroupShape,
	Picture,
} from '../../read/api/shapes.js'
import { isAutoShape, isConnector, isGraphicFrame, isGroupShape, isPicture } from '../../read/api/shapes.js'
import type { GradientFill } from '../../read/api/gradient.js'
import type { NoteScope } from '../fidelity.js'
import type { AssetRef, CallIr, IrValue } from '../ir.js'
import { alphaToTransparency, compact, emu, isWritableSchemeToken, literalColor, orUndefined } from './values.js'
import { pictureFillOption, type PictureFillSubject } from './picture-fill.js'
import { hasEquation, hasIdentityChildSpace, isAudioVideo, isTextBox } from './detect.js'
import { textFrameOptions, textRuns } from './text.js'
import { tableCall } from './table.js'
import { chartCall } from './chart.js'

/** Resolves an image/media part name to bytes the deck-level walk has registered. */
export interface AssetResolver {
	/** Register a part's bytes and hand back the reference that stands in for them. */
	assetFor(partName: string): AssetRef | null
	/**
	 * The part's content type, or `null` when it is not in the package.
	 *
	 * Separate from {@link assetFor} because it has to be answerable *without* registering
	 * anything: a caller that rejects a part on its type (a picture fill's SVG blip, which
	 * the write path refuses) would otherwise leave bytes in the asset list that no call
	 * references, and those bytes are emitted as a file next to the script.
	 */
	contentTypeOf(partName: string): string | null
}

/** How {@link pictureFillOption}'s notes name a shape's surface. */
const SHAPE_PICTURE_FILL: PictureFillSubject = {
	construct: 'fill.picture',
	subject: "this shape's surface",
	element: 'a:blipFill',
}

/** Arrowhead types `ShapeLineProps` accepts; `a:headEnd/@type` uses the same tokens. */
const WRITABLE_ARROWS = new Set(['none', 'arrow', 'diamond', 'oval', 'stealth', 'triangle'])

/** Dash tokens `ShapeLineProps.dashType` accepts — the whole `ST_PresetLineDashVal` set. */
const WRITABLE_DASHES = new Set([
	'solid',
	'dot',
	'dash',
	'lgDash',
	'dashDot',
	'lgDashDot',
	'lgDashDotDot',
	'sysDash',
	'sysDot',
	'sysDashDot',
	'sysDashDotDot',
])

/**
 * Map one shape. Returns `null` when the shape produces no call at all — a note is always
 * recorded in that case, so a dropped shape is never silent.
 */
export function shapeCall(shape: AnyShape, notes: NoteScope, assets: AssetResolver): CallIr | null {
	const scoped = notes.forShape(shape.name || null)

	if (shape.hidden) {
		// `p:cNvPr/@hidden` has no write-API counterpart, so a hidden shape would come back
		// visible. Omitting it preserves what the deck looks like, which is the lesser loss.
		scoped.note(
			'shape.hidden',
			'dropped',
			'unwritable',
			'a hidden shape has no write-API expression; it is omitted rather than emitted visible'
		)
		return null
	}

	if (isGroupShape(shape)) return groupCall(shape, scoped, assets)
	if (isPicture(shape)) return pictureCall(shape, scoped, assets)
	if (isConnector(shape)) return connectorCall(shape, scoped)
	if (isGraphicFrame(shape)) return graphicFrameCall(shape, scoped, assets)
	if (isAutoShape(shape)) return autoShapeCall(shape, scoped, assets)

	// Unreachable: the five guards above exhaust `AnyShape`, which is why `shape` narrows to
	// `never` here. Left as a floor so a sixth shape kind fails to compile at the call sites
	// that assume a call, rather than silently falling through to a silent drop.
	return null
}

/* ===== position, fill, line, effects — shared by every shape kind ===== */

/**
 * Position as `Coord`-typed EMU strings, so the source geometry survives exactly.
 *
 * `absoluteFrame` rather than the raw `left`/`top`: a shape inside a group is positioned in
 * its group's child coordinate space, which a flattened call list cannot express, so the
 * group transform has to be composed in here. For a top-level shape the two are identical.
 */
function positionOptions(shape: AnyShape, notes?: NoteScope): Record<string, IrValue> {
	const frame = shape.absoluteFrame
	if (frame) return { x: emu(frame.left), y: emu(frame.top), w: emu(frame.width), h: emu(frame.height) }

	// No transform of its own — a placeholder taking its geometry from the layout or master.
	// `resolvedFrame` walks that chain, and using it is not optional: omitting the geometry
	// does *not* leave it to be inherited, because a regenerated slide's placeholder inherits
	// nothing. It produces `x=0 y=0 w=<slide width> h=0` — a zero-height box in the corner,
	// which is broken output rather than lossy output.
	const resolved = shape.resolvedFrame
	if (!resolved) return {}
	notes?.note(
		'shape.frameInherited',
		'flattened',
		'unsupported',
		`this shape has no transform of its own and takes its geometry from the ${resolved.source}; a regenerated slide inherits nothing, so the resolved position is baked in and stops tracking later edits to that ${resolved.source}`
	)
	return { x: emu(resolved.left), y: emu(resolved.top), w: emu(resolved.width), h: emu(resolved.height) }
}

/**
 * Rotation and flips. Taken from `absoluteFrame`, which composes enclosing group rotations
 * and XOR-composes group flips — the same reason position is absolute.
 */
function transformOptions(shape: AnyShape): Record<string, IrValue | undefined> {
	const frame = shape.absoluteFrame
	const rotate = frame ? frame.rotation : (shape.rotation ?? 0)
	return {
		rotate: rotate === 0 ? undefined : rotate,
		flipH: (frame?.flipH ?? shape.flipH) ? true : undefined,
		flipV: (frame?.flipV ?? shape.flipV) ? true : undefined,
	}
}

/**
 * A shape's fill as `ShapeFillProps`.
 *
 * Prefers the raw scheme token over `resolvedFill` for the same reason run colour does: a
 * token keeps tracking the destination theme. `resolvedFill` is the fallback because it is
 * the only accessor that sees a `p:style/a:fillRef` — a shape styled entirely from the
 * theme has no `a:solidFill` of its own, so without it the shape would come out unfilled.
 */
function fillOption(shape: AnyShape, notes: NoteScope, assets: AssetResolver): IrValue | undefined {
	const gradient = shape.gradientFill
	if (gradient) {
		const stops = gradientStops(gradient, notes, 'fill')
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const pattern = shape.patternFill
	if (pattern?.preset) {
		return {
			type: 'pattern',
			pattern: compact({
				preset: pattern.preset,
				fgColor: pattern.foreground ? literalColor(pattern.foreground.effectiveHex) : undefined,
				bgColor: pattern.background ? literalColor(pattern.background.effectiveHex) : undefined,
			}) ?? { preset: pattern.preset },
		}
	}

	// An image-filled *surface* is not a picture object — it is `ShapeFillProps.image`, so
	// the bytes are re-embedded through the same asset resolver an `addImage` uses. Before
	// the fill option itself, since a shape carrying a `a:blipFill` has no `a:solidFill` for
	// the colour legs below to find.
	const picture = shape.pictureFill
	if (picture) return pictureFillOption(picture, assets, notes, SHAPE_PICTURE_FILL)

	const scheme = shape.fillSchemeColor
	if (isWritableSchemeToken(scheme)) return { color: scheme as string }
	if (shape.fillColor !== null) return { color: literalColor(shape.fillColor) }

	const resolved = shape.resolvedFill
	if (!resolved) return undefined
	if (scheme !== null) {
		notes.note(
			'fill.schemeToken',
			'approximated',
			'unwritable',
			`fill scheme colour "${scheme}" is outside the ten tokens the write path maps, so it is baked to a literal hex and stops tracking the theme`
		)
	}
	return compact({
		color: literalColor(resolved.effectiveHex),
		transparency: alphaToTransparency(resolved.alpha),
	})
}

/**
 * A shape's outline as `ShapeLineProps`.
 *
 * This is where the measured silent loss lives. When a shape takes its outline from
 * `p:style/a:lnRef`, `resolvedLine` supplies the colour but `lineWidthPt` and `lineDash`
 * both report `null` — the theme's `a:lnStyleLst` entry, where the width and dash actually
 * live, has no accessor. Emitting the colour and noting the rest is the honest outcome;
 * inventing a width would be worse, because it would look deliberate.
 */
function lineOption(shape: AnyShape, notes: NoteScope): Record<string, IrValue> | undefined {
	if (shape.lineNoFill) return { type: 'none' }

	const gradient = shape.lineGradient
	if (gradient) {
		const stops = gradientStops(gradient, notes, 'line')
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const widthPt = shape.lineWidthPt
	const dash = shape.lineDash
	const scheme = shape.lineSchemeColor
	const resolved = shape.resolvedLine

	let color: string | undefined
	if (isWritableSchemeToken(scheme)) color = scheme as string
	else if (shape.lineColor !== null) color = literalColor(shape.lineColor)
	else if (resolved) color = literalColor(resolved.effectiveHex)

	// A colour only `resolvedLine` could supply means the stroke came from the theme style
	// list, whose width and dash are unreadable.
	if (resolved !== null && shape.lineColor === null && scheme === null && widthPt === null) {
		notes.note(
			'line.width',
			'dropped',
			'unread',
			'this outline comes from the theme style list (p:style/a:lnRef); its width and dash live in the theme fmtScheme a:lnStyleLst, which has no accessor, so only the colour carries'
		)
	}
	if (dash !== null && !WRITABLE_DASHES.has(dash)) {
		notes.note(
			'line.dash',
			'approximated',
			'unwritable',
			`dash style "${dash}" is outside the eight the write API accepts, so the outline falls back to solid`
		)
	}

	const options = compact({
		color,
		width: orUndefined(widthPt),
		dashType: dash !== null && WRITABLE_DASHES.has(dash) ? dash : undefined,
		transparency: alphaToTransparency(resolved?.alpha),
		...arrowOptions(shape, notes),
	})
	return options
}

/** Arrowheads. Types map directly; the `sm`/`med`/`lg` size classes have no option. */
function arrowOptions(shape: AnyShape, notes: NoteScope): Record<string, IrValue | undefined> {
	const ends = shape.lineEnds
	if (!ends) return {}
	const sized = [ends.head, ends.tail].some((end) => end && (end.width !== null || end.length !== null))
	if (sized) {
		notes.note(
			'line.arrowSize',
			'dropped',
			'unwritable',
			'arrowhead width and length classes (@w / @len) have no write-API option, so arrowheads render at the default size'
		)
	}
	return {
		beginArrowType: ends.head && WRITABLE_ARROWS.has(ends.head.type) ? ends.head.type : undefined,
		endArrowType: ends.tail && WRITABLE_ARROWS.has(ends.tail.type) ? ends.tail.type : undefined,
	}
}

/**
 * A read gradient as `GradientFillProps`. Stop positions convert from the read model's
 * 0–1 fraction to the write API's 0–100 percentage; the angle needs no conversion, since
 * both sides use OOXML degrees (clockwise from 3 o'clock).
 */
function gradientStops(gradient: GradientFill, notes: NoteScope, where: 'fill' | 'line'): IrValue | undefined {
	const stops = gradient.stops
		.map((stop) => {
			const color = isWritableSchemeToken(stop.schemeColor)
				? (stop.schemeColor as string)
				: stop.effectiveHex === null
					? null
					: literalColor(stop.effectiveHex)
			if (color === null) return null
			return compact({
				color,
				position: Math.round((stop.position ?? 0) * 100),
				transparency: alphaToTransparency(stop.alpha),
			}) as IrValue
		})
		.filter((stop): stop is IrValue => stop !== null)

	if (stops.length < 2) {
		notes.note(
			`${where}.gradient`,
			'dropped',
			'unsupported',
			'a gradient with fewer than two resolvable stops cannot be expressed, so this falls back to no gradient'
		)
		return undefined
	}

	if (gradient.kind === 'path') {
		// `a:path` covers circle/rect/shape; the write API models only the radial case, so
		// a rectangular or shape-following gradient becomes a circular one.
		if (gradient.path !== null && gradient.path !== 'circle') {
			notes.note(
				`${where}.gradient.path`,
				'approximated',
				'unwritable',
				`a "${gradient.path}" path gradient is emitted as a radial one; the write API models no other path shape`
			)
		}
		return { kind: 'radial', stops }
	}
	return compact({ kind: 'linear', angle: gradient.angleDeg ?? 0, stops })
}

/** Outer shadow as `ShadowProps`; an inner shadow uses the same option with `type: 'inner'`. */
function shadowOption(shape: AnyShape, notes: NoteScope): IrValue | undefined {
	if (shape.reflection || shape.softEdge) {
		notes.note(
			'shape.effects',
			'dropped',
			'unwritable',
			'a:reflection and a:softEdge are read but have no write-API emitter, so those effects are lost'
		)
	}

	const outer = shape.shadow
	const source = outer ?? shape.innerShadow
	if (!source) return undefined

	return compact({
		type: outer ? 'outer' : 'inner',
		color: source.color === null ? undefined : literalColor(source.color),
		blur: source.blurPt,
		offset: source.offsetPt,
		angle: source.angleDeg,
		transparency: alphaToTransparency(source.alpha),
	})
}

/** Glow. The read model reports opacity as a 0–1 alpha, which is the write API's unit too. */
function glowOption(shape: AnyShape): IrValue | undefined {
	const glow = shape.glow
	if (!glow) return undefined
	return compact({
		size: glow.radiusPt,
		color: glow.color === null ? undefined : literalColor(glow.color),
		opacity: glow.alpha,
	})
}

/** The style block every shape kind shares. */
function styleOptions(shape: AnyShape, notes: NoteScope, assets: AssetResolver): Record<string, IrValue | undefined> {
	return {
		fill: fillOption(shape, notes, assets),
		line: lineOption(shape, notes),
		shadow: shadowOption(shape, notes),
		glow: glowOption(shape),
	}
}

/* ===== per-kind mappers ===== */

/**
 * An `AutoShape` becomes an `addShape` when it is bare geometry and an `addText` when it
 * carries text. A shape with both takes the `addText` form, because `addShape` has no text
 * argument while `addText` accepts a `shape` option — so only that form expresses both.
 */
function autoShapeCall(shape: AutoShape, notes: NoteScope, assets: AssetResolver): CallIr | null {
	const preset = shape.presetGeometry
	const custom = shape.customGeometry
	const frame = shape.hasTextFrame ? shape.textFrame : null

	// "Has text" means text this converter can author, which is *runs*. `TextFrame.text` also
	// counts `a:fld` field text — a slide number, date or footer placeholder — and there is no
	// run behind it, so measuring by `text` made every field placeholder emit an `addText`
	// carrying a single empty string: an invisible empty box where a number used to be.
	const runText = frame === null ? '' : frame.paragraphs.flatMap((p) => p.runs.map((run) => run.text)).join('')
	const hasText = runText.length > 0

	// Run text is a subsequence of the frame's text, so anything longer came from a field.
	if (frame !== null && frame.text.replaceAll('\n', '').length > runText.length) {
		notes.note(
			'text.field',
			'dropped',
			'unread',
			'this shape holds an automatic field (a:fld — a slide number, date or footer), which has no accessor and no addText expression, so its text is not reproduced'
		)
	}

	// An OMML equation contributes nothing to `TextFrame.text`, so an equation-only shape
	// looks like an empty box and would otherwise emit as bare geometry.
	if (frame !== null && hasEquation(shape.element_)) {
		notes.note(
			'text.equation',
			'dropped',
			'unread',
			'this shape holds an OMML equation, which no accessor exposes, so the shape is emitted without it — even though TextProps.math (and the ts-pptx/math subpath) could author one'
		)
	}

	if (shape.placeholder) {
		notes.note(
			'shape.placeholder',
			'flattened',
			'unsupported',
			'a placeholder is emitted as a plain shape with its inherited geometry and styling baked in, so it no longer follows its layout'
		)
	}

	const common = {
		...positionOptions(shape, notes),
		...transformOptions(shape),
		...styleOptions(shape, notes, assets),
		objectName: shape.name || undefined,
	}

	if (custom) {
		notes.note(
			'shape.custGeom.guides',
			'dropped',
			'unread',
			'custGeom guides, adjust handles and connection sites (a:gdLst / a:ahLst / a:cxnLst) have no accessor, so only the path outline carries'
		)
		const options = compact({ ...common, points: customGeometryPoints(shape, custom) })
		return { method: 'addShape', args: ['custGeom', options ?? {}], ...nameOf(shape) }
	}

	if (hasText && frame !== null) {
		// `isTextBox` and `shape` are independent, not alternatives: a PowerPoint text box
		// carries `txBox="1"` *and* an explicit `prstGeom rect`, so reading text-box-ness off
		// the absence of geometry misclassified every one of them as an auto shape.
		const options = compact({
			...common,
			...textFrameOptions(frame, notes),
			...(isTextBox(shape.element_) ? { isTextBox: true } : {}),
			...(preset === null ? {} : { shape: preset, ...adjustOptions(shape) }),
		})
		return { method: 'addText', args: [textRuns(frame, notes), options ?? {}], ...nameOf(shape) }
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

/**
 * Preset-geometry adjust handles (`a:avLst`) as `shapeAdjust`.
 *
 * The read model reports each guide's raw formula string; a preset's adjust values are
 * `val <n>` where `<n>` is in 1000ths of a percent, and the write API takes the same handle
 * as a `0.0–1.0` fraction. Anything that is not a plain `val` is a computed guide, which
 * has no write-API expression, so it is skipped rather than mis-scaled.
 */
function adjustOptions(shape: AutoShape): Record<string, IrValue | undefined> {
	const values = shape.adjustValues
	const guides: IrValue[] = []
	for (const name of Object.keys(values).sort()) {
		const match = /^val\s+(-?\d+)$/.exec((values[name] ?? '').trim())
		if (!match?.[1]) continue
		guides.push({ name, value: Number(match[1]) / 100000 })
	}
	return guides.length ? { shapeAdjust: guides } : {}
}

/**
 * `custGeom` path commands as the write API's `points` array.
 *
 * Path coordinates are integers in the path's own `0..w`/`0..h` viewport, not EMU, so each
 * is scaled onto the shape's box before being printed as EMU. A path that declares no
 * viewport (`@w`/`@h` of 0) is already in the shape's own space and passes through.
 */
function customGeometryPoints(shape: AutoShape, custom: CustomGeometry): IrValue {
	const frame = shape.absoluteFrame
	const points: IrValue[] = []

	for (const path of custom.paths) {
		const sx = path.w > 0 && frame ? frame.width / path.w : 1
		const sy = path.h > 0 && frame ? frame.height / path.h : 1
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
function pictureCall(shape: Picture, notes: NoteScope, assets: AssetResolver): CallIr | null {
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

	const crop = shape.crop
	const options = compact({
		...positionOptions(shape, notes),
		...transformOptions(shape),
		objectName: shape.name || undefined,
		// `crop`, not `sizing`. Both exist and they are not interchangeable: `crop` is
		// `a:srcRect` emitted verbatim as percentage edge insets — the exact model the read
		// model reports — while `sizing: 'crop'` cuts a window in *displayed inches* against
		// the image's measured natural size. Feeding fractions to `sizing` read them as
		// inches, which shrank every cropped picture to a fraction of its box.
		crop:
			crop === null ? undefined : { l: crop.left * 100, t: crop.top * 100, r: crop.right * 100, b: crop.bottom * 100 },
	})
	return { method: 'addImage', args: [{ ...(options ?? {}), data: asset }], ...nameOf(shape) }
}

/**
 * `ConnectorProps` keys that {@link lineOption} can produce. A connector styles its stroke
 * with *flat* options rather than the nested `line` object every other shape takes, and it
 * accepts a strict subset: no `transparency`, no gradient stroke, and its own `type` names
 * the routing (`straight`/`elbow`/`curved`), not a fill kind — so spreading a line option
 * object into it wholesale would both drop silently and collide.
 */
const CONNECTOR_LINE_KEYS = ['color', 'width', 'dashType', 'beginArrowType', 'endArrowType'] as const

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
	const frame = shape.absoluteFrame
	if (!frame) return null

	if (shape.startConnection || shape.endConnection) {
		notes.note(
			'connector.binding',
			'dropped',
			'unsupported',
			'connector endpoint bindings reference source shape ids with no counterpart in the output, so the connector lands unbound and no longer follows its shapes'
		)
	}
	if (shape.rotation) {
		notes.note(
			'connector.rotation',
			'dropped',
			'unwritable',
			'addConnector takes two endpoints and no rotation, so a rotated connector lands along the unrotated diagonal of its box'
		)
	}

	const line = lineOption(shape, notes) ?? {}
	const stroke: Record<string, IrValue> = {}
	for (const key of CONNECTOR_LINE_KEYS) {
		const value = line[key]
		if (value !== undefined) stroke[key] = value
	}
	if (line['type'] === 'gradient' || line['transparency'] !== undefined) {
		notes.note(
			'connector.line',
			'flattened',
			'unwritable',
			'a connector styles its stroke with flat colour/width options, which cannot express a gradient stroke or a stroke transparency, so those fall back to a plain line'
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
				objectName: shape.name || undefined,
				...stroke,
			}) ?? {},
		],
		...nameOf(shape),
	}
}

/** A `GraphicFrame` hosts a table or a chart; both have their own mapper. */
function graphicFrameCall(shape: GraphicFrame, notes: NoteScope, assets: AssetResolver): CallIr | null {
	const table = shape.table
	if (table) return tableCall(shape, table, notes, assets)

	const chart = shape.chart
	if (chart) return chartCall(shape, chart, notes)

	if (shape.hasChartEx) {
		// Reached only if the deck-level walk did not already mark the slide `carried`.
		notes.note(
			'chartEx.all',
			'dropped',
			'unwritable',
			'an extended chart (waterfall, funnel, box-and-whisker, …) has a full reader but no write-API counterpart'
		)
		return null
	}

	notes.note(
		'graphicFrame.unknown',
		'dropped',
		'unread',
		'this graphic frame hosts neither a table nor a chart (SmartArt or an OLE object), which the read model does not decode'
	)
	return null
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
 * Charts, tables, media and placeholders are excluded from groups by the write API, so a
 * group holding one loses that child.
 */
function groupCall(shape: GroupShape, notes: NoteScope, assets: AssetResolver): CallIr | null {
	const children: IrValue[] = []
	for (const child of shape.shapes) {
		const call = shapeCall(child, notes, assets)
		if (!call) continue
		const tagged = asGroupChild(call)
		if (!tagged) {
			notes.note(
				'group.child',
				'dropped',
				'unsupported',
				`addGroup accepts no ${call.method.replace(/^add/, '').toLowerCase()} child (charts, tables and media are excluded by design), so this child is omitted from the group`
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
		args: [children, compact({ ...positionOptions(shape, notes), objectName: shape.name || undefined }) ?? {}],
		...nameOf(shape),
	}
}

/**
 * Re-tag a call as a `GroupChildProps` entry, or `null` when the content type is one
 * `addGroup` excludes. The tagged form is a single-key object, not a `{ type, args }` pair.
 */
function asGroupChild(call: CallIr): IrValue | null {
	const [first, second] = call.args
	switch (call.method) {
		case 'addText':
			return { text: compact({ text: first, options: second }) ?? {} }
		case 'addShape':
			return { shape: compact({ type: first, options: second }) ?? {} }
		case 'addImage':
			return { image: first ?? {} }
		case 'addGroup':
			return { group: compact({ children: first, options: second }) ?? {} }
		default:
			// Charts, tables, connectors and media have no GroupChildProps variant.
			return null
	}
}

/** Attach the source name to a call, when the shape had one. */
function nameOf(shape: AnyShape): { sourceName?: string } {
	return shape.name ? { sourceName: shape.name } : {}
}
