/**
 * ts-pptx: 3D model Definition (`addModel3d()` — PowerPoint's Insert ▸ 3D Models).
 *
 * Registers the `.glb` payload as an embedded media part plus a preview-picture image rel, and
 * pushes a `SlideObject{ _type: model3d, model3d }` for `gen/slide/objects/model3d.ts` to emit as
 * an `<mc:AlternateContent>` wrapping a `<p:graphicFrame>`.
 *
 * Ground truth for the package-level choices here (verified against a deck authored by PowerPoint
 * via `Shapes.Add3DModel` — `test/read/fixtures/model3d.pptx`):
 *   - the payload is referenced with the Microsoft `…/office/2017/06/relationships/model3d` rel
 *     type and registered by a content-type `Default` for `glb` — `model/gltf.binary`, spelled
 *     with a dot, not the `model/gltf-binary` of the IANA-style media type;
 *   - both the payload and the preview picture live in `ppt/media/`, the payload named `model3dN`;
 *   - the preview picture is an ordinary image rel, referenced twice — once by `am3d:raster`'s
 *     `am3d:blip` and once by the `mc:Fallback` picture.
 *
 * Camera: PowerPoint derives its camera from the model's bounding box (it normalizes the largest
 * extent to 1 metre, then frames the resulting bounding sphere). This library does not parse
 * glTF, so it emits the camera PowerPoint wrote for a 2×2×2 cube and lets the caller override the
 * scale — see `docs/3d-models.md`.
 */
import { SlideObjectType } from '../../enums.js'
import type { Model3dProps, Model3dPoint, Model3dCameraProps } from '../../types/model3d.js'
import type { PresSlideInternal, SlideObject, Model3dInternal } from '../../types/internal.js'
import { getNewRelId, nextMediaTarget } from '../utils.js'
import { resolveObjectName } from './object-name.js'
import { registerPreviewImage } from './preview-image.js'
import { InvalidOptionError } from '../../errors.js'
import { warn } from '../../diagnostics.js'
import { ANGLE_UNITS_PER_DEGREE } from '../../units.js'

/** OPC content type for a glTF binary, as PowerPoint spells it in `[Content_Types].xml`. */
const GLB_CONTENT_TYPE = 'model/gltf.binary'
/** `.rels` `Type` URI for the embedded 3D model part. Note `2017/06` — not the namespace's `2017`. */
const MODEL3D_REL_TYPE = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d'

/** Fixed-point denominator for `am3d` linear values: `am3d:up@dy="36000000"` is the unit vector. */
const AM3D_UNIT = 36000000
/** Denominator of every `am3d` rational (`meterPerModelUnit`, `scale`, `illuminance`, `intensity`). */
const AM3D_RATIO_DEN = 1000000

/**
 * The camera PowerPoint wrote for the 2×2×2 cube fixture, in caller units (metres / degrees).
 * `z` is `sqrt(3)/2 / sin(22.5°)` — the bounding sphere of a 1-metre cube framed at a 45° fov.
 */
const DEFAULT_CAMERA = {
	pos: { x: 0, y: 0, z: 2.2630334 },
	lookAt: { x: 0, y: 0, z: 0 },
	up: { x: 0, y: 1, z: 0 },
	fov: 45,
} as const
/** `am3d:meterPerModelUnit` for the cube fixture: 1 / (its 2-unit largest extent). */
const DEFAULT_METER_PER_MODEL_UNIT = 0.5

/** Reject a non-finite camera component rather than letting `NaN` reach the XML as `"NaN"`. */
function toAm3dUnits(point: Model3dPoint | undefined, fallback: Model3dPoint, label: string): Model3dPoint {
	const src = point ?? fallback
	for (const axis of ['x', 'y', 'z'] as const) {
		if (!Number.isFinite(src[axis])) {
			throw new InvalidOptionError(
				'model3d/invalid-camera',
				`addModel3d(): camera.${label}.${axis} must be a finite number, got ${String(src[axis])}.`
			)
		}
	}
	return { x: Math.round(src.x * AM3D_UNIT), y: Math.round(src.y * AM3D_UNIT), z: Math.round(src.z * AM3D_UNIT) }
}

/** Resolve `camera` + `meterPerModelUnit` into the wire units {@link Model3dInternal} carries. */
function resolveCamera(
	camera: Model3dCameraProps | undefined,
	meterPerModelUnit: number | undefined
): Omit<Model3dInternal, 'modelRid' | 'previewRid'> {
	const fovDeg = camera?.fov ?? DEFAULT_CAMERA.fov
	if (!Number.isFinite(fovDeg) || fovDeg <= 0 || fovDeg >= 180) {
		throw new InvalidOptionError(
			'model3d/invalid-fov',
			`addModel3d(): camera.fov must be greater than 0 and less than 180 degrees, got ${String(fovDeg)}.`
		)
	}
	const scale = meterPerModelUnit ?? DEFAULT_METER_PER_MODEL_UNIT
	if (!Number.isFinite(scale) || scale <= 0) {
		throw new InvalidOptionError(
			'model3d/invalid-scale',
			`addModel3d(): meterPerModelUnit must be a finite number greater than 0, got ${String(scale)}.`
		)
	}
	const up = toAm3dUnits(camera?.up, DEFAULT_CAMERA.up, 'up')
	return {
		pos: toAm3dUnits(camera?.pos, DEFAULT_CAMERA.pos, 'pos'),
		lookAt: toAm3dUnits(camera?.lookAt, DEFAULT_CAMERA.lookAt, 'lookAt'),
		up: { dx: up.x, dy: up.y, dz: up.z },
		fov: Math.round(fovDeg * ANGLE_UNITS_PER_DEGREE),
		meterPerModelUnitN: Math.round(scale * AM3D_RATIO_DEN),
	}
}

/**
 * Adds an embedded 3D model to a slide definition.
 * @param {PresSlideInternal} target - slide the model will be added to
 * @param {Model3dProps} opt - 3D model options
 */
export function addModel3dDefinition(target: PresSlideInternal, opt: Model3dProps): void {
	const strData = opt.data || ''
	const strPath = opt.path || ''

	// STEP 1: REALITY-CHECK. The payload is the whole point; there is no meaningful default.
	if (!strPath && !strData) {
		throw new InvalidOptionError('model3d/missing-source', 'addModel3d(): either `data` or `path` are required!')
	}

	// STEP 2: Resolve the camera before touching the rel table, so an invalid option throws without
	// leaving a half-registered part behind.
	const camera = resolveCamera(opt.camera, opt.meterPerModelUnit)
	const objectName = resolveObjectName(target, SlideObjectType.model3d, {
		label: '3D Model',
		kind: 'model3d',
		supplied: opt.objectName,
	})

	// STEP 3: Register the `.glb` payload rel. Every model gets its own rel and its own camera; the
	// package-part de-dup in `package/assemble.ts` still collapses byte-identical payloads to one
	// part, which is wanted here (a model is read-only geometry, unlike an OLE payload, which is
	// exempted there because double-clicking one would rewrite the other's source).
	const modelRid = getNewRelId(target)
	target._relsMedia.push({
		path: strPath || 'preencoded.glb',
		type: GLB_CONTENT_TYPE,
		extn: 'glb',
		data: strData,
		rId: modelRid,
		model3dRelType: MODEL3D_REL_TYPE,
		Target: nextMediaTarget(target, 'model3d', 'glb'),
	})

	// STEP 4: Register the preview picture. Unlike OLE's silent placeholder, say so out loud — a
	// missing preview is invisible in PowerPoint (which draws the live model over it) and shows up
	// only in thumbnails, PDF export, and older viewers, i.e. long after the deck was checked.
	if (!opt.preview?.path && !opt.preview?.data) {
		warn(
			'model3d/preview-missing',
			'addModel3d(): no `preview` image supplied, so a gray placeholder is embedded. PowerPoint 2019+ ' +
				'draws the live model over it, but thumbnails, PDF export and older viewers show the placeholder.'
		)
	}
	const previewRid = registerPreviewImage(target, opt.preview)

	// LAST: Push the slide object for the `<mc:AlternateContent>` emitter.
	const slideData: SlideObject = {
		_type: SlideObjectType.model3d,
		options: {
			x: opt.x ?? 0,
			y: opt.y ?? 0,
			// A 3D model has no aspect ratio and the library never opens the payload, so — as with
			// `addOleObject` — there is no natural size to measure.
			w: opt.w ?? 4,
			h: opt.h ?? 3,
			objectName,
			...(opt.altText ? { altText: opt.altText } : {}),
			...(opt.objectLock ? { objectLock: opt.objectLock } : {}),
		},
		model3d: { modelRid, previewRid, ...camera },
	}
	target._slideObjects.push(slideData)
}
