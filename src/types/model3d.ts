/**
 * ts-pptx: 3D model types (PowerPoint's Insert ▸ 3D Models — `slide.addModel3d()`).
 *
 * A 3D model is a glTF binary (`.glb`) embedded in the package and drawn live by PowerPoint
 * 2019+. It is emitted as a `<p:graphicFrame>` in the 2017 `am3d` namespace, wrapped in
 * `<mc:AlternateContent>` with a picture fallback — the same structural shape as a zoom.
 *
 * PREVIEW IMAGE: everything that is not PowerPoint 2019+ — including PowerPoint's own
 * `mc:Fallback` path, thumbnails, and PDF export — draws the preview picture, never the model.
 * ts-pptx is Node-first and cannot rasterize a 3D scene, so it emits a neutral gray
 * **placeholder** when `preview` is omitted. Supply one for any deck meant to read correctly
 * outside PowerPoint. This is the same bargain `addOleObject()`'s `cover` makes.
 *
 * CAMERA: PowerPoint derives its camera from the model's bounding box; ts-pptx does not parse
 * glTF, so it emits a fixed default framed for a model roughly 2 units across. See
 * {@link Model3dCameraProps} and `docs/3d-models.md` for the formula and when to override.
 */
import type { DataOrPathRequiredProps, PositionProps } from './core.js'
import type { ObjectNameProps } from './object.js'

/** A point or direction in the model's 3D scene, in metres. */
export interface Model3dPoint {
	x: number
	y: number
	z: number
}

/**
 * Camera override for `addModel3d()`, mapping onto `am3d:camera`.
 *
 * The defaults are what PowerPoint itself wrote for a 2×2×2 cube, and they are correct for any
 * model whose bounding box is roughly cubic **once `meterPerModelUnit` is right** — the camera is
 * expressed in metres, so the scale is what actually decides framing.
 *
 * PowerPoint's own rule, measured against three models: it normalizes the model's largest
 * bounding-box extent to 1 metre and then places the camera at
 *
 * ```
 * pos.z = |halfExtents / maxExtent| / sin(fov / 2)
 * ```
 *
 * looking at the origin. A caller who knows their model's bounding box can reproduce that
 * exactly; a caller who does not should at least set `meterPerModelUnit` to `1 / maxExtent`.
 */
export interface Model3dCameraProps {
	/**
	 * Camera position in metres (`am3d:camera/am3d:pos`).
	 * @default { x: 0, y: 0, z: 2.2630334 }
	 */
	pos?: Model3dPoint
	/**
	 * Point the camera aims at, in metres (`am3d:camera/am3d:lookAt`).
	 * @default { x: 0, y: 0, z: 0 }
	 */
	lookAt?: Model3dPoint
	/**
	 * Camera "up" direction (`am3d:camera/am3d:up`). Need not be normalized.
	 * @default { x: 0, y: 1, z: 0 }
	 */
	up?: Model3dPoint
	/**
	 * Vertical field of view in **degrees** (`am3d:perspective@fov`).
	 * - range: greater than 0 and less than 180
	 * @default 45
	 */
	fov?: number
}

/**
 * Add an embedded 3D model (PowerPoint's Insert ▸ 3D Models) to a slide.
 *
 * The `.glb` bytes travel inside the `.pptx` (in `ppt/media/`), so PowerPoint 2019+ renders and
 * lets the viewer orbit the model. Requires either `data` or `path`. Linked (non-embedded)
 * models are not supported.
 */
interface Model3dBaseProps extends PositionProps, ObjectNameProps {
	/**
	 * Picture of the model shown wherever the live 3D view is unavailable — a raster image `path`
	 * (Node/local) or base64 `data:` URI.
	 * - the library cannot render a 3D scene, so when this is omitted a neutral gray placeholder is
	 *   embedded; PowerPoint 2019+ draws the live model over it, but every other consumer — and
	 *   PowerPoint's own `mc:Fallback` path, thumbnails, and PDF export — shows exactly what is here
	 * - supply a real render whenever the deck is meant to read correctly outside PowerPoint
	 * @example { path: 'assets/engine-render.png' }
	 */
	preview?: { path?: string; data?: string }
	/**
	 * Camera override. Omitted, ts-pptx emits the camera PowerPoint wrote for a 2×2×2 cube.
	 * @see {@link Model3dCameraProps} for the framing formula and when this matters.
	 */
	camera?: Model3dCameraProps
	/**
	 * Metres per model unit (`am3d:trans/am3d:meterPerModelUnit`) — the scale that decides how
	 * large the model is in the scene the {@link camera} looks at.
	 *
	 * This is the single override most models need. PowerPoint sets it to `1 / maxExtent`, where
	 * `maxExtent` is the model's largest bounding-box dimension in model units, which normalizes
	 * every model to 1 metre across. ts-pptx does not parse the `.glb`, so it cannot measure that
	 * — it defaults to `0.5`, correct for a model 2 units across. A model 100 units across left at
	 * the default is a 50-metre object with the camera 2.26 metres from its centre, i.e. the
	 * viewer is inside it.
	 * - range: greater than 0
	 * @default 0.5
	 * @example 1 / 100 // a model whose largest bounding-box dimension is 100 units
	 */
	meterPerModelUnit?: number
	/**
	 * Alt text (`p:cNvPr@descr`) — what a screen reader announces for the model.
	 */
	altText?: string
}
/**
 * Options for `slide.addModel3d()`. Requires either `data` (base64, with or without a
 * `data:...;base64,` header) or `path` (a local/remote `.glb` read at export time).
 *
 * Sizing note: `w`/`h` default to 4 × 3 inches rather than being measured. A 3D model has no
 * intrinsic aspect ratio, and the library never opens the payload.
 */
export type Model3dProps = Model3dBaseProps & DataOrPathRequiredProps
