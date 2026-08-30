/**
 * ts-pptx: 3D model (`am3d:model3d`) slide-object serialization
 *
 * Emits a `model3d` slide object as an `<mc:AlternateContent>`. The `mc:Choice` carries the real
 * `<p:graphicFrame>` in the 2017 `am3d` namespace, rendered live by PowerPoint 2019+; the
 * `mc:Fallback` carries a plain preview picture, which is what every other consumer — older
 * PowerPoint, thumbnails, PDF export — actually draws. See `gen/define/model3d.ts`.
 *
 * Every constant below is transcribed from a PowerPoint-authored deck
 * (`test/read/fixtures/model3d.pptx`, `Shapes.Add3DModel`), including the lighting rig, which came
 * out byte-identical for every probe model and is therefore emitted as a fixed studio setup rather
 * than derived from anything.
 */

import type { Model3dInternal } from '../../../types/internal.js'
import { el, raw, voidEl } from '../../oxml/el.js'
import {
	FALLBACK_PICTURE_LOCKS,
	MC_NS,
	type RenderContext,
	cNvPrOpen,
	graphicFrameEl,
	previewPicBody,
	xfrmEl,
} from './shared.js'
import { GRAPHIC_FRAME_LOCK_ATTRS, PICTURE_LOCK_ATTRS, genXmlObjectLock } from '../../drawingml/locks.js'

/** The `am3d` namespace, doubling as `a:graphicData@uri` and the `mc:Choice Requires` token's URI. */
const AM3D_NS = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
/** Denominator of every `am3d` rational element (`@n`/`@d`). */
const RATIO_DEN = 1000000
/**
 * `am3d:raster@rName`/`@rVer` — the renderer that produced the cached preview. PowerPoint stamps
 * its own renderer here; ts-pptx did not rasterize anything, but the attributes are what PowerPoint
 * writes and the element is where the preview blip lives, so they are reproduced verbatim.
 */
const RASTER = { rName: 'Office3DRenderer', rVer: '16.0.8326' }
/**
 * `am3d:objViewport@viewportSz`. Bounding-box derived in PowerPoint's output (a cube and a slab
 * gave different values); the cube's is emitted, matching the fixed default camera.
 */
const OBJ_VIEWPORT_SZ = 3338805

/** One `am3d:ptLight` of the fixed studio rig: colour (scRGB, 1000ths of a percent), intensity, position. */
interface PtLight {
	clr: { r: number; g: number; b: number }
	intensityN: number
	pos: { x: number; y: number; z: number }
}
/** The three point lights PowerPoint writes for every inserted model — key, fill and rim. */
const PT_LIGHTS: readonly PtLight[] = [
	{ clr: { r: 100000, g: 75000, b: 50000 }, intensityN: 9765625, pos: { x: 21959998, y: 70920001, z: 16344003 } },
	{ clr: { r: 40000, g: 60000, b: 95000 }, intensityN: 12250000, pos: { x: -37964106, y: 51130435, z: 57631972 } },
	{ clr: { r: 86837, g: 72700, b: 100000 }, intensityN: 3125000, pos: { x: -37739122, y: 58056624, z: -34769649 } },
]
/** `am3d:ambientLight` — neutral 50% scRGB at half illuminance, constant like the point lights. */
const AMBIENT = { clr: { r: 50000, g: 50000, b: 50000 }, illuminanceN: 500000 }

/** `<am3d:clr><a:scrgbClr …/></am3d:clr>` — the colour wrapper both light kinds share. */
function lightColor(clr: { r: number; g: number; b: number }): string {
	return el('am3d:clr', null, raw(voidEl('a:scrgbClr', clr)))
}

/** The `am3d:camera` block: eye position, up vector, aim point and perspective fov. */
function cameraEl(model: Model3dInternal): string {
	return el('am3d:camera', null, [
		raw(voidEl('am3d:pos', model.pos)),
		raw(voidEl('am3d:up', model.up)),
		raw(voidEl('am3d:lookAt', model.lookAt)),
		raw(voidEl('am3d:perspective', { fov: model.fov })),
	])
}

/**
 * The `am3d:trans` block: model-unit scale plus an identity placement. Only `meterPerModelUnit`
 * varies — `preTrans`/`scale`/`rot`/`postTrans` are the identity PowerPoint writes for a model
 * inserted without any subsequent 3D manipulation, and `am3d:rot` is emitted attribute-less.
 */
function transEl(model: Model3dInternal): string {
	const unit = { n: RATIO_DEN, d: RATIO_DEN }
	return el('am3d:trans', null, [
		raw(voidEl('am3d:meterPerModelUnit', { n: model.meterPerModelUnitN, d: RATIO_DEN })),
		raw(voidEl('am3d:preTrans', { dx: 0, dy: 0, dz: 0 })),
		raw(
			el('am3d:scale', null, [raw(voidEl('am3d:sx', unit)), raw(voidEl('am3d:sy', unit)), raw(voidEl('am3d:sz', unit))])
		),
		raw(voidEl('am3d:rot')),
		raw(voidEl('am3d:postTrans', { dx: 0, dy: 0, dz: 0 })),
	])
}

/** The fixed lighting rig: one ambient light and three point lights, in PowerPoint's order. */
function lightsXml(): string {
	return (
		el('am3d:ambientLight', null, [
			raw(lightColor(AMBIENT.clr)),
			raw(voidEl('am3d:illuminance', { n: AMBIENT.illuminanceN, d: RATIO_DEN })),
		]) +
		PT_LIGHTS.map((light) =>
			el('am3d:ptLight', { rad: 0 }, [
				raw(lightColor(light.clr)),
				raw(voidEl('am3d:intensity', { n: light.intensityN, d: RATIO_DEN })),
				raw(voidEl('am3d:pos', light.pos)),
			])
		).join('')
	)
}

/** The `mc:Fallback` picture — the cached preview, drawn by everything that is not PowerPoint 2019+. */
function fallbackPic(
	picId: number,
	objectName: string | undefined,
	altText: string,
	model: Model3dInternal,
	xf: { x: number; y: number; cx: number; cy: number }
): string {
	return el('p:pic', null, [
		raw(
			el('p:nvPicPr', null, [
				raw(cNvPrOpen(picId, objectName, altText) + '/>'),
				raw(
					el(
						'p:cNvPicPr',
						null,
						raw(
							// PowerPoint's model-3d lock set: the zoom/picture set plus `noCrop` — a 3D model is
							// reframed by its camera, never by cropping the cached raster. Fixed even though
							// `Model3dProps` has an `objectLock`, for the reason the zoom emitter spells out:
							// this is the `mc:Fallback` picture, and `a:picLocks` takes a different flag set
							// from the `a:graphicFrameLocks` the caller's locks land on.
							genXmlObjectLock('a:picLocks', PICTURE_LOCK_ATTRS, { ...FALLBACK_PICTURE_LOCKS, noCrop: true })
						)
					)
				),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(previewPicBody(model.previewRid, xf)),
	])
}

/**
 * Render a `model3d` slide object to its `<mc:AlternateContent>` XML. The `mc:Choice` carries the
 * real `<p:graphicFrame>` (the `am3d:model3d` graphicData PowerPoint 2019+ renders live); the
 * `mc:Fallback` carries the cached preview picture. See `gen/define/model3d.ts`.
 */
export function renderModel3dObject(ctx: RenderContext): string {
	const {
		obj: slideItemObj,
		idx,
		frame: { x, y, cx, cy },
	} = ctx
	const model = slideItemObj.model3d
	if (!model) return ''
	const opts = slideItemObj.options || {}
	const altText = opts.altText || ''

	// `am3d:spPr`'s xfrm is FRAME-LOCAL: origin 0,0 with the graphic frame's own extent. The
	// slide-absolute position lives on `p:xfrm` below (and on the fallback picture's `p:spPr`).
	const spPr = el('am3d:spPr', null, [
		raw(xfrmEl('a:xfrm', { x: 0, y: 0, cx, cy })),
		raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
	])

	const model3d = el('am3d:model3d', { 'r:embed': `rId${model.modelRid}` }, [
		raw(spPr),
		raw(cameraEl(model)),
		raw(transEl(model)),
		// The raster's blip and the fallback picture share one image rel, as PowerPoint writes it.
		raw(el('am3d:raster', RASTER, raw(voidEl('am3d:blip', { 'r:embed': `rId${model.previewRid}` })))),
		raw(voidEl('am3d:objViewport', { viewportSz: OBJ_VIEWPORT_SZ })),
		raw(lightsXml()),
	])

	const nvGraphicFramePr = el('p:nvGraphicFramePr', null, [
		raw(cNvPrOpen(idx + 2, opts.objectName, altText) + '/>'),
		raw(
			el(
				'p:cNvGraphicFramePr',
				null,
				// PowerPoint writes an empty `<a:graphicFrameLocks/>` here — no flags, but the
				// element present — so that is what an unlocked model emits. `genXmlObjectLock`
				// returns `''` when nothing is set, hence the fallback.
				raw(
					genXmlObjectLock('a:graphicFrameLocks', GRAPHIC_FRAME_LOCK_ATTRS, opts.objectLock, opts.objectName) ||
						voidEl('a:graphicFrameLocks')
				)
			)
		),
		raw(voidEl('p:nvPr')),
	])
	const graphicFrame = graphicFrameEl({
		nvGraphicFramePr,
		frame: { x, y, cx, cy },
		uri: AM3D_NS,
		payload: model3d,
	})

	return el('mc:AlternateContent', { 'xmlns:mc': MC_NS }, [
		raw(el('mc:Choice', { 'xmlns:am3d': AM3D_NS, Requires: 'am3d' }, raw(graphicFrame))),
		raw(el('mc:Fallback', null, raw(fallbackPic(idx + 2, opts.objectName, altText, model, { x, y, cx, cy })))),
	])
}
