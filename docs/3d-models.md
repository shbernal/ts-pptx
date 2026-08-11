---
doc-schema-version: 1
title: "3D Models"
summary: "Embed a glTF binary (.glb) into a slide with addModel3d(), so PowerPoint 2019+ renders it live and everything else falls back to a preview picture."
read_when:
  - Embedding a 3D model so PowerPoint renders and orbits it
  - Reproducing PowerPoint's Insert > 3D Models
  - Working out why an embedded model renders far too large, or from the wrong angle
  - Choosing the preview picture a slide shows outside PowerPoint 2019+
doc_type: "guide"
---

# 3D Models

`slide.addModel3d()` embeds a glTF binary inside the `.pptx` and places it on the
slide as a live 3D model: PowerPoint's **Insert ▸ 3D Models**. In PowerPoint
2019 or later the viewer can click and drag to orbit it.

```js
const s = pptx.addSlide()
s.addModel3d({
	path: 'assets/engine.glb',
	preview: { path: 'assets/engine-render.png' },
	meterPerModelUnit: 1 / 240, // the model's largest bounding-box dimension, in model units
	x: 1,
	y: 1,
	w: 6,
	h: 4,
})
```

The model's bytes travel inside the package, so the deck stays self-contained.
Either `data` (base64, with or without a `data:...;base64,` header) or `path` is
required; everything else is optional.

> Linked models (a model referenced from outside the package) are not
> supported. Only embedded payloads. Neither are animation scenes: a `.glb` with
> animation clips embeds fine, but ts-pptx emits no `Model3DFormat` animation
> settings, so PowerPoint shows the model at rest.

Only `.glb` (the binary container) is embedded. A `.gltf` + separate `.bin` +
loose textures is a *set* of files, and PowerPoint stores a single part; convert
to `.glb` first.

## The preview picture (`preview`)

**Almost everything that looks at your deck will see this picture, not the
model.** PowerPoint 2019+ reads the `mc:Choice` branch and renders the live
model; every other consumer reads the `mc:Fallback` branch, and so do
PowerPoint's own slide thumbnails, PDF export, and print.

This library is Node-first: it has no 3D renderer, so it cannot produce one.
Supply your own with `preview`, taking a `path` or base64 `data` just like
`addImage`:

```js
s.addModel3d({ path: 'engine.glb', preview: { path: 'engine-render.png' } })
```

Omit it and a neutral gray placeholder is embedded, and the library emits a
[`model3d/preview-missing`](diagnostics.md) warning, because the gap is
invisible exactly where you would check for it (on screen, in PowerPoint) and
shows up later, in the thumbnail or the PDF someone else opens.

The most convenient way to produce one is to let PowerPoint render it: insert the
model once by hand, export the slide as a picture, and crop. That is how
`demos/common/images/cube_3d_preview.png` was made.

## Scale: `meterPerModelUnit`

**This is the setting most models need, and the one most likely to make a model
look wrong if left alone.**

The `am3d` scene is measured in metres, so PowerPoint has to know how big a model
unit is. When *PowerPoint* inserts a model it reads the bounding box out of the
file and normalizes the largest dimension to 1 metre: a 240-unit-wide model gets
`meterPerModelUnit = 1/240`.

ts-pptx never opens the `.glb` (there is no glTF parser here, deliberately), so it
cannot measure that. It emits `0.5`, which is correct for a model 2 units across
and wrong for everything else. Left at the default, a model 240 units across
becomes a 120-metre object with the camera 2.26 metres from its centre: the
viewer is inside it, and the slide shows a wall of shading.

So: set it to `1 / <largest bounding-box dimension, in model units>`. Most
exporters report that dimension, and it is also the span of `accessors[].min/max`
on the `POSITION` accessors in the file's JSON chunk.

```js
s.addModel3d({ path: 'engine.glb', meterPerModelUnit: 1 / 240 })
```

With the scale right, the default camera frames a roughly-cubic model correctly,
because the camera is expressed in metres and the model is now 1 metre across.

## Camera

`camera` overrides the viewpoint. The defaults are what PowerPoint wrote for a
2×2×2 cube:

| Field | Default | Meaning |
| --- | --- | --- |
| `pos` | `{ x: 0, y: 0, z: 2.2630334 }` | eye position, in metres |
| `lookAt` | `{ x: 0, y: 0, z: 0 }` | the point it aims at, in metres |
| `up` | `{ x: 0, y: 1, z: 0 }` | up direction (need not be normalized) |
| `fov` | `45` | vertical field of view, in **degrees** |

```js
// Orbit to 35° azimuth / 25° elevation at 2.6 m, looking at the origin.
s.addModel3d({
	path: 'cube.glb',
	camera: { pos: { x: 1.3516, y: 1.0988, z: 1.9305 }, lookAt: { x: 0, y: 0, z: 0 }, fov: 45 },
})
```

If you know your model's bounding box you can reproduce PowerPoint's own framing
exactly. Measured against three models, it places the camera at

```
pos.z = |halfExtents / maxExtent| / sin(fov / 2)
```

with `lookAt` at the origin: i.e. far enough back to contain the bounding sphere
of the scaled model. For the default 45° fov and a cube that is
`(√3 / 2) / sin(22.5°) = 2.2630334`, which is where the default comes from.

Values outside their valid ranges are rejected rather than coerced: a non-finite
camera component, an `fov` outside `0 < fov < 180`, or a `meterPerModelUnit` that
is not greater than zero each throw an
[`InvalidOptionError`](errors.md).

### What is *not* configurable

The lighting is a fixed studio rig: one ambient light and three point lights,
transcribed from PowerPoint's output, which emitted the identical rig for every
model tested. PowerPoint's own UI exposes lighting presets; ts-pptx does not, and
a model authored here always gets the default rig.

## Sizing

`w`/`h` default to **4 × 3 inches**. A 3D model has no intrinsic aspect ratio and
the library never opens the payload, so there is nothing to measure: set them
explicitly. The model is drawn inside the frame at its camera's framing; a cube
in a 4 × 3 frame renders as a 3 × 3 square, centred.

## Emitted OOXML

Each model is an `<mc:AlternateContent>` whose `mc:Choice` holds a
`<p:graphicFrame>` in the 2017 `am3d` namespace, and whose `mc:Fallback` holds an
ordinary `<p:pic>` on the preview image:

```xml
<mc:Choice xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" Requires="am3d">
  <p:graphicFrame>
    <!-- nvGraphicFramePr, p:xfrm (slide-absolute) -->
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">
      <am3d:model3d r:embed="rId1">
        <am3d:spPr><!-- frame-LOCAL xfrm: off 0,0 --></am3d:spPr>
        <am3d:camera>…</am3d:camera>
        <am3d:trans><am3d:meterPerModelUnit n="500000" d="1000000"/>…</am3d:trans>
        <am3d:raster rName="Office3DRenderer" rVer="16.0.8326">
          <am3d:blip r:embed="rId2"/>
        </am3d:raster>
        <am3d:objViewport viewportSz="3338805"/>
        <am3d:ambientLight>…</am3d:ambientLight>
        <am3d:ptLight rad="0">…</am3d:ptLight><!-- x3 -->
      </am3d:model3d>
    </a:graphicData></a:graphic>
  </p:graphicFrame>
</mc:Choice>
<mc:Fallback><p:pic><!-- the preview picture, slide-absolute --></p:pic></mc:Fallback>
```

Two relationships, one part each:

| Rel type | Target | Content type |
| --- | --- | --- |
| `…/office/2017/06/relationships/model3d` | `ppt/media/model3d-…​.glb` | `<Default Extension="glb" ContentType="model/gltf.binary"/>` |
| `…/officeDocument/2006/relationships/image` | `ppt/media/image-…​.png` | the usual image `Default` |

Details worth knowing, all transcribed from a PowerPoint-authored deck
(`test/read/fixtures/model3d.pptx`, produced by `Shapes.Add3DModel`):

- The rel type is under `2017/06`, which is **not** the `2017` of the namespace.
- The content type is `model/gltf.binary`, spelled with a dot: not the
  IANA-style `model/gltf-binary`.
- `am3d:raster`'s blip and the `mc:Fallback` picture share the *same* image
  relationship; the preview is stored once.
- `am3d:spPr`'s `a:xfrm` is frame-local (offset 0,0, extent equal to the frame's).
  The slide-absolute position lives on `p:xfrm` and on the fallback picture.
- Linear `am3d` values are fixed point over 36,000,000 (`up@dy="36000000"` is the
  unit vector); `fov` is in 60000ths of a degree, like every other DrawingML angle.
- PowerPoint re-exports the `.glb` through its own glTF exporter when *it* inserts
  a model. ts-pptx embeds your bytes unchanged; PowerPoint reads them fine.

## Reading a deck that contains a model

There is no typed read accessor. A model read through `ts-pptx/read` surfaces as
an inert `graphicFrame` shape carrying its `objectName`: visible to anything
enumerating shapes, and preserved byte-intact through load → save and through
`importSlide`, but with no camera or payload accessor. That is deliberate for
now; see `test/read/model3d-roundtrip.test.js`, which pins the survival.

## A note on testing

The Open XML SDK validator does **not** descend into an `mc:Choice`: it checks
only the `mc:Fallback` branch. Measured at `Microsoft365`: a bogus attribute, an
unknown child element, elements out of document order, a non-numeric `fov`, and
even a missing required `r:embed` inside `am3d:model3d` all validate clean, while
the same mutation inside `mc:Fallback` is caught. This is a property of
`mc:AlternateContent`, so it applies equally to the zoom and OLE emitters.

Schema validation is therefore not evidence about this construct. What covers it
is a byte-for-byte comparison of the emitted `am3d:model3d` body against the
PowerPoint-authored fixture, plus `pnpm run test:com`, which opens a generated
deck in the real application, reads the shape back as `msoShape3DModel`, **and**
exports the slide to PNG to confirm the model actually rasterized. The last step
is not redundant: with a deliberately corrupted `.glb`, PowerPoint still reports a
3D model with the right camera, and renders nothing at all.
