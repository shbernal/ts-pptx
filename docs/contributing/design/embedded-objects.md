---
doc-schema-version: 1
title: "Embedded objects: design"
summary: "How addOleObject() and addModel3d() write a payload behind a preview picture: the mc:AlternateContent envelope and parts they share, what each object puts in its two branches, and the evidence the markup rests on."
read_when:
  - Changing the slide XML, relationships or parts that addOleObject() or addModel3d() write
  - Adding another object that embeds a payload behind a preview picture
  - Checking an am3d or p:oleObj constant against PowerPoint's own output
doc_type: "architecture"
---

# Embedded objects: design

The user guides are [OLE embedded objects](../../ole-objects.md) and [3D models](../../3d-models.md). This page is
how the code behind them is shaped.

## Where the code lives

| Concern | File |
| --- | --- |
| option checks, payload kind, relationship registration | `src/gen/define/ole.ts`, `src/gen/define/model3d.ts` |
| the payload source check and the frame | `src/gen/define/object-options.ts` |
| the preview picture and its gray placeholder | `src/gen/define/preview-image.ts` |
| slide XML | `src/gen/slide/objects/ole.ts`, `src/gen/slide/objects/model3d.ts` |
| the frame and the fallback picture body | `src/gen/slide/objects/shared.ts` |
| the `mc:AlternateContent` envelope | `src/gen/oxml/alternate-content.ts` |
| relationship types in the slide's `.rels` | `src/gen/slide/object.ts` |
| part names | `src/gen/utils.ts` |
| part de-duplication and zip storage | `src/package/assemble.ts` |

## What both objects share

Each object stores a payload part and a preview picture part, and writes one `mc:AlternateContent` whose
`mc:Fallback` branch draws the preview.

| Piece | Function | What it does |
| --- | --- | --- |
| payload check | `requirePayloadSource()` | throws `ole/missing-source` or `model3d/missing-source` when neither `data` nor `path` is set |
| frame | `framedObjectOptions()` | resolves the frame with `x` and `y` at 0 and `w` and `h` at 4 by 3 inches when unset, and names the object |
| payload relationship | `pushMediaRel()` | registers the caller's `path` or `data`; the bytes load when the deck is written and are stored unchanged |
| preview relationship | `registerPreviewImage()` | registers an ordinary image relationship, de-duplicated on the slide like `addImage` |
| missing or headerless preview | `registerPreviewImage()` | substitutes a 32 by 32 `#E7E6E6` PNG |
| envelope | `alternateContentEl()` | declares the `Requires` prefix on the `mc:Choice` element itself, where a reader resolves it |
| fallback picture body | `previewPicBody()` | writes a stretched `p:blipFill` and a `p:spPr` at the slide-absolute frame |

Three rules hold for both definers:

- Every option check runs before `framedObjectOptions()` names the object and before any relationship is
  registered. A throw leaves no part, relationship or name counter behind.
- Part names come from `nextMediaTarget()`: `{kind}-{slide}-{n}.{ext}`, where `{n}` counts the media
  relationships already on the slide. The payload registers first, so its preview takes the next `{n}`.
- The relationship writer in `src/gen/slide/object.ts` picks a type by looking for `image`, `audio` or `video`
  in the part's content type, and has no final `else`. A payload's content type matches none of them, so the
  writer checks `oleRelType` and `model3dRelType` first. Without those checks the relationship is dropped
  and the slide keeps a dangling `r:id`.

## Where the two objects differ

| | OLE object | 3D model |
| --- | --- | --- |
| outer element | `p:graphicFrame`, with the envelope inside `a:graphicData` | the envelope, with a `p:graphicFrame` inside `mc:Choice` |
| `a:graphicData@uri` | `http://schemas.openxmlformats.org/presentationml/2006/ole` | `http://schemas.microsoft.com/office/drawing/2017/model3d` |
| `mc:Choice@Requires` | `v`, declared as `urn:schemas-microsoft-com:vml` | `am3d`, declared as the model3d namespace |
| `mc:Choice` holds | `p:oleObj` with `p:embed` | `p:graphicFrame` with `am3d:model3d` |
| `mc:Fallback` holds | the same `p:oleObj`, plus a `p:pic` | a `p:pic` |
| preview referenced by | the fallback `p:pic` | `am3d:blip` and the fallback `p:pic` |
| fallback `p:cNvPr` | `id="0" name=""` | the frame's own id, name and alt text |
| fallback `a:picLocks` | none, `p:cNvPicPr` is empty | `FALLBACK_PICTURE_LOCKS` plus `noCrop` |
| frame `a:graphicFrameLocks` | `noChangeAspect`, with `objectLock` spread over it | `objectLock`, or an empty element when it sets no flag |
| payload part | `ppt/embeddings/oleObject-{slide}-{n}.{extn}` | `ppt/media/model3d-{slide}-{n}.glb` |
| payload relationship type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/package`, or `.../relationships/oleObject` for `bin` | `http://schemas.microsoft.com/office/2017/06/relationships/model3d` |
| payload content type | a `Default` for the Office extension, or `application/vnd.openxmlformats-officedocument.oleObject` for `bin` | `<Default Extension="glb" ContentType="model/gltf.binary"/>` |
| deck-wide payload de-duplication | never | by extension plus data string |
| zip storage of the payload | STORE for the six Office extensions, DEFLATE for `bin` | DEFLATE |

`src/package/assemble.ts` collapses identical media parts across the deck, keyed on extension plus the data
string. It skips any relationship with `oleRelType`, because two objects sharing one embedding would make
editing either one rewrite the other's source. A model is read-only geometry, so it takes the collapse. The
key is the string, not the decoded bytes: the same model passed once as bare base64 and once with a `data:`
header lands in two parts.

The six Office payloads are ZIP archives themselves, so `ZIP_CONTAINER_EXTN` stores them without a second
DEFLATE pass.

## The OLE object

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="2" name="Object 1" descr=""/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">
          <p:oleObj name="Worksheet" r:id="rId1" imgW="3657600" imgH="1828800" progId="Excel.Sheet.12">
            <p:embed/>
          </p:oleObj>
        </mc:Choice>
        <mc:Fallback>
          <p:oleObj name="Worksheet" r:id="rId1" imgW="3657600" imgH="1828800" progId="Excel.Sheet.12">
            <p:embed/>
            <p:pic><!-- nvPicPr with cNvPr id="0", blipFill on rId2, spPr at the frame --></p:pic>
          </p:oleObj>
        </mc:Fallback>
      </mc:AlternateContent>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

- `p:oleObj` attributes go in PowerPoint's order: `name`, `showAsIcon`, `r:id`, `imgW`, `imgH`, `progId`.
  `showAsIcon` is written only when true.
- `name` is the kind's label from `OLE_FORMATS`: `Worksheet`, `Document`, `Presentation`, or `Object` for a
  `bin` payload.
- `resolvePartExtn()` settles the kind from `extn`, then a `data:` URI's MIME type, then the `path` extension,
  then `progId`. The reverse maps `EXTN_BY_CONTENT_TYPE` and `EXTN_BY_PROG_ID` are built from `OLE_FORMATS`,
  so a seventh kind is one entry there.
- `Requires="v"` declares the VML namespace, and no VML follows. PowerPoint writes no `spid` attribute and no
  `vmlDrawing` part for an embedded object, so neither is emitted.
- The fallback picture takes `id="0"` and an empty name. It is another rendering of the frame, not a sibling
  shape, so it takes no id from the slide's shape-id space.
- `oleImageSize()` bounds `imgW` and `imgH` to `ST_PositiveCoordinate32`, 0 to 2147483647, and rounds a
  fraction.
- `bin` is `BIN_FORMAT`. The library writes the caller's bytes into it and builds no compound file.

## The 3D model

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" Requires="am3d">
    <p:graphicFrame>
      <!-- p:nvGraphicFramePr, then p:xfrm at the slide-absolute frame -->
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">
          <am3d:model3d r:embed="rId1">
            <am3d:spPr><!-- a:xfrm with a:off 0,0 and the frame's a:ext, then prstGeom rect --></am3d:spPr>
            <am3d:camera>
              <am3d:pos x="0" y="0" z="81469202"/>
              <am3d:up dx="0" dy="36000000" dz="0"/>
              <am3d:lookAt x="0" y="0" z="0"/>
              <am3d:perspective fov="2700000"/>
            </am3d:camera>
            <am3d:trans>
              <am3d:meterPerModelUnit n="500000" d="1000000"/>
              <!-- identity preTrans, scale, rot and postTrans -->
            </am3d:trans>
            <am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="rId2"/></am3d:raster>
            <am3d:objViewport viewportSz="3338805"/>
            <am3d:ambientLight><!-- fixed --></am3d:ambientLight>
            <am3d:ptLight rad="0"><!-- fixed, three of them --></am3d:ptLight>
          </am3d:model3d>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </mc:Choice>
  <mc:Fallback>
    <p:pic><!-- the preview picture on rId2, at the slide-absolute frame --></p:pic>
  </mc:Fallback>
</mc:AlternateContent>
```

- `am3d:model3d` children go in this order: `spPr`, `camera`, `trans`, `raster`, `objViewport`,
  `ambientLight`, three `ptLight`.
- `am3d:spPr`'s `a:xfrm` is frame-local: offset 0,0, extent equal to the frame's. The slide position lives on
  `p:xfrm` and on the fallback picture.
- `resolveCamera()` converts the caller's units once, so the emitter does no arithmetic. Points are fixed
  point over 36,000,000 (`up@dy="36000000"` is the unit vector), `fov` is in 60,000ths of a degree like every
  DrawingML angle, and every rational is `n` over 1,000,000. `meterPerModelUnit` is rounded to a whole `n`,
  which is six decimal places, so a positive value below 0.0000005 is written as `n="0"`.
- The relationship type sits under `2017/06`, which is not the `2017` of the namespace. The content type is
  `model/gltf.binary` with a dot, not `model/gltf-binary`.
- These constants in `src/gen/slide/objects/model3d.ts` are copied from PowerPoint's output, not derived:
  `RASTER`, `OBJ_VIEWPORT_SZ`, the identity `am3d:trans` children, `AMBIENT` and `PT_LIGHTS`. PowerPoint
  wrote the same light rig for every model probed. `viewportSz` depends on the bounding box in PowerPoint's
  output, and the cube's value is written.
- `DEFAULT_CAMERA` and `DEFAULT_METER_PER_MODEL_UNIT` are PowerPoint's values for a 2 by 2 by 2 cube.
  PowerPoint scales the largest bounding-box extent to 1 metre and sets
  `pos.z = |halfExtents| / maxExtent / sin(fov / 2)` looking at the origin. That framing is scale-invariant
  and shape-dependent: a 20-unit cube got the cube's camera with `meterPerModelUnit` 1/20, and a 1 by 1 by 8
  box got `pos.z` 1.3268209. The probe data is in `test/read/fixtures/README.md`. The library does not parse
  glTF, so it writes the cube's values and leaves the rest to the caller.
- When PowerPoint inserts a model it re-exports the `.glb` through its own glTF exporter. The library stores
  the caller's bytes, and PowerPoint draws them.
- The fallback picture reuses the frame's id and name, and its `a:picLocks` set is fixed. `objectLock` lands on
  `a:graphicFrameLocks`, which accepts a different flag set.

## Evidence

- `test/read/fixtures/model3d.pptx` is a deck PowerPoint authored with `Shapes.Add3DModel`, from
  `test/read/fixtures/authoring/author-model3d.ps1`. `test/read/model3d-roundtrip.test.js` compares the
  emitted `am3d:model3d` subtree with it byte for byte, apart from relationship ids and the frame extent. The
  same file pins the relationship graph, a byte-identical load and save, and `importSlide` carrying the model
  and both relationships.
- The OLE package choices were checked against a deck PowerPoint authored with `Shapes.AddOLEObject`. That
  deck is not committed.
- `test/schema-cases.js` builds both objects and asserts parts, relationship types, content types,
  attributes and the de-duplication counts. Schema validation reaches only the `mc:Fallback` branch, so it
  says nothing about `p:oleObj` or `am3d:model3d` inside `mc:Choice`. See
  [What the validator cannot see](../testing.md#what-the-validator-cannot-see).
- Option refusals and warnings are pinned in `test/regression/api/definition-reality-checks.test.js`,
  `test/regression/api/non-finite-numbers.test.js`, `test/regression/image/image-source-resolution.test.js`
  and `test/regression/shape/authored-frame.test.js`.
- `pnpm run test:com` opens both in PowerPoint. The `ole` leg reads each `progId` back, and the `model3d` leg
  reads the camera back and exports the slide to PNG. See
  [Check rendering with pixels, not COM properties](../testing.md#check-rendering-with-pixels-not-com-properties).
- No showcase deck builds an OLE object, so the byte-identity harness passes an OLE refactor without looking at
  it. See [What the demos verify](../testing.md#what-the-demos-verify).
