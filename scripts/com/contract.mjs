/**
 * What the generated decks, the VBScripts and the verifiers all have to agree on.
 *
 * The three used to sit in one 848-line file, so the agreement was implicit -- a shape name in a
 * deck builder, the same string quoted inside a VBScript template three hundred lines down, and a
 * third copy in the verifier that reads the output back. Splitting the file makes the agreement
 * an import, which is the point of splitting it.
 *
 * Its own module rather than a corner of `decks.mjs`, because half of these are facts about what
 * PowerPoint must report -- a `PpActionType` number, a `ProgID`, a camera position -- which the
 * deck builders never see.
 */

// PpActionType enum values PowerPoint resolves `ppaction://hlinkshowjump?jump=<x>` into.
// The whole point of the read-back: prove the jump became a live navigation action, not a
// dead ppActionHyperlink (7) / ppActionNone (0).
export const EXPECTED_ACTION = {
	nextslide: 1, // ppActionNextSlide
	previousslide: 2, // ppActionPreviousSlide
	firstslide: 3, // ppActionFirstSlide
	lastslide: 4, // ppActionLastSlide
	lastslideviewed: 5, // ppActionLastSlideViewed
	endshow: 6, // ppActionEndShow
}

// The custGeom deck binds the connector's start to connection site index 1 (0-based, OOXML
// `<a:stCxn idx="1">`) of the shape named `freeform`. PowerPoint numbers connection sites
// 1-based, so the resolved `BeginConnectionSite` is idx+1 = 2.
export const GEOM_TARGET_NAME = 'freeform'
export const GEOM_START_IDX = 1
export const GEOM_EXPECTED_SITE = GEOM_START_IDX + 1

// The OLE deck's objects, by `objectName` → the ProgID PowerPoint must resolve each back to.
// A schema-valid `<p:oleObj>` that PowerPoint quietly discards enumerates as *no shape at all*,
// so the read-back is the only thing proving the embedded object survived the round trip.
export const EXPECTED_OLE_PROGID = {
	OlePackage: 'Excel.Sheet.12', // embedded OPC package → `.../package` rel
	OleBlob: 'Package', // generic OLE blob → `.../oleObject` rel, `.bin` part
}
/** An empty (but structurally valid) ZIP — enough of an "xlsx" for PowerPoint to bind the OLE server. */
export const EMPTY_ZIP_B64 = 'UEsFBgAAAAAAAAAAAAAAAAAAAAAAAA=='

// The 3D-model deck. `msoShape3DModel` = 30 — a shape PowerPoint bound as a live model rather
// than leaving as an inert graphicFrame or collapsing to the fallback picture. The camera is the
// library's default, in metres, and PowerPoint exposes it as `Model3DFormat.CameraPositionZ`; a
// reading that far off means the `am3d:camera` subtree did not parse.
export const MODEL3D_SHAPE_NAME = 'Cube3D'
export const MODEL3D_SHAPE_TYPE = 30
export const MODEL3D_EXPECTED_CAMERA_Z = 2.2630334
/**
 * The model deck's preview picture is solid magenta, which no lit gray model can produce. Type 30
 * plus a good camera still only proves PowerPoint resolved the *model*, not that it drew one — an
 * unreadable payload renders as the fallback or as nothing. So the deck is exported to PNG and the
 * pixels are read: any magenta means we are looking at the fallback, and an all-white frame means
 * nothing was drawn at all. Both absent ⇒ the `.glb` really rasterized.
 */
export const MODEL3D_PREVIEW_RGB = [255, 0, 255]
/** Export size, and the frame's slide-relative rect (inches) — see `generateModel3dDeck`. */
export const MODEL3D_EXPORT = { w: 960, h: 540 }
export const MODEL3D_FRAME_IN = { x: 2, y: 1, w: 4, h: 3 }
/**
 * The model deck's slide size, in inches.
 *
 * The deck builder defines the layout and the verifier scales the frame rect into exported
 * pixels with it. Those two were 145 lines apart in one file and are now in two, so the size is
 * stated here rather than written out twice: a layout change that only one of them followed
 * would move the sampled rectangle off the model and read the slide background as "not drawn".
 */
export const MODEL3D_LAYOUT_IN = { w: 10, h: 5.625 }

/**
 * The preset-geometry deck. Each row is one slide: a shape drawn at {@link PRSTGEOM_FRAME_IN}
 * with one adjustment guide driven to `adj`, in the raw 1/100000 units `<a:gd fmla="val N">`
 * carries. `same` names the slide whose exported pixels this one must match.
 *
 * The claim under test is that an out-of-range guide is *inert*, not corrupting: the preset's own
 * formula pins it, so PowerPoint paints the same shape it paints at the bound. That is the whole
 * reason this library emits a finite adjustment verbatim instead of clamping it against a
 * per-preset range table it would have to invent -- see `gen/drawingml/geometry.ts`.
 *
 * The third row of each pair is the sensitivity check and is load-bearing: it is an IN-range value,
 * so it must paint *differently*. Without it a run that exported six identical blank slides would
 * pass every equality assertion it makes.
 */
export const PRSTGEOM_CASES = [
	{ label: 'roundRect@max', shape: 'roundRect', adj: { adj: 50000 } },
	{ label: 'roundRect@over', shape: 'roundRect', adj: { adj: 266667 }, same: 'roundRect@max' },
	{ label: 'roundRect@half', shape: 'roundRect', adj: { adj: 25000 }, differs: 'roundRect@max' },
	{ label: 'blockArc@max', shape: 'blockArc', adj: { adj1: 0, adj2: 10800000, adj3: 50000 } },
	{ label: 'blockArc@over', shape: 'blockArc', adj: { adj1: 0, adj2: 10800000, adj3: 5000000 }, same: 'blockArc@max' },
	{ label: 'blockArc@half', shape: 'blockArc', adj: { adj1: 0, adj2: 10800000, adj3: 25000 }, differs: 'blockArc@max' },
]
/** Export size, the shape's rect (inches), and the slide size the deck is built at. */
export const PRSTGEOM_EXPORT = { w: 960, h: 540 }
export const PRSTGEOM_FRAME_IN = { x: 1, y: 1, w: 4, h: 3 }
export const PRSTGEOM_LAYOUT_IN = { w: 10, h: 5.625 }
