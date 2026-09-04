/**
 * The VBScripts that drive desktop PowerPoint, one per deck.
 *
 * Every builder returns a complete `.vbs` source string: {@link vbsOpenHeader} opens the deck and
 * prints `OPEN_OK` or `OPEN_ERR`, the middle reads the construct under test back out over COM, and
 * {@link vbsFooter} closes and quits. The entry point writes the string to a temp file, runs it
 * under `cscript`, and hands the stdout lines to a verifier.
 *
 * String building, not string escaping: every value spliced in here is a repo constant or a path
 * this script made, never caller input.
 */

import { MODEL3D_EXPORT, PRSTGEOM_CASES, PRSTGEOM_EXPORT } from './contract.mjs'

// --- 3. the VBScripts that drive PowerPoint ---------------------------------
/**
 * Shared header: create the app, open the deck, emit OPEN_OK/OPEN_ERR.
 *
 * Default is headless + read-only. Pass `withWindow` for checks that read OLE objects back:
 * a windowless PowerPoint does not instantiate embedded objects, so `Shapes` comes back without
 * them — indistinguishable from PowerPoint having discarded them. (A deck PowerPoint authored
 * itself enumerates zero shapes under a headless open too, which is how that was pinned down.)
 */
/**
 * @param {string} pptxFile
 * @param {boolean} [withWindow] open with a window, for the features a headless open will not instantiate
 * @returns {string}
 */
export function vbsOpenHeader(pptxFile, withWindow = false) {
	// WithWindow:=msoFalse (0) keeps it headless; ReadOnly avoids touching the file.
	const openArgs = withWindow ? '0, 0, -1' : '-1, 0, 0'
	return `Option Explicit
Dim ppt, pres, sld, shp
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "NO_POWERPOINT\t" & Err.Description
  WScript.Quit 3
End If
Err.Clear
Set pres = ppt.Presentations.Open("${pptxFile.replace(/\\/g, '\\\\')}", ${openArgs})
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "OPEN_ERR\t" & Hex(Err.Number) & "\t" & Err.Description
  ppt.Quit
  WScript.Quit 2
End If
WScript.StdOut.WriteLine "OPEN_OK\t" & pres.Slides.Count
`
}

export function vbsFooter() {
	return `pres.Close
ppt.Quit
WScript.StdOut.WriteLine "DONE"
WScript.Quit 0
`
}

/** @param {string} pptxFile @returns {string} */
export function buildNavVbs(pptxFile) {
	// Emits tab-separated `ACTION` lines (slideIdx, objectName, resolvedActionNum).
	return (
		vbsOpenHeader(pptxFile) +
		`Dim act
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    Set act = shp.ActionSettings(1) ' ppMouseClick = 1
    If Not act Is Nothing Then
      If act.Action <> 0 Then
        WScript.StdOut.WriteLine "ACTION\t" & sld.SlideIndex & "\t" & shp.Name & "\t" & act.Action
      End If
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} pptxFile @returns {string} */
export function buildGeomVbs(pptxFile) {
	// Emits one tab-separated `CONN` line per connector shape:
	//   CONN <name> <beginConnected(-1/0)> <beginConnectedShapeName> <beginConnectionSite>
	return (
		vbsOpenHeader(pptxFile) +
		`Dim cf, tgt
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    If shp.Connector Then
      Set cf = shp.ConnectorFormat
      tgt = ""
      If cf.BeginConnected Then tgt = cf.BeginConnectedShape.Name
      WScript.StdOut.WriteLine "CONN\t" & shp.Name & "\t" & cf.BeginConnected & "\t" & tgt & "\t" & cf.BeginConnectionSite
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} pptxFile @returns {string} */
export function buildOleVbs(pptxFile) {
	// Emits one tab-separated `OLE` line per embedded-object shape: name, resolved ProgID.
	// msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10.
	return (
		vbsOpenHeader(pptxFile, true) +
		`Dim i, j
For i = 1 To pres.Slides.Count
  Set sld = pres.Slides(i)
  For j = 1 To sld.Shapes.Count
    Set shp = sld.Shapes(j)
    If shp.Type = 7 Or shp.Type = 10 Then
      WScript.StdOut.WriteLine "OLE\t" & shp.Name & "\t" & shp.OLEFormat.ProgID
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} pptxFile @returns {string} */
export function buildModel3dVbs(pptxFile) {
	// Emits `M3D <name> <Shape.Type> <Model3D.CameraPositionZ>` per shape, then exports slide 1 to
	// PNG so the caller can prove the model actually rasterized. A window is needed here for the
	// same reason as OLE: a headless PowerPoint does not instantiate the 3D renderer.
	const png = pptxFile.replace(/\.pptx$/i, '.png').replace(/\\/g, '\\\\')
	return (
		vbsOpenHeader(pptxFile, true) +
		`Dim j, camZ
Set sld = pres.Slides(1)
For j = 1 To sld.Shapes.Count
  Set shp = sld.Shapes(j)
  camZ = ""
  On Error Resume Next
  camZ = shp.Model3D.CameraPositionZ
  On Error Goto 0
  WScript.StdOut.WriteLine "M3D" & vbTab & shp.Name & vbTab & shp.Type & vbTab & camZ
Next
Err.Clear
sld.Export "${png}", "PNG", ${MODEL3D_EXPORT.w}, ${MODEL3D_EXPORT.h}
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "EXPORT_ERR" & vbTab & Hex(Err.Number) & vbTab & Err.Description
Else
  WScript.StdOut.WriteLine "EXPORT" & vbTab & "${png}"
End If
` +
		vbsFooter()
	)
}

/** @param {string} pptxFile @returns {string} */
export function buildPresetGeomVbs(pptxFile) {
	// Exports every slide to `<deck>-<n>.png` and emits one `PNG <n> <path>` line each. Nothing is
	// read back over COM on purpose: `Shape.Adjustments` reports the *stored* guide, out-of-range
	// value and all, so it cannot answer what PowerPoint paints. Only the pixels can.
	const base = pptxFile.replace(/\.pptx$/i, '').replace(/\\/g, '\\\\')
	return (
		vbsOpenHeader(pptxFile) +
		`Dim i, png
For i = 1 To pres.Slides.Count
  png = "${base}-" & i & ".png"
  Err.Clear
  On Error Resume Next
  pres.Slides(i).Export png, "PNG", ${PRSTGEOM_EXPORT.w}, ${PRSTGEOM_EXPORT.h}
  If Err.Number <> 0 Then
    WScript.StdOut.WriteLine "EXPORT_ERR" & vbTab & i & vbTab & Hex(Err.Number) & vbTab & Err.Description
  Else
    WScript.StdOut.WriteLine "PNG" & vbTab & i & vbTab & png
  End If
  On Error Goto 0
Next
WScript.StdOut.WriteLine "SLIDES" & vbTab & ${PRSTGEOM_CASES.length}
` +
		vbsFooter()
	)
}
